package main

import (
	"bufio"
	"encoding/json"
	"fmt"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"sync"
	"time"
)

type Record struct {
	Employee  string    `json:"employee"`
	Count     int       `json:"count"`
	Timestamp time.Time `json:"timestamp"`
}

type EmployeeStat struct {
	Employee            string  `json:"employee"`
	CurrentHourCount    int     `json:"currentHourCount"`
	CurrentHourRate     float64 `json:"currentHourRate"`
	AverageHourlyRate   float64 `json:"averageHourlyRate"`
	TotalCount          int     `json:"totalCount"`
	ActiveHoursRecorded int     `json:"activeHoursRecorded"`
}

type StatsResponse struct {
	ServerTime  time.Time      `json:"serverTime"`
	Leaderboard []EmployeeStat `json:"leaderboard"`
}

type Store struct {
	mu      sync.RWMutex
	records []Record
	clients map[chan struct{}]struct{}
	file    *os.File
}

func NewStore(path string) (*Store, error) {
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		return nil, err
	}
	f, err := os.OpenFile(path, os.O_CREATE|os.O_RDWR|os.O_APPEND, 0o644)
	if err != nil {
		return nil, err
	}

	s := &Store{
		records: make([]Record, 0),
		clients: make(map[chan struct{}]struct{}),
		file:    f,
	}

	if err := s.loadFromFile(); err != nil {
		_ = f.Close()
		return nil, err
	}
	return s, nil
}

func (s *Store) loadFromFile() error {
	if _, err := s.file.Seek(0, 0); err != nil {
		return err
	}
	scanner := bufio.NewScanner(s.file)
	for scanner.Scan() {
		line := strings.TrimSpace(scanner.Text())
		if line == "" {
			continue
		}
		var r Record
		if err := json.Unmarshal([]byte(line), &r); err != nil {
			log.Printf("skip invalid line: %v", err)
			continue
		}
		s.records = append(s.records, r)
	}
	if err := scanner.Err(); err != nil {
		return err
	}
	_, err := s.file.Seek(0, 2)
	return err
}

func (s *Store) Add(r Record) error {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.records = append(s.records, r)
	b, err := json.Marshal(r)
	if err != nil {
		return err
	}
	if _, err := s.file.Write(append(b, '\n')); err != nil {
		return err
	}
	if err := s.file.Sync(); err != nil {
		return err
	}
	for ch := range s.clients {
		select {
		case ch <- struct{}{}:
		default:
		}
	}
	return nil
}

func (s *Store) Stats(now time.Time) StatsResponse {
	s.mu.RLock()
	defer s.mu.RUnlock()

	type agg struct {
		total      int
		hourTotals map[string]int
	}

	byEmployee := map[string]*agg{}
	currentHourKey := now.Format("2006-01-02 15")

	for _, r := range s.records {
		employee := strings.TrimSpace(r.Employee)
		if employee == "" {
			continue
		}
		a, ok := byEmployee[employee]
		if !ok {
			a = &agg{hourTotals: map[string]int{}}
			byEmployee[employee] = a
		}
		a.total += r.Count
		hourKey := r.Timestamp.In(now.Location()).Format("2006-01-02 15")
		a.hourTotals[hourKey] += r.Count
	}

	stats := make([]EmployeeStat, 0, len(byEmployee))
	for employee, a := range byEmployee {
		activeHours := len(a.hourTotals)
		if activeHours == 0 {
			activeHours = 1
		}
		currentCount := a.hourTotals[currentHourKey]
		avg := float64(a.total) / float64(activeHours)
		stats = append(stats, EmployeeStat{
			Employee:            employee,
			CurrentHourCount:    currentCount,
			CurrentHourRate:     float64(currentCount),
			AverageHourlyRate:   avg,
			TotalCount:          a.total,
			ActiveHoursRecorded: len(a.hourTotals),
		})
	}

	sort.Slice(stats, func(i, j int) bool {
		if stats[i].CurrentHourRate == stats[j].CurrentHourRate {
			return stats[i].AverageHourlyRate > stats[j].AverageHourlyRate
		}
		return stats[i].CurrentHourRate > stats[j].CurrentHourRate
	})

	return StatsResponse{ServerTime: now, Leaderboard: stats}
}

func (s *Store) registerClient() chan struct{} {
	s.mu.Lock()
	defer s.mu.Unlock()
	ch := make(chan struct{}, 1)
	s.clients[ch] = struct{}{}
	return ch
}

func (s *Store) unregisterClient(ch chan struct{}) {
	s.mu.Lock()
	defer s.mu.Unlock()
	delete(s.clients, ch)
	close(ch)
}

func main() {
	store, err := NewStore("data/records.jsonl")
	if err != nil {
		log.Fatalf("init store failed: %v", err)
	}
	defer store.file.Close()

	mux := http.NewServeMux()

	mux.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/" {
			http.NotFound(w, r)
			return
		}
		w.Header().Set("Content-Type", "text/html; charset=utf-8")
		_, _ = w.Write([]byte(indexHTML))
	})

	mux.HandleFunc("/api/records", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
			return
		}
		var req struct {
			Employee string `json:"employee"`
			Count    int    `json:"count"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, "invalid json", http.StatusBadRequest)
			return
		}
		req.Employee = strings.TrimSpace(req.Employee)
		if req.Employee == "" || req.Count <= 0 {
			http.Error(w, "employee and count must be valid", http.StatusBadRequest)
			return
		}
		rec := Record{Employee: req.Employee, Count: req.Count, Timestamp: time.Now().UTC()}
		if err := store.Add(rec); err != nil {
			http.Error(w, "failed to save record", http.StatusInternalServerError)
			return
		}
		w.Header().Set("Content-Type", "application/json")
		_ = json.NewEncoder(w).Encode(map[string]any{"ok": true})
	})

	mux.HandleFunc("/api/stats", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodGet {
			http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
			return
		}
		resp := store.Stats(time.Now().UTC())
		w.Header().Set("Content-Type", "application/json")
		_ = json.NewEncoder(w).Encode(resp)
	})

	mux.HandleFunc("/api/events", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodGet {
			http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
			return
		}
		w.Header().Set("Content-Type", "text/event-stream")
		w.Header().Set("Cache-Control", "no-cache")
		w.Header().Set("Connection", "keep-alive")

		flusher, ok := w.(http.Flusher)
		if !ok {
			http.Error(w, "stream unsupported", http.StatusInternalServerError)
			return
		}

		client := store.registerClient()
		defer store.unregisterClient(client)

		send := func() {
			resp := store.Stats(time.Now().UTC())
			b, _ := json.Marshal(resp)
			fmt.Fprintf(w, "data: %s\n\n", b)
			flusher.Flush()
		}

		send()
		ticker := time.NewTicker(30 * time.Second)
		defer ticker.Stop()

		for {
			select {
			case <-r.Context().Done():
				return
			case <-client:
				send()
			case <-ticker.C:
				send()
			}
		}
	})

	addr := ":8080"
	log.Printf("employee tracker listening on http://localhost%s", addr)
	if err := http.ListenAndServe(addr, mux); err != nil {
		log.Fatal(err)
	}
}

const indexHTML = `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>员工采集速率看板</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 24px; background: #f5f7fb; }
    .wrap { display: grid; grid-template-columns: 320px 1fr; gap: 24px; }
    .card { background: #fff; border-radius: 10px; padding: 16px; box-shadow: 0 3px 12px rgba(0,0,0,.06); }
    h1,h2 { margin: 0 0 12px; }
    label { display:block; margin-top:8px; font-size:14px; }
    input,button { width:100%; padding:10px; margin-top:6px; box-sizing:border-box; }
    button { background:#1864ff; color:#fff; border:none; cursor:pointer; border-radius:6px; }
    button.secondary { background:#5c677d; margin-top:8px; }
    table { width:100%; border-collapse: collapse; }
    th,td { padding:10px; border-bottom:1px solid #eee; text-align:left; }
    th { background:#fafafa; }
    .muted { color:#666; font-size: 12px; }
    .ok { color: green; }
  </style>
</head>
<body>
  <h1>员工采集实时看板</h1>
  <p class="muted">支持在线部署，也支持在无本地开发环境的电脑上直接运行打包好的可执行文件。</p>
  <div class="wrap">
    <div class="card">
      <h2>员工录入</h2>
      <label>员工姓名</label>
      <input id="employee" placeholder="例如：张三" />
      <label>本次采集条数</label>
      <input id="count" type="number" min="1" placeholder="例如：35" />
      <button id="submit">提交记录</button>
      <p id="msg" class="muted"></p>

      <hr style="margin:14px 0;border:none;border-top:1px solid #eee;" />
      <h2 style="font-size:18px;">定时提醒填写</h2>
      <label>提醒间隔（分钟）</label>
      <input id="remindMinutes" type="number" min="1" value="60" />
      <button id="startReminder" class="secondary">开启提醒</button>
      <button id="stopReminder" class="secondary">关闭提醒</button>
      <p id="reminderMsg" class="muted">默认关闭，开启后会按设定时间弹窗提醒。</p>
    </div>
    <div class="card">
      <h2>每小时速率排行榜</h2>
      <p class="muted" id="time"></p>
      <table>
        <thead>
          <tr>
            <th>排名</th>
            <th>员工</th>
            <th>本小时条数</th>
            <th>平均小时速率</th>
            <th>总条数</th>
            <th>活跃小时数</th>
          </tr>
        </thead>
        <tbody id="rows"></tbody>
      </table>
    </div>
  </div>

  <script>
    const rows = document.getElementById('rows');
    const timeEl = document.getElementById('time');
    const msg = document.getElementById('msg');
    const reminderMsg = document.getElementById('reminderMsg');
    let reminderTimer = null;

    function notifyReminder() {
      const text = '请及时填写本小时采集条数。';
      alert(text);
      if ('Notification' in window && Notification.permission === 'granted') {
        new Notification('采集填报提醒', { body: text });
      }
    }

    async function ensureNotificationPermission() {
      if (!('Notification' in window)) return;
      if (Notification.permission === 'default') {
        try { await Notification.requestPermission(); } catch (e) { console.error(e); }
      }
    }

    async function loadStats() {
      const res = await fetch('/api/stats');
      const data = await res.json();
      render(data);
    }

    function render(data) {
      const list = data.leaderboard || [];
      timeEl.textContent = '服务器时间：' + new Date(data.serverTime).toLocaleString();
      rows.innerHTML = list.map((x, idx) =>
        '<tr>' +
          '<td>' + (idx + 1) + '</td>' +
          '<td>' + x.employee + '</td>' +
          '<td>' + x.currentHourCount + '</td>' +
          '<td>' + x.averageHourlyRate.toFixed(2) + '</td>' +
          '<td>' + x.totalCount + '</td>' +
          '<td>' + x.activeHoursRecorded + '</td>' +
        '</tr>'
      ).join('') || '<tr><td colspan="6">暂无数据</td></tr>';
    }

    document.getElementById('submit').onclick = async () => {
      const employee = document.getElementById('employee').value.trim();
      const count = Number(document.getElementById('count').value);
      msg.className = 'muted';
      if (!employee || count <= 0) {
        msg.textContent = '请填写有效的员工姓名和条数';
        return;
      }
      const res = await fetch('/api/records', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ employee, count })
      });
      if (!res.ok) {
        msg.textContent = '提交失败，请重试';
        return;
      }
      msg.textContent = '提交成功';
      msg.className = 'ok';
      document.getElementById('count').value = '';
      await loadStats();
    };

    document.getElementById('startReminder').onclick = async () => {
      const minutes = Number(document.getElementById('remindMinutes').value);
      if (!minutes || minutes <= 0) {
        reminderMsg.textContent = '提醒间隔必须大于 0 分钟';
        reminderMsg.className = 'muted';
        return;
      }

      await ensureNotificationPermission();
      if (reminderTimer) clearInterval(reminderTimer);
      reminderTimer = setInterval(notifyReminder, minutes * 60 * 1000);
      reminderMsg.textContent = '提醒已开启：每 ' + minutes + ' 分钟提醒一次';
      reminderMsg.className = 'ok';
    };

    document.getElementById('stopReminder').onclick = () => {
      if (reminderTimer) clearInterval(reminderTimer);
      reminderTimer = null;
      reminderMsg.textContent = '提醒已关闭';
      reminderMsg.className = 'muted';
    };

    const es = new EventSource('/api/events');
    es.onmessage = (evt) => {
      try { render(JSON.parse(evt.data)); } catch (e) { console.error(e); }
    };
    es.onerror = () => { loadStats(); };

    loadStats();
  </script>
</body>
</html>`
