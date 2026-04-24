package main

import (
	"bufio"
	"crypto/rand"
	"encoding/csv"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
)

const (
	sessionCookieName = "tracker_session"
	adminPassword     = "Adm!n-7Qx9P2Lm"
)

var defaultEmployees = []string{
	"张杰", "余锐芝", "唐丹", "叶深岚", "杨依然",
	"周小远", "庄锐", "李秀媚", "何乐怡",
	"陈近珊", "钟雪", "黄小米",
	"黄佳魏", "陈君伟", "邱殊艺",
	"甘喜圆", "罗子龙",
}

type Record struct {
	Employee  string    `json:"employee"`
	Count     int       `json:"count"`
	Timestamp time.Time `json:"timestamp"`
}

type ReminderConfig struct {
	Mode        string   `json:"mode"`
	CustomTimes []string `json:"customTimes"`
}

type Settings struct {
	Employees []string       `json:"employees"`
	OpenTime  string         `json:"openTime"`
	CloseTime string         `json:"closeTime"`
	Reminder  ReminderConfig `json:"reminder"`
}

type Session struct {
	Role      string
	Employee  string
	ExpiresAt time.Time
}

type LeaderboardItem struct {
	Rank              int     `json:"rank"`
	Employee          string  `json:"employee"`
	AverageHourlyRate float64 `json:"averageHourlyRate"`
	TotalCount        int     `json:"totalCount"`
}

type EmployeeRecordView struct {
	Employee  string    `json:"employee"`
	Count     int       `json:"count"`
	DayTotal  int       `json:"dayTotal"`
	Timestamp time.Time `json:"timestamp"`
}

type EmployeeDashboardResponse struct {
	ServerTime    time.Time            `json:"serverTime"`
	Employee      string               `json:"employee"`
	CanSubmit     bool                 `json:"canSubmit"`
	SubmitMessage string               `json:"submitMessage"`
	OpenTime      string               `json:"openTime"`
	CloseTime     string               `json:"closeTime"`
	Reminder      ReminderConfig       `json:"reminder"`
	LatestRecord  time.Time            `json:"latestRecord"`
	TodayRecords  []EmployeeRecordView `json:"todayRecords"`
	Leaderboard   []LeaderboardItem    `json:"leaderboard"`
}

type AdminDashboardResponse struct {
	ServerTime   time.Time         `json:"serverTime"`
	Employees    []string          `json:"employees"`
	OpenTime     string            `json:"openTime"`
	CloseTime    string            `json:"closeTime"`
	Reminder     ReminderConfig    `json:"reminder"`
	LatestRecord time.Time         `json:"latestRecord"`
	Leaderboard  []LeaderboardItem `json:"leaderboard"`
}

type Store struct {
	mu         sync.RWMutex
	records    []Record
	sessions   map[string]Session
	clients    map[chan struct{}]struct{}
	recordFile *os.File
	configPath string
	settings   Settings
}

func NewStore(recordsPath, configPath string) (*Store, error) {
	if err := os.MkdirAll(filepath.Dir(recordsPath), 0o755); err != nil {
		return nil, err
	}
	if err := os.MkdirAll(filepath.Dir(configPath), 0o755); err != nil {
		return nil, err
	}
	f, err := os.OpenFile(recordsPath, os.O_CREATE|os.O_RDWR|os.O_APPEND, 0o644)
	if err != nil {
		return nil, err
	}
	s := &Store{
		records:    make([]Record, 0),
		sessions:   map[string]Session{},
		clients:    map[chan struct{}]struct{}{},
		recordFile: f,
		configPath: configPath,
		settings: Settings{
			Employees: append([]string{}, defaultEmployees...),
			OpenTime:  "09:00",
			CloseTime: "18:00",
			Reminder: ReminderConfig{
				Mode:        "hourly",
				CustomTimes: []string{},
			},
		},
	}
	if err := s.loadRecords(); err != nil {
		_ = f.Close()
		return nil, err
	}
	if err := s.loadConfig(); err != nil {
		_ = f.Close()
		return nil, err
	}
	return s, nil
}

func (s *Store) loadRecords() error {
	if _, err := s.recordFile.Seek(0, 0); err != nil {
		return err
	}
	scanner := bufio.NewScanner(s.recordFile)
	for scanner.Scan() {
		line := strings.TrimSpace(scanner.Text())
		if line == "" {
			continue
		}
		var r Record
		if err := json.Unmarshal([]byte(line), &r); err != nil {
			log.Printf("skip invalid record: %v", err)
			continue
		}
		s.records = append(s.records, r)
	}
	if err := scanner.Err(); err != nil {
		return err
	}
	_, err := s.recordFile.Seek(0, 2)
	return err
}

func (s *Store) loadConfig() error {
	b, err := os.ReadFile(s.configPath)
	if err != nil {
		if os.IsNotExist(err) {
			return s.saveConfigLocked()
		}
		return err
	}
	var cfg Settings
	if err := json.Unmarshal(b, &cfg); err != nil {
		return err
	}
	if len(cfg.Employees) == 0 {
		cfg.Employees = append([]string{}, defaultEmployees...)
	}
	if cfg.OpenTime == "" {
		cfg.OpenTime = "09:00"
	}
	if cfg.CloseTime == "" {
		cfg.CloseTime = "18:00"
	}
	if cfg.Reminder.Mode == "" {
		cfg.Reminder.Mode = "hourly"
	}
	s.settings = cfg
	return nil
}

func (s *Store) saveConfigLocked() error {
	b, err := json.MarshalIndent(s.settings, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(s.configPath, b, 0o644)
}

func randomToken() (string, error) {
	b := make([]byte, 16)
	if _, err := rand.Read(b); err != nil {
		return "", err
	}
	return hex.EncodeToString(b), nil
}

func (s *Store) createSession(role, employee string) (string, error) {
	token, err := randomToken()
	if err != nil {
		return "", err
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	s.sessions[token] = Session{Role: role, Employee: employee, ExpiresAt: time.Now().Add(24 * time.Hour)}
	return token, nil
}

func (s *Store) getSession(token string) (Session, bool) {
	s.mu.RLock()
	defer s.mu.RUnlock()
	sess, ok := s.sessions[token]
	if !ok || sess.ExpiresAt.Before(time.Now()) {
		return Session{}, false
	}
	return sess, true
}

func (s *Store) deleteSession(token string) {
	s.mu.Lock()
	defer s.mu.Unlock()
	delete(s.sessions, token)
}

func (s *Store) settingsCopy() Settings {
	s.mu.RLock()
	defer s.mu.RUnlock()
	cp := s.settings
	cp.Employees = append([]string{}, s.settings.Employees...)
	cp.Reminder.CustomTimes = append([]string{}, s.settings.Reminder.CustomTimes...)
	return cp
}

func (s *Store) employeeExists(name string) bool {
	s.mu.RLock()
	defer s.mu.RUnlock()
	for _, e := range s.settings.Employees {
		if e == name {
			return true
		}
	}
	return false
}

func (s *Store) AddRecord(r Record) error {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.records = append(s.records, r)
	b, err := json.Marshal(r)
	if err != nil {
		return err
	}
	if _, err := s.recordFile.Write(append(b, '\n')); err != nil {
		return err
	}
	if err := s.recordFile.Sync(); err != nil {
		return err
	}
	s.broadcastLocked()
	return nil
}

func (s *Store) setSettings(openTime, closeTime, mode string, custom []string) error {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.settings.OpenTime = openTime
	s.settings.CloseTime = closeTime
	s.settings.Reminder.Mode = mode
	s.settings.Reminder.CustomTimes = append([]string{}, custom...)
	if err := s.saveConfigLocked(); err != nil {
		return err
	}
	s.broadcastLocked()
	return nil
}

func (s *Store) addEmployee(name string) error {
	s.mu.Lock()
	defer s.mu.Unlock()
	for _, n := range s.settings.Employees {
		if n == name {
			return fmt.Errorf("员工已存在")
		}
	}
	s.settings.Employees = append(s.settings.Employees, name)
	sort.Strings(s.settings.Employees)
	if err := s.saveConfigLocked(); err != nil {
		return err
	}
	s.broadcastLocked()
	return nil
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

func (s *Store) broadcastLocked() {
	for ch := range s.clients {
		select {
		case ch <- struct{}{}:
		default:
		}
	}
}

func validateHHMM(v string) bool {
	_, err := time.Parse("15:04", v)
	return err == nil
}

func inOpenRange(now time.Time, open, close string) bool {
	o, err := time.Parse("15:04", open)
	if err != nil {
		return true
	}
	c, err := time.Parse("15:04", close)
	if err != nil {
		return true
	}
	start := time.Date(now.Year(), now.Month(), now.Day(), o.Hour(), o.Minute(), 0, 0, now.Location())
	end := time.Date(now.Year(), now.Month(), now.Day(), c.Hour(), c.Minute(), 0, 0, now.Location())
	if !end.After(start) {
		end = end.Add(24 * time.Hour)
		if now.Before(start) {
			now = now.Add(24 * time.Hour)
		}
	}
	return (now.Equal(start) || now.After(start)) && now.Before(end)
}

func buildLeaderboard(records []Record, now time.Time) []LeaderboardItem {
	type agg struct {
		total int
		hours map[string]int
	}
	by := map[string]*agg{}
	for _, r := range records {
		a := by[r.Employee]
		if a == nil {
			a = &agg{hours: map[string]int{}}
			by[r.Employee] = a
		}
		a.total += r.Count
		hour := r.Timestamp.In(now.Location()).Format("2006-01-02 15")
		a.hours[hour] += r.Count
	}
	list := make([]LeaderboardItem, 0, len(by))
	for emp, a := range by {
		active := len(a.hours)
		if active == 0 {
			active = 1
		}
		list = append(list, LeaderboardItem{Employee: emp, AverageHourlyRate: float64(a.total) / float64(active), TotalCount: a.total})
	}
	sort.Slice(list, func(i, j int) bool {
		if list[i].AverageHourlyRate == list[j].AverageHourlyRate {
			return list[i].TotalCount > list[j].TotalCount
		}
		return list[i].AverageHourlyRate > list[j].AverageHourlyRate
	})
	for i := range list {
		list[i].Rank = i + 1
	}
	return list
}

func (s *Store) latestRecordLocked() time.Time {
	if len(s.records) == 0 {
		return time.Time{}
	}
	return s.records[len(s.records)-1].Timestamp
}

func (s *Store) employeeDashboard(name string, now time.Time) EmployeeDashboardResponse {
	s.mu.RLock()
	defer s.mu.RUnlock()
	leaderboard := buildLeaderboard(s.records, now)
	dayKey := now.Format("2006-01-02")
	today := make([]EmployeeRecordView, 0)
	dayTotal := 0
	for _, r := range s.records {
		if r.Employee != name {
			continue
		}
		if r.Timestamp.In(now.Location()).Format("2006-01-02") != dayKey {
			continue
		}
		dayTotal += r.Count
		today = append(today, EmployeeRecordView{
			Employee:  r.Employee,
			Count:     r.Count,
			DayTotal:  dayTotal,
			Timestamp: r.Timestamp,
		})
	}
	canSubmit := inOpenRange(now, s.settings.OpenTime, s.settings.CloseTime)
	msg := "当前允许录入"
	if !canSubmit {
		msg = "当前不在开放录入时间段"
	}
	return EmployeeDashboardResponse{
		ServerTime:    now,
		Employee:      name,
		CanSubmit:     canSubmit,
		SubmitMessage: msg,
		OpenTime:      s.settings.OpenTime,
		CloseTime:     s.settings.CloseTime,
		Reminder:      s.settings.Reminder,
		LatestRecord:  s.latestRecordLocked(),
		TodayRecords:  today,
		Leaderboard:   leaderboard,
	}
}

func (s *Store) adminDashboard(now time.Time) AdminDashboardResponse {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return AdminDashboardResponse{
		ServerTime:   now,
		Employees:    append([]string{}, s.settings.Employees...),
		OpenTime:     s.settings.OpenTime,
		CloseTime:    s.settings.CloseTime,
		Reminder:     s.settings.Reminder,
		LatestRecord: s.latestRecordLocked(),
		Leaderboard:  buildLeaderboard(s.records, now),
	}
}

func (s *Store) exportCSV(w http.ResponseWriter, now time.Time) error {
	s.mu.RLock()
	defer s.mu.RUnlock()
	cw := csv.NewWriter(w)
	if err := cw.Write([]string{"类型", "员工姓名", "录入条数", "当日总条数/平均小时速率", "录入总数", "录入时间/排名"}); err != nil {
		return err
	}
	dayKey := now.Format("2006-01-02")
	totalByEmployee := map[string]int{}
	for _, r := range s.records {
		if r.Timestamp.In(now.Location()).Format("2006-01-02") != dayKey {
			continue
		}
		totalByEmployee[r.Employee] += r.Count
		if err := cw.Write([]string{"记录", r.Employee, strconv.Itoa(r.Count), strconv.Itoa(totalByEmployee[r.Employee]), "", r.Timestamp.In(now.Location()).Format("2006-01-02 15:04:05")}); err != nil {
			return err
		}
	}
	for _, item := range buildLeaderboard(s.records, now) {
		if err := cw.Write([]string{"排行榜", item.Employee, "", fmt.Sprintf("%.2f", item.AverageHourlyRate), strconv.Itoa(item.TotalCount), strconv.Itoa(item.Rank)}); err != nil {
			return err
		}
	}
	cw.Flush()
	return cw.Error()
}

func sessionFromRequest(store *Store, r *http.Request) (Session, bool) {
	ck, err := r.Cookie(sessionCookieName)
	if err != nil {
		return Session{}, false
	}
	return store.getSession(ck.Value)
}

func main() {
	store, err := NewStore("data/records.jsonl", "data/config.json")
	if err != nil {
		log.Fatalf("init store failed: %v", err)
	}
	defer store.recordFile.Close()
	log.Printf("管理员密码: %s", adminPassword)

	mux := http.NewServeMux()
	mux.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/" {
			http.NotFound(w, r)
			return
		}
		w.Header().Set("Content-Type", "text/html; charset=utf-8")
		_, _ = w.Write([]byte(indexHTML))
	})

	mux.HandleFunc("/api/login/options", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodGet {
			http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
			return
		}
		cfg := store.settingsCopy()
		_ = json.NewEncoder(w).Encode(map[string]any{"employees": cfg.Employees})
	})

	mux.HandleFunc("/api/login/employee", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
			return
		}
		var req struct {
			Employee string `json:"employee"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, "invalid json", http.StatusBadRequest)
			return
		}
		req.Employee = strings.TrimSpace(req.Employee)
		if req.Employee == "" || !store.employeeExists(req.Employee) {
			http.Error(w, "employee invalid", http.StatusBadRequest)
			return
		}
		token, err := store.createSession("employee", req.Employee)
		if err != nil {
			http.Error(w, "login failed", http.StatusInternalServerError)
			return
		}
		http.SetCookie(w, &http.Cookie{Name: sessionCookieName, Value: token, Path: "/", HttpOnly: true, SameSite: http.SameSiteLaxMode})
		_ = json.NewEncoder(w).Encode(map[string]any{"ok": true})
	})

	mux.HandleFunc("/api/login/admin", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
			return
		}
		var req struct {
			Password string `json:"password"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, "invalid json", http.StatusBadRequest)
			return
		}
		if req.Password != adminPassword {
			http.Error(w, "password invalid", http.StatusUnauthorized)
			return
		}
		token, err := store.createSession("admin", "")
		if err != nil {
			http.Error(w, "login failed", http.StatusInternalServerError)
			return
		}
		http.SetCookie(w, &http.Cookie{Name: sessionCookieName, Value: token, Path: "/", HttpOnly: true, SameSite: http.SameSiteLaxMode})
		_ = json.NewEncoder(w).Encode(map[string]any{"ok": true})
	})

	mux.HandleFunc("/api/logout", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
			return
		}
		if ck, err := r.Cookie(sessionCookieName); err == nil {
			store.deleteSession(ck.Value)
		}
		http.SetCookie(w, &http.Cookie{Name: sessionCookieName, Value: "", Path: "/", MaxAge: -1, HttpOnly: true})
		_ = json.NewEncoder(w).Encode(map[string]any{"ok": true})
	})

	mux.HandleFunc("/api/me", func(w http.ResponseWriter, r *http.Request) {
		sess, ok := sessionFromRequest(store, r)
		if !ok {
			http.Error(w, "unauthorized", http.StatusUnauthorized)
			return
		}
		_ = json.NewEncoder(w).Encode(map[string]any{"role": sess.Role, "employee": sess.Employee})
	})

	mux.HandleFunc("/api/records", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
			return
		}
		sess, ok := sessionFromRequest(store, r)
		if !ok || sess.Role != "employee" {
			http.Error(w, "unauthorized", http.StatusUnauthorized)
			return
		}
		var req struct {
			Count int `json:"count"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, "invalid json", http.StatusBadRequest)
			return
		}
		now := time.Now()
		cfg := store.settingsCopy()
		if !inOpenRange(now, cfg.OpenTime, cfg.CloseTime) {
			http.Error(w, "not in open time", http.StatusForbidden)
			return
		}
		if req.Count <= 0 {
			http.Error(w, "count must be > 0", http.StatusBadRequest)
			return
		}
		if err := store.AddRecord(Record{Employee: sess.Employee, Count: req.Count, Timestamp: now}); err != nil {
			http.Error(w, "failed", http.StatusInternalServerError)
			return
		}
		_ = json.NewEncoder(w).Encode(map[string]any{"ok": true})
	})

	mux.HandleFunc("/api/employee/dashboard", func(w http.ResponseWriter, r *http.Request) {
		sess, ok := sessionFromRequest(store, r)
		if !ok || sess.Role != "employee" {
			http.Error(w, "unauthorized", http.StatusUnauthorized)
			return
		}
		_ = json.NewEncoder(w).Encode(store.employeeDashboard(sess.Employee, time.Now()))
	})

	mux.HandleFunc("/api/admin/dashboard", func(w http.ResponseWriter, r *http.Request) {
		sess, ok := sessionFromRequest(store, r)
		if !ok || sess.Role != "admin" {
			http.Error(w, "unauthorized", http.StatusUnauthorized)
			return
		}
		_ = json.NewEncoder(w).Encode(store.adminDashboard(time.Now()))
	})

	mux.HandleFunc("/api/admin/settings", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
			return
		}
		sess, ok := sessionFromRequest(store, r)
		if !ok || sess.Role != "admin" {
			http.Error(w, "unauthorized", http.StatusUnauthorized)
			return
		}
		var req struct {
			OpenTime     string   `json:"openTime"`
			CloseTime    string   `json:"closeTime"`
			ReminderMode string   `json:"reminderMode"`
			CustomTimes  []string `json:"customTimes"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, "invalid json", http.StatusBadRequest)
			return
		}
		if !validateHHMM(req.OpenTime) || !validateHHMM(req.CloseTime) {
			http.Error(w, "invalid open/close", http.StatusBadRequest)
			return
		}
		mode := strings.TrimSpace(req.ReminderMode)
		if mode != "hourly" && mode != "custom" {
			http.Error(w, "invalid reminder mode", http.StatusBadRequest)
			return
		}
		custom := make([]string, 0, len(req.CustomTimes))
		if mode == "custom" {
			seen := map[string]struct{}{}
			for _, t := range req.CustomTimes {
				t = strings.TrimSpace(t)
				if t == "" {
					continue
				}
				if !validateHHMM(t) {
					http.Error(w, "invalid custom time", http.StatusBadRequest)
					return
				}
				if _, ok := seen[t]; ok {
					continue
				}
				seen[t] = struct{}{}
				custom = append(custom, t)
			}
			sort.Strings(custom)
		}
		if err := store.setSettings(req.OpenTime, req.CloseTime, mode, custom); err != nil {
			http.Error(w, "save failed", http.StatusInternalServerError)
			return
		}
		_ = json.NewEncoder(w).Encode(map[string]any{"ok": true})
	})

	mux.HandleFunc("/api/admin/employees", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
			return
		}
		sess, ok := sessionFromRequest(store, r)
		if !ok || sess.Role != "admin" {
			http.Error(w, "unauthorized", http.StatusUnauthorized)
			return
		}
		var req struct {
			Name string `json:"name"`
		}
		if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
			http.Error(w, "invalid json", http.StatusBadRequest)
			return
		}
		name := strings.TrimSpace(req.Name)
		if name == "" {
			http.Error(w, "name empty", http.StatusBadRequest)
			return
		}
		if err := store.addEmployee(name); err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}
		_ = json.NewEncoder(w).Encode(map[string]any{"ok": true})
	})

	mux.HandleFunc("/api/export.csv", func(w http.ResponseWriter, r *http.Request) {
		sess, ok := sessionFromRequest(store, r)
		if !ok || sess.Role != "admin" {
			http.Error(w, "unauthorized", http.StatusUnauthorized)
			return
		}
		w.Header().Set("Content-Type", "text/csv; charset=utf-8")
		w.Header().Set("Content-Disposition", "attachment; filename=report.csv")
		if err := store.exportCSV(w, time.Now()); err != nil {
			http.Error(w, "export failed", http.StatusInternalServerError)
		}
	})

	mux.HandleFunc("/api/events", func(w http.ResponseWriter, r *http.Request) {
		sess, ok := sessionFromRequest(store, r)
		if !ok {
			http.Error(w, "unauthorized", http.StatusUnauthorized)
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
			now := time.Now()
			var payload any
			if sess.Role == "admin" {
				payload = map[string]any{"role": "admin", "dashboard": store.adminDashboard(now)}
			} else {
				payload = map[string]any{"role": "employee", "dashboard": store.employeeDashboard(sess.Employee, now)}
			}
			b, _ := json.Marshal(payload)
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
  <title>员工采集实时看板</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 20px; background: #f5f7fb; }
    .card { background:#fff; border-radius:10px; padding:16px; box-shadow:0 3px 12px rgba(0,0,0,.06); margin-bottom:16px; }
    .row { display:grid; grid-template-columns: repeat(2, minmax(0,1fr)); gap:16px; }
    .hidden { display:none; }
    input,select,button,textarea { width:100%; padding:8px; margin-top:6px; box-sizing:border-box; }
    button { background:#1864ff; color:#fff; border:none; border-radius:6px; cursor:pointer; }
    table { width:100%; border-collapse:collapse; margin-top:10px; }
    th,td { border-bottom:1px solid #eee; padding:8px; text-align:left; }
    .muted { color:#666; font-size:12px; }
    .ok { color:green; }
    .err { color:#c92a2a; }
  </style>
</head>
<body>
  <h1>员工采集实时看板</h1>

  <div id="loginCard" class="card">
    <h2>登录入口</h2>
    <div class="row">
      <div>
        <h3>员工登录</h3>
        <label>选择员工姓名</label>
        <select id="employeeSelect"></select>
        <button id="employeeLoginBtn">员工登录</button>
      </div>
      <div>
        <h3>管理员登录</h3>
        <label>管理员密码</label>
        <input id="adminPassword" type="password" placeholder="请输入管理员密码" />
        <button id="adminLoginBtn">管理员登录</button>
      </div>
    </div>
    <p id="loginMsg" class="muted"></p>
  </div>

  <div id="employeePanel" class="hidden">
    <div class="card">
      <h2>员工录入</h2>
      <p id="employeeWelcome" class="muted"></p>
      <label>本次采集条数</label>
      <input id="countInput" type="number" min="1" />
      <button id="submitBtn">提交记录</button>
      <p id="submitMsg" class="muted"></p>
    </div>

    <div class="card">
      <h2>我的当日录入信息</h2>
      <table>
        <thead><tr><th>员工姓名</th><th>录入条数</th><th>当日总条数</th><th>录入时间</th></tr></thead>
        <tbody id="myRows"></tbody>
      </table>
    </div>

    <div class="card">
      <h2>速率排行榜</h2>
      <table>
        <thead><tr><th>排名</th><th>员工姓名</th><th>平均小时速率</th><th>录入总数</th></tr></thead>
        <tbody id="rankRows"></tbody>
      </table>
    </div>
  </div>

  <div id="adminPanel" class="hidden">
    <div class="card">
      <h2>管理员配置</h2>
      <div class="row">
        <div>
          <label>开放录入开始时间</label>
          <input id="openTime" type="time" />
        </div>
        <div>
          <label>开放录入结束时间</label>
          <input id="closeTime" type="time" />
        </div>
      </div>
      <label><input type="radio" name="reminderMode" value="hourly" checked style="width:auto;" /> 整点提醒（默认）</label>
      <label><input type="radio" name="reminderMode" value="custom" style="width:auto;" /> 自定义提醒时间（HH:MM，逗号分隔）</label>
      <textarea id="customTimes" rows="2" placeholder="例如：09:00,10:30,14:00"></textarea>
      <button id="saveSettingsBtn">保存设置</button>
      <p id="adminMsg" class="muted"></p>
      <hr />
      <label>新增员工姓名</label>
      <input id="newEmployee" placeholder="输入员工姓名" />
      <button id="addEmployeeBtn">添加员工</button>
      <button id="exportBtn">导出 CSV</button>
    </div>

    <div class="card">
      <h2>实时速率排行榜</h2>
      <table>
        <thead><tr><th>排名</th><th>员工姓名</th><th>平均小时速率</th><th>录入总数</th></tr></thead>
        <tbody id="adminRankRows"></tbody>
      </table>
    </div>
  </div>

  <button id="logoutBtn" class="hidden">退出登录</button>

  <script>
    const loginMsg = document.getElementById('loginMsg');
    const loginCard = document.getElementById('loginCard');
    const employeePanel = document.getElementById('employeePanel');
    const adminPanel = document.getElementById('adminPanel');
    const logoutBtn = document.getElementById('logoutBtn');
    let me = null;
    let es = null;
    let reminderInterval = null;
    let repeatCheckInterval = null;
    let lastReminderKey = '';

    function setMsg(el, text, cls='muted') { el.className = cls; el.textContent = text; }
    function fmt(t) { return t ? new Date(t).toLocaleString() : '-'; }

    async function api(url, options={}) {
      const res = await fetch(url, options);
      if (!res.ok) throw new Error(await res.text());
      const ct = res.headers.get('content-type') || '';
      if (ct.includes('application/json')) return await res.json();
      return null;
    }

    function reminderTimes(d) {
      if (!d || !d.reminder) return [];
      if (d.reminder.mode === 'hourly') return Array.from({length:24}, (_,i)=>String(i).padStart(2,'0')+':00');
      return d.reminder.customTimes || [];
    }

    function notify(text) {
      alert(text);
      if ('Notification' in window && Notification.permission === 'granted') new Notification('采集提醒', {body:text});
    }

    async function ensureNotificationPermission() {
      if ('Notification' in window && Notification.permission === 'default') {
        try { await Notification.requestPermission(); } catch(e) {}
      }
    }

    function setupReminderWatcher(dashboard) {
      if (me.role !== 'employee') return;
      clearInterval(reminderInterval); clearInterval(repeatCheckInterval);
      const times = reminderTimes(dashboard);
      let latestAt = dashboard.latestRecord || '';
      reminderInterval = setInterval(() => {
        const now = new Date();
        const key = now.toTimeString().slice(0,5);
        if (!times.includes(key) || lastReminderKey === now.toDateString()+key) return;
        lastReminderKey = now.toDateString()+key;
        const snapshot = latestAt;
        notify('请填写当前采集条数。');
        clearInterval(repeatCheckInterval);
        repeatCheckInterval = setInterval(() => {
          if (latestAt && latestAt !== snapshot) { clearInterval(repeatCheckInterval); return; }
          notify('15秒内未检测到新录入，请尽快填写。');
        }, 15000);
      }, 1000);

      window.__updateLatestRecord = (v) => { latestAt = v; };
    }

    function renderLeaderboard(rowsEl, list) {
      rowsEl.innerHTML = (list||[]).map(x =>
        '<tr><td>' + x.rank + '</td><td>' + x.employee + '</td><td>' + x.averageHourlyRate.toFixed(2) + '</td><td>' + x.totalCount + '</td></tr>'
      ).join('') || '<tr><td colspan="4">暂无数据</td></tr>';
    }

    function renderEmployee(d) {
      document.getElementById('employeeWelcome').textContent = '当前员工：' + d.employee + '；开放时间 ' + d.openTime + ' - ' + d.closeTime + '；' + d.submitMessage;
      const myRows = document.getElementById('myRows');
      myRows.innerHTML = (d.todayRecords||[]).map(r =>
        '<tr><td>' + r.employee + '</td><td>' + r.count + '</td><td>' + r.dayTotal + '</td><td>' + fmt(r.timestamp) + '</td></tr>'
      ).join('') || '<tr><td colspan="4">今日暂无记录</td></tr>';
      renderLeaderboard(document.getElementById('rankRows'), d.leaderboard);
      window.__updateLatestRecord && window.__updateLatestRecord(d.latestRecord);
      setupReminderWatcher(d);
    }

    function renderAdmin(d) {
      document.getElementById('openTime').value = d.openTime;
      document.getElementById('closeTime').value = d.closeTime;
      document.querySelector('input[name="reminderMode"][value="' + d.reminder.mode + '"]').checked = true;
      document.getElementById('customTimes').value = (d.reminder.customTimes || []).join(',');
      renderLeaderboard(document.getElementById('adminRankRows'), d.leaderboard);
    }

    async function loadLoginOptions() {
      const d = await api('/api/login/options');
      const sel = document.getElementById('employeeSelect');
      sel.innerHTML = (d.employees||[]).map(x => '<option value="' + x + '">' + x + '</option>').join('');
    }

    function startEvents() {
      if (es) es.close();
      es = new EventSource('/api/events');
      es.onmessage = (evt) => {
        const data = JSON.parse(evt.data);
        if (data.role === 'employee') renderEmployee(data.dashboard);
        if (data.role === 'admin') renderAdmin(data.dashboard);
      };
    }

    async function refreshByRole() {
      if (me.role === 'employee') renderEmployee(await api('/api/employee/dashboard'));
      if (me.role === 'admin') renderAdmin(await api('/api/admin/dashboard'));
      startEvents();
      await ensureNotificationPermission();
    }

    async function checkMe() {
      try {
        me = await api('/api/me');
        loginCard.classList.add('hidden');
        logoutBtn.classList.remove('hidden');
        if (me.role === 'employee') employeePanel.classList.remove('hidden');
        if (me.role === 'admin') adminPanel.classList.remove('hidden');
        await refreshByRole();
      } catch {
        me = null;
        loginCard.classList.remove('hidden');
      }
    }

    document.getElementById('employeeLoginBtn').onclick = async () => {
      try {
        await api('/api/login/employee', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({employee: document.getElementById('employeeSelect').value})});
        setMsg(loginMsg, '员工登录成功', 'ok');
        location.reload();
      } catch (e) { setMsg(loginMsg, '员工登录失败：'+e.message, 'err'); }
    };

    document.getElementById('adminLoginBtn').onclick = async () => {
      try {
        await api('/api/login/admin', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({password: document.getElementById('adminPassword').value})});
        setMsg(loginMsg, '管理员登录成功', 'ok');
        location.reload();
      } catch (e) { setMsg(loginMsg, '管理员登录失败：'+e.message, 'err'); }
    };

    document.getElementById('submitBtn').onclick = async () => {
      const count = Number(document.getElementById('countInput').value);
      const msg = document.getElementById('submitMsg');
      if (!count || count <= 0) return setMsg(msg, '请输入正确条数', 'err');
      try {
        await api('/api/records', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({count})});
        setMsg(msg, '提交成功', 'ok');
        document.getElementById('countInput').value = '';
      } catch (e) { setMsg(msg, '提交失败：'+e.message, 'err'); }
    };

    document.getElementById('saveSettingsBtn').onclick = async () => {
      const mode = document.querySelector('input[name="reminderMode"]:checked').value;
      const customTimes = document.getElementById('customTimes').value.split(',').map(x=>x.trim()).filter(Boolean);
      try {
        await api('/api/admin/settings', {
          method:'POST', headers:{'Content-Type':'application/json'},
          body:JSON.stringify({openTime:document.getElementById('openTime').value, closeTime:document.getElementById('closeTime').value, reminderMode:mode, customTimes})
        });
        setMsg(document.getElementById('adminMsg'), '设置已保存', 'ok');
      } catch (e) { setMsg(document.getElementById('adminMsg'), '保存失败：'+e.message, 'err'); }
    };

    document.getElementById('addEmployeeBtn').onclick = async () => {
      try {
        await api('/api/admin/employees', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({name: document.getElementById('newEmployee').value})});
        setMsg(document.getElementById('adminMsg'), '员工已添加', 'ok');
        document.getElementById('newEmployee').value = '';
        await loadLoginOptions();
      } catch (e) { setMsg(document.getElementById('adminMsg'), '添加失败：'+e.message, 'err'); }
    };

    document.getElementById('exportBtn').onclick = () => { window.location.href = '/api/export.csv'; };

    logoutBtn.onclick = async () => { await api('/api/logout', {method:'POST'}); location.reload(); };

    loadLoginOptions();
    checkMe();
  </script>
</body>
</html>`
