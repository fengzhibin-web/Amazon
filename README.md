# 员工采集速率统计系统（在线/离线可用）

这是一个可打包成单文件可执行程序的 Go 系统，支持员工登录录入、管理员统一配置提醒与录入时间、实时排行榜、CSV 导出。

## 主要功能

- 员工登录：从预设名单选择姓名登录（无需密码）。
- 管理员登录：密码登录后可管理提醒、录入时段、员工名单。
- 员工录入：登录后仅填写“本次采集条数”，不再填写姓名，不支持修改历史记录。
- 员工面板：查看自己的当日录入明细（员工姓名、录入条数、当日总条数、录入时间）和速率排行榜。
- 管理员面板：实时查看排行榜、设置整点或自定义提醒、设置开放录入时段（例如 09:00-18:00）、新增员工。
- 定时提醒：默认整点提醒；提醒后每 15 秒检测是否有新录入，没有则继续提醒直到有新记录。
- 数据落盘：
  - `data/records.jsonl`：录入数据
  - `data/config.json`：员工名单与系统设置
- 导出：管理员可一键导出 CSV（包含员工录入信息与排行榜数据）。

## 默认管理员密码

程序启动时控制台会打印管理员密码：

`Adm!n-7Qx9P2Lm`

> 建议内网使用并妥善保管密码。

## 启动（开发模式）

需要 Go 1.22+：

```bash
go run .
```

访问 `http://localhost:8080`

## 从 GitHub 下载

```bash
git clone https://github.com/fengzhibin-web/Amazon.git
cd Amazon
```

## 打包成可执行文件

### Windows

```bash
set GOOS=windows
set GOARCH=amd64
go build -o employee-tracker.exe .
```

### macOS

```bash
GOOS=darwin GOARCH=amd64 go build -o employee-tracker-macos .
```

### Linux

```bash
GOOS=linux GOARCH=amd64 go build -o employee-tracker-linux .
```

## 局域网部署（统计主机）

1. 在统计主机运行可执行文件。
2. 本机先访问 `http://localhost:8080` 确认服务正常。
3. 员工通过 `http://统计主机IP:8080` 访问。
4. 若访问失败，放行防火墙 TCP 8080 入站规则。
