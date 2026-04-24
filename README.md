# 员工采集速率统计系统（在线/离线可用）

这是一个**可直接打包成单文件可执行程序**的轻量系统，用来记录每位员工每次采集条数，并实时展示：

- 每位员工的填写记录（会持续累计并落盘保存）
- 每小时速率排行榜（按“当前小时条数”排序）
- 平均速率统计（总条数 / 活跃小时数）

## 功能说明

- 员工端录入：输入姓名 + 本次条数
- 定时提醒：支持按分钟设置弹窗提醒填写（可开启/关闭）
- 管理看板：实时更新排行榜（SSE 推送）
- 数据落盘：`data/records.jsonl`，重启后自动恢复
- 在线/离线：
  - **在线部署**：放在公司服务器，所有员工浏览器访问
  - **离线局域网**：在一台电脑启动后，其他电脑通过 `http://该电脑IP:8080` 访问
  - **单机离线**：只在一台电脑本地使用 `http://localhost:8080`

## 启动（开发模式）

> 需要 Go 1.22+

```bash
go run .
```

打开 `http://localhost:8080`

## 从 GitHub 下载到本地

你有两种方式：

1. **推荐：git clone**

```bash
git clone <你的仓库地址>
cd Amazon
```

2. **不装 Git：下载 ZIP**

- 在 GitHub 仓库页点击 `Code` -> `Download ZIP`
- 解压后进入项目目录

## 打包（无需本地开发环境即可运行）

在你的打包机上执行：

### Windows 可执行文件

```bash
set GOOS=windows
set GOARCH=amd64
go build -o employee-tracker.exe .
```

### macOS 可执行文件

```bash
GOOS=darwin GOARCH=amd64 go build -o employee-tracker-macos .
```

### Linux 可执行文件

```bash
GOOS=linux GOARCH=amd64 go build -o employee-tracker-linux .
```

把打包产物发给新电脑，双击/命令行启动即可，不需要再配 Go 环境。

## 使用建议（你这个场景）

1. 指定一台“统计主机”运行程序（可执行文件）。
2. 让员工都通过浏览器访问该主机地址进行填写。
3. 主管打开同一地址即可实时看排行。
4. 每天定时备份 `data/records.jsonl`。
5. 员工可在左侧开启“定时提醒填写”，到点后会浏览器弹窗提醒。

## 后续可扩展

- 登录与权限（员工账号、主管账号）
- 数据导出 Excel/CSV
- 按天/按班次报表
- 补录与审核机制
- 部门维度排行榜
