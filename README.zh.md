# git-calendar-sync

[![README 中文](https://img.shields.io/badge/README-中文-2563eb?style=for-the-badge)](README.zh.md)
[![README English](https://img.shields.io/badge/README-English-16a34a?style=for-the-badge)](README.md)

把 git commit 记录自动转换成日历事件。支持 `.ics` 导出、Google Calendar、Microsoft Graph API 和经典版 Outlook。

## 两种使用方式

### 方式 1 - 下载 exe，免装 Python

从 [Releases](https://github.com/chushug/git-calendar-sync/releases) 下载 `git-calendar-sync-windows.zip`，解压后双击 **git-calendar-sync.exe**。

GUI 可以直接填写仓库路径、同步方式和凭据，并在下次启动时保留这些设置。界面默认使用 Catppuccin Latte，小窗口下左侧选项栏也支持滚动查看。Google Calendar 和 Microsoft Graph 的凭据可以直接在应用内填写，不需要额外准备 `.env`。

打包版不需要 Python，但本机仍然需要安装 Git，程序才能读取 commit 历史。

如果想每天自动同步，而不打开 GUI，exe 也可以作为 CLI 使用：

```powershell
git-calendar-sync.exe sync --method google --days 1 --repo "C:\Projects\MyApp"
```

也可以注册 Windows 计划任务：

```powershell
.\scheduler\setup_windows.ps1 -Method google -Repo "C:\Projects\MyApp"
```

### 方式 2 - 从源码运行

```bash
pip install -r requirements.txt
python sync.py generate --days 7
python sync.py sync --method google --days 1
python gui.py
```

源码运行时，凭据也可以放在 `.env` 中。后面的安装需求和源码快速开始只适用于这条源码路径；exe 用户不需要安装 Python 依赖，也不需要准备 `.env`。

## 功能特性

- 导出为标准 `.ics` 文件（无需授权，任何日历应用都能导入）
- 直接同步到 Google Calendar、Microsoft Calendar（新版 Outlook）或经典版 Outlook
- Textual TUI 界面，支持设置持久化、Catppuccin Latte 默认主题和实时输出日志
- 小窗口下左侧选项栏可滚动，支持 `PageUp` / `PageDown`
- Windows 打包版既能启动 GUI，也能直接执行无界面的 CLI 命令
- 幂等操作：多次运行不会产生重复事件
- 稳定 UID：重新导入同一个 `.ics` 文件会更新事件，而不是重复创建
- 所有事件标记为 **空闲（Free）**，不会占用你的日历时间
- 支持 Windows Task Scheduler 和 cron 自动化

## 兼容性说明

| 功能 | 详情 |
|------|------|
| `.ics` 生成 | 在 Windows 上测试通过；macOS/Linux 应该可用但未测试 |
| `.ics` 导入目标 | 经典版 Outlook、Apple Calendar、Google Calendar、Thunderbird |
| 直接同步 — Google Calendar | 任意系统（需 Python 3.9+）|
| 直接同步 — Microsoft Graph | 任意系统（需 Python 3.9+），支持**新版 Outlook for Windows**、Outlook 网页版、Microsoft 365 |
| 直接同步 — 经典版 Outlook COM | **仅限 Windows**，需要 Office 2016 及以上版本 |
| 新版 Outlook for Windows（COM）| **不支持**。Microsoft 官方文档明确说明新版 Outlook 不支持 COM/VBA/VSTO 自动化 |
| 每日自动同步 | Windows Task Scheduler；Linux/macOS 用 cron |

> **关于新版 Outlook：** 新版 Outlook for Windows 是基于 Web 的应用，没有 COM 接口。
> 请改用 `--method graph`（Microsoft Graph API）或 `--method google`（如果已绑定 Google 账号）。

## 安装要求

### Windows exe 打包版

- Windows
- 已安装 Git，并且 `git` 命令在 `PATH` 中可用
- 不需要安装 Python

只有在使用计划任务脚本或经典版 Outlook COM 时，才需要设置 PowerShell 执行策略：

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

### 从源码运行

- Python 3.9+
- Git

安装依赖：

```bash
pip install -r requirements.txt
```

如果只需要生成 `.ics`（不需要 Google/Graph 同步）：

```bash
pip install python-dotenv textual
```

Windows 下使用 `outlook-com` 方法或计划任务脚本时，需要设置 PowerShell 执行策略：

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

> 不要使用 `Unrestricted`，`RemoteSigned` 是所需的最低权限，更加安全。

## 源码快速开始

本节默认你是从源码运行，因此示例使用 `python sync.py`。
如果你下载的是 exe，请按上面的 GUI 流程操作；下面的 `.env` 配置只适用于源码运行。

### 方式 A — 生成 `.ics` 文件（无需授权）

```bash
python sync.py generate --days 7
# 默认输出路径：
#   Windows/macOS: ~/Documents/CommitCalendar/YYYY-MM-DD_commits.ics
#   Linux:         $XDG_DATA_HOME/commit-calendar/YYYY-MM-DD_commits.ics
#   Linux 备用:    ~/.local/share/commit-calendar/YYYY-MM-DD_commits.ics

python sync.py generate --days 7 --out ~/Desktop/commits.ics  # 自定义路径
```

双击 `.ics` 文件导入到日历应用即可。

> **注意：** 导入 `.ics` 是一次性快照，不是持续同步。
> 需要更新时重新运行命令并重新导入。
> 要实现持续同步，请使用方式 B 或 C。

### 方式 B — 同步到 Google Calendar

**一次性设置：**

1. 进入 [Google Cloud Console](https://console.cloud.google.com/)
2. 创建项目 → 启用 **Google Calendar API**
3. 创建 **OAuth 2.0 凭据**（桌面应用）→ 下载 JSON 文件
4. 复制 `.env.example` 为 `.env`，设置：
   ```
   GOOGLE_CREDENTIALS_FILE=path/to/client_secret.json
   ```
5. 运行授权（会打开浏览器一次）：
   ```bash
   python sync.py sync --method google --setup
   ```

**日常使用：**

```bash
python sync.py sync --method google --days 1
```

> **新版 Outlook 用户提示：** 在新版 Outlook 设置里添加 Google 账号（设置 → 账号 → 添加账号），Google Calendar 事件会自动显示。

### 方式 C — 同步到 Microsoft Calendar（新版 Outlook / Microsoft 365）

使用 Microsoft Graph API，无需 COM，完全免费。

**一次性 Azure 配置（免费）：**

1. 进入 [portal.azure.com](https://portal.azure.com) → **应用注册** → 新注册
   - 名称：`git-calendar-sync`（随意）
   - 支持的账户类型：**任何组织目录中的账户 + 个人 Microsoft 账户**
2. 身份验证 → 添加平台 → **移动和桌面应用程序**
   - 重定向 URI 添加：`https://login.microsoftonline.com/common/oauth2/nativeclient`
3. API 权限 → 添加 → Microsoft Graph → 委托权限 → **Calendars.ReadWrite**
   - 如果组织策略要求，再授予管理员同意
4. 复制**应用程序（客户端）ID** → 添加到 `.env`：
   ```
   GRAPH_CLIENT_ID=你的客户端ID
   ```
5. 运行授权（会显示验证码，访问指定 URL 即可，不需要跳转浏览器）：
   ```bash
   python sync.py sync --method graph --setup
   ```

**日常使用：**

```bash
python sync.py sync --method graph --days 1
```

### 方式 D — 经典版 Outlook COM（仅 Windows）

需要安装经典版 Microsoft Outlook（Office 2016+）。**不支持新版 Outlook for Windows**。

```bash
python sync.py sync --method outlook-com --days 1
python sync.py sync --method outlook-com --days 1 --dry-run  # 预览模式
```

## 完整参数说明

### `generate` — 生成 .ics 文件

```
python sync.py generate
  [--days N]          最近 N 天的 commit（默认：1）
  [--since DATE]      开始日期 YYYY-MM-DD（设置了 --days 则忽略）
  [--until DATE]      结束日期/时间（设置了 --days 则忽略）
  [--repo PATH]       Git 仓库路径（默认：当前目录）
  [--out FILE]        输出 .ics 路径（默认：平台特定路径）
  [--duration MINS]   事件时长（分钟，默认：15）
  [--dry-run]         只打印，不生成文件
```

### `sync` — 直接同步到日历服务

```
python sync.py sync --method google|graph|outlook-com
  [--days N]          最近 N 天的 commit（默认：1）
  [--since DATE]      开始日期 YYYY-MM-DD
  [--until DATE]      结束日期/时间
  [--repo PATH]       Git 仓库路径（默认：当前目录）
  [--credentials FILE] [google] client_secret.json 路径
  [--category LABEL]  [outlook-com] Outlook 分类标签（默认："Git Commit"）
  [--setup]           运行首次 OAuth 授权（仅 google/graph）
  [--dry-run]         只打印，不写入
```

## 日历事件格式

每个 commit 映射为以下日历事件：

```
标题:     [仓库名] commit 提交信息
开始时间: commit 时间戳（committer date，UTC）
结束时间: 开始时间 + 时长（默认 15 分钟）
UID:      <完整 hash>@<仓库名>.commitcal（稳定，永不改变）
描述:     Hash: <完整哈希>
          Author: <邮箱>
          Repo: <名称>
          Branch: <分支>
          Remote: <origin URL>
分类:     Git, Commit, <仓库名>
忙碌状态: 空闲（不会占用日历时间）
提醒:     关闭
```

**时长选项：** 通过 `--duration` 设置 5 / 10 / 15（默认）/ 30 / 60 分钟。

**UID 稳定性：** 同一个 commit 总是生成相同的 UID。重新导入同一个 `.ics` 到支持 UID 匹配的日历应用时，会更新现有事件而不是重复创建。

## 自动每日同步

### Windows Task Scheduler

```powershell
# 注册每日 23:30 运行的计划任务
.\scheduler\setup_windows.ps1 -Method google

# 可选方法：google | graph | outlook-com
.\scheduler\setup_windows.ps1 -Method graph -Time 22:00

# 删除任务
.\scheduler\setup_windows.ps1 -Uninstall
```

> **说明：** 本工具不是后台服务。`setup_windows.ps1` 注册一个每天定时运行一次的计划任务，两次执行之间不会有任何进程在后台运行。

### macOS / Linux（cron）

```bash
crontab -e
# 添加（每天 23:30 运行）：
30 23 * * * cd /你的仓库路径 && python /git-calendar-sync路径/sync.py sync --method google --days 1
```

## 去重机制

对同一时间段多次运行不会产生重复事件。

| 方法 | 去重策略 |
|------|----------|
| `generate`（.ics）| 每次重新生成完整文件；支持 UID 匹配的日历应用重新导入时会更新现有事件 |
| `google` | 在 `extendedProperties.private` 中存储 `gitHash`，插入前查询 |
| `graph` | 在事件描述中搜索 commit hash，插入前检查 |
| `outlook-com` | 扫描现有日历条目的描述字段中的 `Hash: <sha>` |

## 输出文件位置

| 平台 | 默认 `.ics` 路径 |
|------|-----------------|
| Windows | `~/Documents/CommitCalendar/YYYY-MM-DD_commits.ics` |
| macOS | `~/Documents/CommitCalendar/YYYY-MM-DD_commits.ics` |
| Linux | `$XDG_DATA_HOME/commit-calendar/YYYY-MM-DD_commits.ics`（备用：`~/.local/share/commit-calendar/YYYY-MM-DD_commits.ics`）|

可以用 `--out /你的路径/file.ics` 覆盖默认路径。

> **提示：** 不要把 `.ics` 文件输出到 Git 仓库目录内，除非你确实想发布这个日历文件。

## 安全与隐私

生成的事件中包含 commit 提交信息、作者邮箱、分支名和远程 URL。

**在公开或分享 `.ics` 文件前，请确认其中不包含：**
- 私有仓库名称或内部分支名
- commit 提交信息中的 Ticket ID 或内部标识符
- 个人邮箱地址

GUI 会把非敏感偏好设置保存在 `~/.git-calendar-sync/settings.json`，包括你填写的凭据文件路径和 Azure client ID。OAuth token 文件（`~/.git-calendar-sync/`）和 `.env` 已加入 `.gitignore`，不会被提交到版本库。

## 常见问题排查

**`ModuleNotFoundError: No module named 'dotenv'`**
```bash
pip install python-dotenv
```

**google 方法：`FileNotFoundError: client_secret.json`**
从 Google Cloud Console 下载 OAuth 凭据。使用 exe GUI 时，在 Google 凭据输入框里填写 JSON 路径；从源码运行时，在 `.env` 中设置 `GOOGLE_CREDENTIALS_FILE`，或传入 `--credentials`。

**graph 方法：`GRAPH_CLIENT_ID not set`**
使用 exe GUI 时，在 Microsoft Graph 输入框里填写 Azure Application/client ID；从源码运行时，把它添加到 `.env`。

**outlook-com 方法：`Cannot complete the operation. You are not connected.`**
经典版 Outlook 必须已安装并配置好账号。新版 Outlook for Windows 不支持此方法。

**PowerShell 脚本被阻止**
```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

**小窗口下 GUI 底部选项看不到**
使用左侧滚动条，或按 `PageUp` / `PageDown` 翻页。`Ctrl+Home` 和 `Ctrl+End` 可以跳到选项栏顶部或底部。

## 项目结构

```
git-calendar-sync/
├── gui.py                        # Textual TUI 入口
├── sync.py                       # CLI 入口（generate / sync）
├── core/
│   ├── models.py                 # Commit 数据类，带稳定 UID
│   └── git_log.py                # git log 解析，平台路径处理
├── methods/
│   ├── ics.py                    # generate：RFC 5545 .ics 输出
│   ├── google_cal.py             # sync：Google Calendar API
│   ├── graph_api.py              # sync：Microsoft Graph API
│   └── outlook_com.py            # sync：经典版 Outlook COM（Windows）
├── scheduler/
│   ├── setup_windows.ps1         # Task Scheduler 注册脚本
│   └── _outlook_com_worker.ps1   # PowerShell COM 工作脚本
├── build.ps1                     # 构建 Windows .exe 发行版
├── git-calendar-sync.spec        # PyInstaller 配置
├── requirements.txt
├── .env.example
├── .gitignore
├── README.md                     # 英文文档
└── README.zh.md                  # 中文文档（本文件）
```

## 已知限制

- 新版 Outlook for Windows 不支持 COM 同步（这是 Microsoft 的设计限制）
- 本地导入 `.ics` 是快照，不是持续订阅
- 不支持多账号 Outlook（使用默认配置文件）
- Rebase/amend 后的 commit（hash 变化）会被视为新 commit
- macOS 和 Linux 上的跨平台 `.ics` 生成未经测试

## 许可证

MIT
