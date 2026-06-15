# 快手视频评论抓取器 🎬

> 基于 GraphQL API + CDP 浏览器的快手视频评论批量抓取工具  
> 支持根评论、子评论、GUI 一键操作、分享链接自动解析

## ✨ 功能

- 🔌 **自动连接浏览器** — 一键启动 Edge/Chrome CDP 调试端口，免手动配置
- 🍪 **Cookie 持久化** — 登录一次，自动保存，后续离线可用
- 🔗 **分享链接解析** — 支持 `v.kuaishou.com` 短链和 `kuaishou.com/short-video/` 标准链接
- 💬 **根评论 + 子评论** — GraphQL API 拉取根评论，CDP DOM 提取子回复
- 📊 **Excel 导出** — 蓝色表头、冻结首行、自动筛选、按点赞降序
- 🖥 **GUI 界面** — tkinter 可视化操作，实时日志、进度条、批量模式
- 📝 **视频标题文件名** — 自动用原视频标题命名 Excel（最长 10 字）

## 🚀 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

> 无需手动安装 Playwright 浏览器，工具会自动复用系统 Edge。

### 2. 运行 GUI（推荐）

```bash
python gui.py
```

**操作流程：**

1. 点击 **🔌 自动连接** → 弹出 Edge 浏览器
2. 在浏览器中**登录快手**（kuaishou.com）
3. 粘贴视频 URL/ID（每行一个，支持分享短链）
4. 勾选「📎 包含子评论」（可选，需要 CDP 浏览器）
5. 点击 **▶ 开始抓取** → 等待完成 → 点击 **📂 打开输出目录**

### 3. 命令行方式

```bash
# 单个视频
python run.py https://www.kuaishou.com/short-video/3xqi7iru65ut3bk

# 批量（多个 ID 或 URL）
python run.py 3xqi7iru65ut3bk 3xgem97p6hychb9

# 从文件读取
python run.py -f videos.txt

# 从配置文件
python run.py -c config.json
```

### 4. Python 代码调用

```python
from ks_scraper import KuaishouCommentScraper

scraper = KuaishouCommentScraper()
comments = scraper.scrape("3xqi7iru65ut3bk")
scraper.export_excel(comments, "output.xlsx")
```

## 📋 Excel 输出格式

| 序号 | 用户名 | 头像链接 | 评论内容 | 评论时间 | 点赞量 | 回复数 | 评论ID |
|------|--------|----------|----------|----------|--------|--------|--------|

- 🔵 蓝色表头，冻结首行，自动筛选
- 📊 按点赞量降序排列
- 📝 末尾汇总行（总评论数、总点赞）

## ⚙️ 配置说明 (config.json)

```json
{
  "videos": ["视频ID或URL"],
  "cdp": "http://127.0.0.1:28800",
  "output_dir": "ks_output",
  "max_pages": 50,
  "delay": 0.5,
  "batch_delay": 2.0
}
```

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `videos` | 必填 | 视频 ID 或 URL 列表 |
| `cdp` | `http://127.0.0.1:28800` | CDP 调试端口 |
| `output_dir` | `ks_output` | 输出目录 |
| `max_pages` | 50 | 每个视频最大翻页数 |
| `delay` | 0.5 | 翻页间隔（秒） |
| `batch_delay` | 2.0 | 视频间间隔（秒） |

## 🏗 技术架构

```
GUI (gui.py) ──→ KuaishouCommentScraper (ks_scraper.py)
                      │
     ┌────────────────┼────────────────┐
     ▼                ▼                 ▼
  分享链接解析    GraphQL API        CDP 浏览器
  (302 重定向)   (commentListQuery)  (Playwright)
     │                │                 │
     ▼                ▼                 ▼
  提取视频 ID    根评论分页抓取    子评论 DOM 提取
     │                │                 │
     └────────────────┴─────────────────┘
                      │
                      ▼
              Excel 导出 (openpyxl)
```

- **GraphQL 端点**: `https://www.kuaishou.com/graphql`
- **核心查询**: `commentListQuery(photoId, pcursorV2)` → 分页拉取根评论
- **子评论**: 通过 CDP 浏览器点击「展开回复」→ DOM 提取
- **反爬策略**: 禁用 `webDriver` 检测、移除遮罩层、请求间隔 ≥ 0.5s

## ⚠️ 注意事项

- 🍪 **需要快手登录态** — 工具通过 CDP 复用浏览器 cookie 或从 `ks_cookies.json` 加载
- 👶 **子评论需要 CDP 浏览器** — 快手 GraphQL API 不暴露子评论查询，需浏览器 DOM 抓取
- 🛡 **不要频繁请求** — 建议 `delay` ≥ 0.5 秒，避免触发风控
- 🔗 **支持的链接格式**:
  - `https://v.kuaishou.com/xxxxx` （分享短链）
  - `https://www.kuaishou.com/short-video/3xqi7iru65ut3bk`
  - 纯视频 ID: `3xqi7iru65ut3bk`

## 📦 打包为 EXE

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name KuaishouScraper gui.py
```

输出在 `dist/KuaishouScraper.exe`，免安装 Python 即可运行。

## 📄 License

本项目基于逆向方式实现，仅供学习、研究、个人实验和内部验证使用，不提供任何商业授权、稳定性保证或可用性保证。 作者及仓库维护者不对因使用、修改、分发、部署或依赖本项目而产生的任何直接或间接损失、账号封禁、数据丢失、法律风险或第三方索赔负责。

请勿将本项目用于违反服务条款、协议、法律法规或平台规则的场景。商业使用前请自行确认你是否获得了作者的书面许可。
