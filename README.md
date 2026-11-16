# 自动报表开源组件

本目录按“网页定义/查询 + Python 自动生成 + 定时任务”的需求下载了以下 GitHub 项目。默认数据源可以使用 SQLite；商用数据库通过各项目的数据库驱动或 SQLAlchemy 连接 URL 接入。

## 已实现的 Web 门户

`report_portal/` 是在这些组件基础上新增的轻量可运行 MVP，提供 Web 前后端闭环：报表定义（只读 SQL、图表、输出格式和间隔）、Excel 风格统计公式（最大值、最小值、平均值、求和、计数、历史曲线梯形积分）、可编辑单元格版式（字体、颜色、对齐、边框、合并/拆分和区域复制/粘贴）、异步报表执行、调度进程启动/暂停/恢复/停止、运行历史、结果表格查询以及 HTML/Excel/CSV 下载。它默认使用本地 SQLite，并自动生成 `demo_sales` 示例表。启动方式：

```powershell
python .\report_portal\app.py
```

浏览器访问 <http://127.0.0.1:8766>。详细说明见 [report_portal/README.md](report_portal/README.md)。

## 已下载项目

| 目录 | 项目 | 适合的功能 | 许可证 |
|---|---|---|---|
| `seal-report/` | [Ariacom/Seal-Report](https://github.com/ariacom/Seal-Report) | 完整的 Web 报表服务器、报表设计器、SQLite/MySQL/PostgreSQL/SQL Server/Oracle 等数据源、HTML/Excel/PDF、表格/透视表/折线/饼图/柱状图、定时输出 | MIT |
| `redash/` | [getredash/redash](https://github.com/getredash/redash) | Web SQL 查询、结果查询、可视化和仪表盘、定时刷新、Excel/CSV 结果下载、SQLite 和大量商用 SQL 数据源 | BSD 风格许可证 |
| `reportbro-designer/` | [jobsta/reportbro-designer](https://github.com/jobsta/reportbro-designer) | 可嵌入业务系统的浏览器端 PDF/XLSX 模板设计和预览 | AGPL-3.0 |
| `reportbro-lib/` | [jobsta/reportbro-lib](https://github.com/jobsta/reportbro-lib) | Python 3.10+ 按模板生成 PDF/XLSX，支持表格、分页、样式、图片和表达式 | AGPL-3.0 或商业许可 |
| `python-automation/` | [adriasancheza/mysql-report-automation](https://github.com/adriasancheza/mysql-report-automation) | YAML 定义 SQL、参数和格式，使用 SQLAlchemy 从 SQLite/MySQL 等数据库生成格式化 Excel/CSV，由 cron 或 Windows Task Scheduler 定时执行 | MIT |

## 推荐组合

如果优先快速落地，使用 `seal-report/` 作为完整的 Web 报表服务器：它已经覆盖网页报表定义、结果查询、图表、HTML/Excel 输出和调度，SQLite 可直接使用，商用数据库通过 ODBC/原生驱动接入。

如果必须由 Python 控制生成流程，则使用 `redash/` 保存查询、可视化和网页结果，使用 `reportbro-designer/` 保存可视模板，再由 `reportbro-lib/` 或 `python-automation/` 生成 Excel。Python 任务可以通过数据库 URL 读取 SQLite，例如 `sqlite:///D:/data/report.db`；商用数据库只需替换为对应的 SQLAlchemy URL 和驱动。

## 许可证提醒

`reportbro-designer` 和 `reportbro-lib` 的开源许可是 AGPL-3.0；将它们集成到闭源商用系统前，需要评估 AGPL 义务或购买 ReportBro 商业许可。`Seal Report` 和 `python-automation` 是 MIT 许可，更适合直接用于商用项目。各目录中的原始 `LICENSE` 和第三方声明文件优先于本说明。

## 当前下载状态

所有目录均为 GitHub 的浅克隆（`--depth 1`），保留各自的 `origin` 远程地址和当前默认分支；未安装依赖、未连接业务数据库，也未执行生产调度。
