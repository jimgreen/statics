# Report Portal

这是一个独立的轻量 Web 报表门户，用于把已下载的报表项目先串成可运行的“定义 → 运行 → 查询 → 导出”闭环。

## 功能

- Web 前端定义报表：名称、说明、只读 SQL、表格/折线/柱状/饼图、横轴/数值列、输出格式。
- 网格版式设计器：编辑单元格内容，Shift 选择区域，复制/粘贴、合并/拆分、字号、粗体/斜体、文字色、填充色、对齐和线型；版式保存在报表定义中，并输出到 HTML 和 Excel 的“版式”工作表。
- 单元格/图元独立数据源：每个网格元素可单独选择静态内容、主查询字段、独立只读 SQL 或统计公式；元素类型支持文本、数值/字段、动态表格和折线/柱状/饼图。动态表格不预设行数，按独立 SQL 返回结果在指定单元格展开。
- Excel 风格统计公式：每行一条“名称 = 公式”，支持 `MAX`、`MIN`、`AVERAGE`、`SUM`、`COUNT` 和按横轴梯形积分的 `INTEGRAL(数值列, 时间列)`；也支持 `B2:B100` 单元格范围。
- 后端使用 SQLite 保存报表定义和运行历史；默认数据库文件为 `D:\codex\sqldb\realtime\statics\report.db`，外部默认数据库不会创建演示数据。
- 异步执行报表 SQL，生成 HTML、CSV 和带原生图表的 XLSX。
- 定时调度线程：按报表的“间隔分钟”执行；支持启动、暂停、恢复和停止。
- 运行历史、数据表格预览、HTML/Excel 下载。
- 结果页显示统计卡片，HTML 报表带统计值表，Excel 额外生成“统计值”工作表。
- `REPORT_DB_PATH` 环境变量可指定已有 SQLite 数据库，例如：

  ```powershell
  $env:REPORT_DB_PATH='D:\data\business.db'
  python .\report_portal\app.py
  ```

  如果只需要更换默认数据库目录，可设置 `REPORT_DB_DIR`；数据库文件名固定为 `report.db`：

  ```powershell
  $env:REPORT_DB_DIR='D:\data\reports'
  python .\report_portal\app.py
  ```

## 启动

在 `D:\codex\statics` 下执行：

```powershell
python -m pip install -r .\report_portal\requirements.txt
python .\report_portal\app.py
```

打开 <http://127.0.0.1:8766>。如果需要 Excel 导出，请确保环境安装了 `openpyxl`；当前开发环境已具备该依赖。

## 后续接入商用数据库

当前门户直接针对 SQLite，目的是先把 Web 流程和权限边界跑通。若要接入 SQL Server、Oracle、MySQL 或 PostgreSQL，可将 `connection()` 抽象为 SQLAlchemy engine，并为每个数据源保存只读连接配置；前端的报表定义、调度、结果和下载接口无需改变。
