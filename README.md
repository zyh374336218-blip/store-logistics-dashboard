# data-analysis

月度 Partner Level 分析 + 店铺物流表现看板。

## 目录

| 路径 | 用途 |
|------|------|
| `data/current/` | Partner Summary **当前分析用** Excel |
| `data/YYYY-MM/` | Partner 按月归档 |
| `datasource/daily.csv` | 物流看板源数据（演示样例可替换） |
| `dashboard/` | 物流看板（本机打开 `index.html`） |
| `outputs/` | 导出的筛选/汇总结果 |
| `scripts/load_partner_summary.py` | Partner 读表 |
| `scripts/build_logistics_dashboard.py` | 物流看板数据构建 + 质检 |
| `.cursor/skills/partner-summary-analysis/` | Partner 分析约定 |

## 物流看板（本机）

```bash
python scripts/build_logistics_dashboard.py
```

然后浏览器打开 `dashboard/index.html`。说明见 [dashboard/README.md](dashboard/README.md)。

## 每月更新

1. 将新文件放到 `data/YYYY-MM/`（例如 `data/2026-08/`）
2. 用新文件覆盖 `data/current/Summary- CN Partner Level.xlsx`
3. 在 Cursor 打开本文件夹，直接提问即可

## 依赖

```bash
pip install -r requirements.txt
```

## 自检

```bash
python scripts/load_partner_summary.py
```

## 提问示例

- 哪个 `id_partner` 的 Order-to-Ship Avg 最大？（默认按 **MTD**）
- 全部 partner 的 Order-to-Ship Avg 平均数是多少？
- 筛选 Week-2 中 Order-to-Ship Avg > 5 的 partner，并导出到 outputs
