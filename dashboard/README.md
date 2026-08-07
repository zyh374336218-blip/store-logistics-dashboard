# 店铺物流表现看板（本机打开）

## 怎么用（托管方式 A）

1. 更新数据：替换 `datasource/daily.csv`（正式全量替换演示样例即可）
2. 生成看板数据：

```bash
python scripts/build_logistics_dashboard.py
```

3. 用浏览器打开 `dashboard/index.html`（双击或拖到浏览器）

不需要服务器、不需要公网链接。以后若要分享给同事，再选 GitHub Pages 或内网托管。

## 规则摘要

| 项 | 规则 |
|----|------|
| 异常 | `Weight Variance > 0.2` |
| PID | 跨 SA/AE 同号合并 |
| PH NO / 日期 | 与 market、PID 一起过滤；影响 KPI 与视图 A/B/C |
| 质检 | 空值或重复 PH NO → build 报错「数据源有误」 |
| 小数 | 展示统一两位 |

## 文件

- `datasource/daily.csv` — 源数据
- `scripts/build_logistics_dashboard.py` — 清洗与质检
- `dashboard/data/logistics.js` — 前端数据（自动生成）
- `dashboard/index.html` — 看板页面
