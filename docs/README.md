# 店铺物流表现看板

静态数据看板：按 market / PID / PH NO / 日期筛选，查看入库时效与重量差异异常。

## 在线访问

开启 GitHub Pages 后，地址为：

`https://<你的用户名>.github.io/<仓库名>/`

## 本地更新数据

1. 替换 `datasource/daily.csv`
2. 运行 `python scripts/build_logistics_dashboard.py`
3. 将生成的 `dashboard/data/*` 同步到本仓库的 `docs/data/`（或重新复制 `dashboard/index.html` → `docs/`）
4. `git add` / `commit` / `push`

## 说明

当前为演示样例数据。页面为纯静态文件，无需后端。
