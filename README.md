# 古城公司资产差异分析报告

在线访问：[https://zejunwu12.github.io/gucheng-asset-diff-report/](https://zejunwu12.github.io/asset-diff-report/)

## 功能

对比两个时期的资产明细数据，生成交互式可视化差异分析报告：

- **KPI 仪表盘** — 资产总量、经营面积、出租率等核心指标对比
- **桑基图** — 使用状态流向可视化，支持整体/变动流向切换
- **变动概览** — 新增 / 删除 / 状态变更 / 面积调整 统计
- **明细表** — 使用状态与建筑类型变动明细，支持搜索、分页、排序

## 技术栈

- Python (pandas, openpyxl) — 数据处理
- Jinja2 — 模板渲染
- Vue 3 + Plotly.js — 前端交互与可视化
- GitHub Pages — 在线部署

## 使用

```bash
pip install pandas openpyxl jinja2
python -m asset_report -c 当前期.xlsx -p 对比期.xlsx
```

| 参数      | 说明                                     |
| --------- | ---------------------------------------- |
| `-c`      | 当前期 Excel 文件                        |
| `-p`      | 对比期 Excel 文件                        |
| `-o`      | 输出 HTML 路径（可选）                   |
| `--scope` | 分析范围：`自有资产`（默认）/ `全部资产` |

## 实现逻辑

1. 读取两期 Excel，标准化字段
2. 逐条匹配资产编码，识别新增/删除/变更
3. 汇总 KPI、使用状态分布、建筑类型分布
4. 构建桑基图节点与连线数据
5. Jinja2 模板渲染为含 Vue + Plotly 的交互式页面
