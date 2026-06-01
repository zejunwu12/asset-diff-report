# 古城公司资产差异分析报告

## 功能说明

对比两个时期的资产明细数据，生成交互式可视化差异分析报告。

- **KPI 仪表盘** — 资产总量、经营面积、出租率等核心指标对比
- **桑基图** — 使用状态流向可视化，支持整体流向与变动流向切换
- **变动概览** — 新增 / 删除 / 状态变更 / 面积调整 四类变动统计卡片
- **使用状态变动明细** — 汇总表 + 可搜索、分页、排序的明细表
- **建筑类型变动明细** — 汇总表 + 可搜索、分页、排序的明细表

## 技术栈

| 层级       | 技术                      |
| ---------- | ------------------------- |
| 数据处理   | Python (pandas, openpyxl) |
| 模板渲染   | Jinja2                    |
| 前端框架   | Vue 3 (CDN)               |
| 图表可视化 | Plotly.js (CDN)           |
| 在线部署   | GitHub Pages              |

## 本地生成报告

### 环境要求

- Python ≥ 3.9

### 安装依赖

```bash
pip install pandas openpyxl jinja2
```

### 运行命令

```bash
# 进入仓库根目录
cd gucheng-asset-diff-report

# 生成报告
python -m asset_report -c 当前期数据.xlsx -p 对比期数据.xlsx

# 指定输出路径
python -m asset_report -c 当前期数据.xlsx -p 对比期数据.xlsx -o 输出报告.html
```

### 参数说明

| 参数         | 简写 | 必填 | 说明                                     |
| ------------ | ---- | ---- | ---------------------------------------- |
| `--current`  | `-c` | ✅    | 当前期 Excel 文件路径                    |
| `--previous` | `-p` | ✅    | 对比期 Excel 文件路径                    |
| `--output`   | `-o` | ❌    | 输出 HTML 路径（默认自动生成文件名）     |
| `--scope`    | —    | ❌    | 分析范围：`自有资产`（默认）/ `全部资产` |

## 项目结构

```
gucheng-asset-diff-report/
├── index.html                  # 已生成的报告页面（GitHub Pages 入口）
├── asset_diff.py               # 数据读取与差异比对模块
├── README.md                   # 本文件
├── .gitignore                  # Git 忽略规则
│
├── asset_report/               # Python 包（报告生成）
│   ├── __init__.py             # 包标识
│   ├── __main__.py             # 包入口（python -m asset_report）
│   ├── main.py                 # CLI 主逻辑
│   ├── html_builder.py         # HTML 报告构建
│   ├── sankey.py               # 桑基图数据生成
│   ├── utils.py                # 格式化与配色工具
│   └── templates/              # HTML 模板
│       ├── report.html
│       └── report_vue.html
│
└── data/                       # 数据文件
    ├── 20260331.xlsx
    └── 20250131.xlsx
```

## License

私有项目，仅供内部使用。
