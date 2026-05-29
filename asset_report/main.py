#!/usr/bin/env python3
"""资产差异分析可视化报告 — CLI 入口"""

import argparse
import sys
from pathlib import Path

# 确保能导入同目录下的 asset_diff.py
sys.path.insert(0, str(Path(__file__).parent.parent))

from asset_diff import (
    read_asset_file,
    map_status_category,
    split_self_owned,
    standardize_building_type,
    calc_summary_stats,
    compare_assets_by_status,
    compare_assets_by_building_type,
    analyze_category_changes,
)
from asset_report.sankey import build_sankey_data
from asset_report.html_builder import generate_vue_html


def main():
    parser = argparse.ArgumentParser(
        description="资产差异分析可视化报告",
        epilog="使用示例:\n  python -m asset_report -c 当前期.xlsx -p 对比期.xlsx"
    )
    parser.add_argument("-c", "--current", required=True, help="当前期 Excel 文件")
    parser.add_argument("-p", "--previous", required=True, help="对比期 Excel 文件")
    parser.add_argument("-o", "--output", default=None, help="输出 HTML 路径")
    parser.add_argument("--scope", default="自有资产", choices=["自有资产", "全部资产"])
    args = parser.parse_args()

    if args.output is None:
        nc = Path(args.current).stem
        np_ = Path(args.previous).stem
        args.output = f"{nc}_vs_{np_}_report.html"

    # 1. 读取
    print("\n[1/5] 读取数据...")
    df_c_raw, date_c = read_asset_file(args.current)
    df_p_raw, date_p = read_asset_file(args.previous)

    # 2. 预处理
    print("\n[2/5] 数据预处理...")
    for label, df in [("当前期", df_c_raw), ("对比期", df_p_raw)]:
        map_status_category(df)
        standardize_building_type(df)

    # 3. 拆分
    print("\n[3/5] 拆分自有/非自有...")
    if args.scope == "自有资产":
        df_c, _ = split_self_owned(df_c_raw)
        df_p, _ = split_self_owned(df_p_raw)
    else:
        df_c, df_p = df_c_raw, df_p_raw
    print(f"  当前期: {len(df_c)} 条, 对比期: {len(df_p)} 条")

    # 4. 统计 + 比对
    print("\n[4/5] 计算统计指标 & 比对...")
    stats_c = calc_summary_stats(df_c, "当前期")
    stats_p = calc_summary_stats(df_p, "对比期")
    changes_status_df = compare_assets_by_status(df_c, df_p)
    changes_bt_df = compare_assets_by_building_type(df_c, df_p)
    cat_changes_df = analyze_category_changes(df_c, df_p)
    print(f"  使用状态变更: {len(changes_status_df)} 条")
    print(f"  建筑类型变更: {len(changes_bt_df)} 条")

    # 5. 桑基图 + HTML
    print("\n[5/5] 生成报告...")
    sankey_data = build_sankey_data(df_c, df_p)
    print(f"  桑基图: 整体 {len(sankey_data['all']['links'])} 条连线, 变动 {len(sankey_data['changed']['links'])} 条连线")

    html = generate_vue_html(
        date_c, date_p, stats_c, stats_p, cat_changes_df,
        changes_status_df, changes_bt_df, sankey_data,
        len(df_c), len(df_p), df_c, df_p,
    )

    with open(args.output, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n报告已保存: {args.output}")


if __name__ == "__main__":
    main()
