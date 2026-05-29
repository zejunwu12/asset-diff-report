"""HTML 报告生成模块"""

import json
from pathlib import Path
from jinja2 import Environment, FileSystemLoader

from asset_diff import (
    calc_summary_stats,
    compare_assets_by_status,
    compare_assets_by_building_type,
    analyze_category_changes,
    calc_building_type_summary,
)
from asset_report.sankey import build_sankey_data
from asset_report.utils import fmt, diff_color, diff_arrow, badge_cls, PALETTE


def _build_kpis(items, stats_c, stats_p):
    """构建 KPI 数据列表"""
    result = []
    for stats_key, unit, display_label in items:
        vc = stats_c.get(stats_key, 0)
        vp = stats_p.get(stats_key, 0)
        diff = round(vc - vp, 2)
        if "出租率" in stats_key:
            rc = vc / 100 if vc > 1 else vc
            rp = vp / 100 if vp > 1 else vp
            result.append({
                "label": display_label, "unit": "",
                "current": f"{rc:.2%}", "previous": f"{rp:.2%}",
                "diff": f"{diff:+.2f}pp", "diff_raw": diff
            })
        else:
            result.append({
                "label": display_label, "unit": unit,
                "current": fmt(vc), "previous": fmt(vp),
                "diff": f"{diff:+,.2f}", "diff_raw": diff
            })
    return result


def _kpi_card_html(k, accent_color="#1a73e8"):
    """生成单个 KPI 卡片 HTML"""
    dv = k["diff_raw"]
    color = diff_color(dv)
    arrow = diff_arrow(dv)
    return (
        f'<div class="kpi-card" style="border-left-color:{accent_color}">'
        f'<div class="kpi-label">{k["label"]}</div>'
        f'<div class="kpi-values">'
        f'<span class="kpi-current">{k["current"]}</span>'
        f'<span class="kpi-diff" style="color:{color}">{arrow} {k["diff"]}</span>'
        f'</div>'
        f'<div class="kpi-previous">对比期: {k["previous"]} {k["unit"]}</div>'
        f'</div>'
    )


def _detail_to_list(df):
    """将 DataFrame 行转为字典列表"""
    rows = []
    for _, r in df.iterrows():
        rows.append({
            "code": r.get("资产编码", ""),
            "name": str(r.get("资产分层（同系统名称）", ""))[:30],
            "group": str(r.get("管理小组", ""))[:15],
            "change_type": r.get("变更类型", ""),
            "area_c": round(r.get("建筑面积（当前期）", 0), 2),
            "area_p": round(r.get("建筑面积（对比期）", 0), 2),
            "area_diff": round(r.get("建筑面积（变动）", 0), 2),
            "status_c": r.get("使用状态（当前期）", "") or r.get("建筑类型（当前期）", ""),
            "status_p": r.get("使用状态（对比期）", "") or r.get("建筑类型（对比期）", ""),
            "detail": str(r.get("变更详情", ""))[:50],
            "category_c": r.get("使用细分类（当前期）", "") or r.get("建筑类型（当前期）", ""),
            "category_p": r.get("使用细分类（对比期）", "") or r.get("建筑类型（对比期）", ""),
        })
    return rows


def _status_table_rows_html(cat_changes_df):
    """生成使用状态汇总表行 HTML"""
    rows = ""
    for _, r in cat_changes_df.iterrows():
        dc = round(r.get("变动面积", 0), 2)
        color = diff_color(dc)
        rows += (
            f'<tr data-name="{r.get("使用状态", "")}">'
            f'<td>{r.get("使用状态", "")}</td>'
            f'<td>{int(r.get("当前期宗数", 0))}</td>'
            f'<td>{int(r.get("对比期宗数", 0))}</td>'
            f'<td>{fmt(round(r.get("当前期面积", 0), 2))}</td>'
            f'<td>{fmt(round(r.get("对比期面积", 0), 2))}</td>'
            f'<td style="color:{color}">{fmt(dc)}</td>'
            f'<td>{int(r.get("新增宗数", 0))}</td>'
            f'<td>{fmt(round(r.get("新增面积", 0), 2))}</td>'
            f'<td>{int(r.get("删除宗数", 0))}</td>'
            f'<td>{fmt(round(r.get("删除面积", 0), 2))}</td>'
            f'<td>{int(r.get("面积调整宗数", 0))}</td>'
            f'<td>{fmt(round(r.get("面积调整额", 0), 2))}</td>'
            f'</tr>'
        )
    return rows


def _bt_table_rows_html(bt_summary):
    """生成建筑类型汇总表行 HTML"""
    rows = ""
    for item in bt_summary:
        color = diff_color(item["area_diff"])
        rows += (
            f'<tr>'
            f'<td>{item["name"]}</td>'
            f'<td>{item["count_c"]}</td><td>{item["count_p"]}</td>'
            f'<td>{fmt(item["area_c"])}</td><td>{fmt(item["area_p"])}</td>'
            f'<td style="color:{color}">{fmt(item["area_diff"])}</td>'
            f'<td>{item["added_count"]}</td><td>{fmt(item["added_area"])}</td>'
            f'<td>{item["removed_count"]}</td><td>{fmt(item["removed_area"])}</td>'
            f'<td>{item["adj_count"]}</td><td>{fmt(item["adj_area"])}</td>'
            f'</tr>'
        )
    return rows


def _detail_table_html(rows, title, count, table_id, data_attrs=None):
    """生成明细表 HTML"""
    if data_attrs is None:
        data_attrs = ("data-sp", "data-sc")
    da_prev, da_curr = data_attrs
    trs = ""
    for r in rows:
        dc = r["area_diff"]
        color = diff_color(dc)
        sp_val = r.get("category_p") or r['status_p']
        sc_val = r.get("category_c") or r['status_c']
        if r['change_type'] == "新增":
            sp_val = "新增"
        elif r['change_type'] == "删除":
            sc_val = "删除"
        trs += (
            f'<tr {da_prev}="{sp_val}" {da_curr}="{sc_val}" data-ct="{r["change_type"]}">'
            f'<td>{r["code"]}</td><td title="{r["name"]}">{r["name"]}</td>'
            f'<td>{r["group"]}</td>'
            f'<td><span class="badge badge-{badge_cls(r["change_type"])}">{r["change_type"]}</span></td>'
            f'<td>{fmt(r["area_c"])}</td><td>{fmt(r["area_p"])}</td>'
            f'<td style="color:{color}">{fmt(dc)}</td>'
            f'<td>{r["status_c"] or "-"}</td><td>{r["status_p"] or "-"}</td>'
            f'<td>{r["detail"]}</td>'
            f'</tr>'
        )
    return (
        f'<div class="card" id="card-{table_id}">'
        f'<h2>{title} <span>共 <b id="count-{table_id}">{count}</b> 条</span></h2>'
        f'<div class="table-toolbar">'
        f'<input type="text" class="search-input" id="search-{table_id}" placeholder="搜索资产编码/名称/状态..." />'
        f'<button class="btn-clear" id="clear-{table_id}" style="display:none;">✕ 清除筛选</button>'
        f'<div class="pagination" id="pager-{table_id}">'
        f'<button class="page-btn" id="prev-{table_id}">‹ 上一页</button>'
        f'<span class="page-info" id="info-{table_id}"></span>'
        f'<button class="page-btn" id="next-{table_id}">下一页 ›</button>'
        f'<select class="page-size-select" id="size-{table_id}">'
        f'<option value="5">5条/页</option>'
        f'<option value="10" selected>10条/页</option>'
        f'<option value="20">20条/页</option>'
        f'<option value="0">全部</option>'
        f'</select></div></div>'
        f'<div class="table-wrap"><table class="detail-table" id="table-{table_id}">'
        f'<thead><tr>'
        f'<th class="sortable" data-col="0">资产编码</th><th class="sortable" data-col="1">资产名称</th>'
        f'<th class="sortable" data-col="2">管理小组</th><th class="sortable" data-col="3">变更类型</th>'
        f'<th class="sortable" data-col="4">面积(当)</th><th class="sortable" data-col="5">面积(对)</th>'
        f'<th class="sortable" data-col="6">面积变动</th>'
        f'<th class="sortable" data-col="7">状态(当)</th><th class="sortable" data-col="8">状态(对)</th>'
        f'<th>变更详情</th>'
        f'</tr></thead>'
        f'<tbody>{trs}</tbody>'
        f'</table></div></div>'
    )


def _build_legend_html(sankey_data):
    """生成桑基图图例 HTML"""
    all_status_names = []
    seen = set()
    for n in sankey_data["all"]["nodes"]:
        name = n["name"]
        if name in ("新增", "删除") or name in seen:
            continue
        seen.add(name)
        all_status_names.append(name)
    all_status_names.sort()

    items = (
        '<div class="legend-item legend-selectable selected" data-name="__new__">'
        '<div class="legend-color" style="background:#9CA3AF;border:2px dashed #E65100;"></div>'
        '<span class="legend-label">新增</span></div>\n'
        '<div class="legend-item legend-selectable selected" data-name="__del__">'
        '<div class="legend-color" style="background:#9CA3AF;border:2px dashed #E65100;"></div>'
        '<span class="legend-label">删除</span></div>\n'
    )
    for i, name in enumerate(all_status_names):
        color = PALETTE[i % len(PALETTE)]
        items += (
            f'<div class="legend-item legend-selectable selected" data-name="{name}">'
            f'<div class="legend-color" style="background:{color};"></div>'
            f'<span class="legend-label">{name}</span></div>\n'
        )
    return items


def generate_html(date_c, date_p, stats_c, stats_p, cat_changes_df,
                  changes_status_df, changes_bt_df, sankey_data,
                  total_c, total_p, df_current=None, df_previous=None):
    """生成综合 HTML 报告"""

    # KPI 数据
    overview_kpis = [
        ("房产总面积", "㎡", "房产总面积"),
        ("资产总宗数", "宗", "资产总宗数"),
        ("在约合同数", "份", "在约合同数"),
        ("可出租面积", "㎡", "可出租面积"),
        ("已出租面积", "㎡", "已出租面积"),
        ("空置面积", "㎡", "空置面积"),
        ("出租率", "%", "出租率（口径3）"),
    ]
    structure_kpis = [
        ("可开发运作面积", "㎡", "可开发运作面积"),
        ("暂不可开发利用面积", "㎡", "暂不可开发利用面积"),
        ("I类面积", "㎡", "I类面积"),
        ("II类面积", "㎡", "II类面积"),
        ("可自主决策经营面积", "㎡", "可自主决策经营面积"),
        ("无经营决策权面积", "㎡", "无经营决策权面积"),
    ]

    # 变动计数
    changes_status_df_valid = changes_status_df[changes_status_df["变更类型"] != "面积调整"]
    added_count = len(changes_status_df_valid[changes_status_df_valid["变更类型"] == "新增"])
    removed_count = len(changes_status_df_valid[changes_status_df_valid["变更类型"] == "删除"])
    status_changed_count = len(changes_status_df_valid[changes_status_df_valid["变更类型"] == "状态变更"])
    area_adj_count = len(changes_status_df[changes_status_df["变更类型"] == "面积调整"])

    # KPI HTML
    overview = _build_kpis(overview_kpis, stats_c, stats_p)
    structure = _build_kpis(structure_kpis, stats_c, stats_p)
    overview_html = "".join(_kpi_card_html(k) for k in overview)
    structure_html = "".join(_kpi_card_html(k, "#5B6EB5") for k in structure)

    # 汇总表 HTML
    status_table_rows = _status_table_rows_html(cat_changes_df)
    bt_summary = calc_building_type_summary(df_current, df_previous) if df_current is not None and df_previous is not None else []
    bt_table_rows = _bt_table_rows_html(bt_summary)

    # 明细表 HTML
    status_detail = _detail_to_list(changes_status_df)
    bt_detail = _detail_to_list(changes_bt_df)
    status_detail_html = _detail_table_html(status_detail, "使用状态变动明细", len(status_detail), "status")
    bt_detail_html = _detail_table_html(bt_detail, "建筑类型变动明细", len(bt_detail), "building", data_attrs=("data-bt", "data-bc"))

    # 图例 HTML
    legend_items = _build_legend_html(sankey_data)

    # 桑基图 JSON
    all_nodes = json.dumps(sankey_data["all"]["nodes"], ensure_ascii=False)
    all_links = json.dumps(sankey_data["all"]["links"], ensure_ascii=False)
    changed_nodes = json.dumps(sankey_data["changed"]["nodes"], ensure_ascii=False)
    changed_links = json.dumps(sankey_data["changed"]["links"], ensure_ascii=False)

    # 渲染模板
    template_dir = Path(__file__).parent / "templates"
    env = Environment(loader=FileSystemLoader(str(template_dir)), autoescape=False)
    template = env.get_template("report.html")
    return template.render(
        date_c=date_c, date_p=date_p,
        overview_html=overview_html, structure_html=structure_html,
        added_count=added_count, removed_count=removed_count,
        status_changed_count=status_changed_count, area_adj_count=area_adj_count,
        legend_items=legend_items,
        status_table_rows=status_table_rows, bt_table_rows=bt_table_rows,
        status_detail_html=status_detail_html, bt_detail_html=bt_detail_html,
        all_nodes=all_nodes, all_links=all_links,
        changed_nodes=changed_nodes, changed_links=changed_links,
        PALETTE_JSON=json.dumps(PALETTE, ensure_ascii=False),
    )


def _build_kpis_json(items, stats_c, stats_p):
    """构建 KPI 数据列表（JSON 友好格式，含 diffColor/arrow）"""
    result = []
    for stats_key, unit, display_label in items:
        vc = stats_c.get(stats_key, 0)
        vp = stats_p.get(stats_key, 0)
        diff = round(vc - vp, 2)
        if "出租率" in stats_key:
            rc = vc / 100 if vc > 1 else vc
            rp = vp / 100 if vp > 1 else vp
            result.append({
                "label": display_label, "unit": "",
                "current": f"{rc:.2%}", "previous": f"{rp:.2%}",
                "diff": f"{diff:+.2f}pp", "diff_raw": diff,
                "diffColor": diff_color(diff),
                "arrow": diff_arrow(diff),
            })
        else:
            result.append({
                "label": display_label, "unit": unit,
                "current": fmt(vc), "previous": fmt(vp),
                "diff": f"{diff:+,.2f}", "diff_raw": diff,
                "diffColor": diff_color(diff),
                "arrow": diff_arrow(diff),
            })
    return result


def _detail_to_list_json(df):
    """将 DataFrame 行转为字典列表（JSON 友好格式，用于 Vue 模板）"""
    rows = []
    for _, r in df.iterrows():
        change_type = r.get("变更类型", "")
        area_diff = round(r.get("建筑面积（变动）", 0), 2)
        status_c = r.get("使用状态（当前期）", "") or r.get("建筑类型（当前期）", "")
        status_p = r.get("使用状态（对比期）", "") or r.get("建筑类型（对比期）", "")
        category_c = r.get("使用细分类（当前期）", "") or r.get("建筑类型（当前期）", "")
        category_p = r.get("使用细分类（对比期）", "") or r.get("建筑类型（对比期）", "")

        # sp / sc 计算
        sp_val = category_p or status_p
        sc_val = category_c or status_c
        if change_type == "新增":
            sp_val = "新增"
        elif change_type == "删除":
            sc_val = "删除"

        rows.append({
            "code": r.get("资产编码", ""),
            "name": str(r.get("资产分层（同系统名称）", ""))[:30],
            "group": str(r.get("管理小组", ""))[:15],
            "changeType": change_type,
            "badgeCls": badge_cls(change_type),
            "areaC": fmt(round(r.get("建筑面积（当前期）", 0), 2)),
            "areaP": fmt(round(r.get("建筑面积（对比期）", 0), 2)),
            "areaDiff": fmt(area_diff),
            "diffColor": diff_color(area_diff),
            "statusC": status_c,
            "statusP": status_p,
            "sp": sp_val,
            "sc": sc_val,
            "detail": str(r.get("变更详情", ""))[:50],
        })
    return rows


def _status_summary_to_json(cat_changes_df):
    """使用状态汇总表 → 二维数组 JSON"""
    rows = []
    for _, r in cat_changes_df.iterrows():
        dc = round(r.get("变动面积", 0), 2)
        rows.append([
            r.get("使用状态", ""),
            int(r.get("当前期宗数", 0)),
            int(r.get("对比期宗数", 0)),
            fmt(round(r.get("当前期面积", 0), 2)),
            fmt(round(r.get("对比期面积", 0), 2)),
            fmt(dc),
            int(r.get("新增宗数", 0)),
            fmt(round(r.get("新增面积", 0), 2)),
            int(r.get("删除宗数", 0)),
            fmt(round(r.get("删除面积", 0), 2)),
            int(r.get("面积调整宗数", 0)),
            fmt(round(r.get("面积调整额", 0), 2)),
        ])
    return rows


def _building_summary_to_json(bt_summary):
    """建筑类型汇总表 → 二维数组 JSON"""
    rows = []
    for item in bt_summary:
        rows.append([
            item["name"],
            item["count_c"], item["count_p"],
            fmt(item["area_c"]), fmt(item["area_p"]),
            fmt(item["area_diff"]),
            item["added_count"], fmt(item["added_area"]),
            item["removed_count"], fmt(item["removed_area"]),
            item["adj_count"], fmt(item["adj_area"]),
        ])
    return rows


def generate_vue_html(date_c, date_p, stats_c, stats_p, cat_changes_df,
                      changes_status_df, changes_bt_df, sankey_data,
                      total_c, total_p, df_current=None, df_previous=None):
    """生成 Vue 3 CDN 版本的 HTML 报告"""

    # KPI 定义
    overview_kpis_def = [
        ("房产总面积", "㎡", "房产总面积"),
        ("资产总宗数", "宗", "资产总宗数"),
        ("在约合同数", "份", "在约合同数"),
        ("可出租面积", "㎡", "可出租面积"),
        ("已出租面积", "㎡", "已出租面积"),
        ("空置面积", "㎡", "空置面积"),
        ("出租率", "%", "出租率（口径3）"),
    ]
    structure_kpis_def = [
        ("可开发运作面积", "㎡", "可开发运作面积"),
        ("暂不可开发利用面积", "㎡", "暂不可开发利用面积"),
        ("I类面积", "㎡", "I类面积"),
        ("II类面积", "㎡", "II类面积"),
        ("可自主决策经营面积", "㎡", "可自主决策经营面积"),
        ("无经营决策权面积", "㎡", "无经营决策权面积"),
    ]

    # 变动计数
    changes_status_df_valid = changes_status_df[changes_status_df["变更类型"] != "面积调整"]
    added_count = len(changes_status_df_valid[changes_status_df_valid["变更类型"] == "新增"])
    removed_count = len(changes_status_df_valid[changes_status_df_valid["变更类型"] == "删除"])
    status_changed_count = len(changes_status_df_valid[changes_status_df_valid["变更类型"] == "状态变更"])
    area_adj_count = len(changes_status_df[changes_status_df["变更类型"] == "面积调整"])

    # KPI JSON
    overview_kpis_json = json.dumps(
        _build_kpis_json(overview_kpis_def, stats_c, stats_p), ensure_ascii=False
    )
    structure_kpis_json = json.dumps(
        _build_kpis_json(structure_kpis_def, stats_c, stats_p), ensure_ascii=False
    )

    # 汇总表 JSON
    status_summary_json = json.dumps(
        _status_summary_to_json(cat_changes_df), ensure_ascii=False
    )
    bt_summary = calc_building_type_summary(df_current, df_previous) if df_current is not None and df_previous is not None else []
    building_summary_json = json.dumps(
        _building_summary_to_json(bt_summary), ensure_ascii=False
    )

    # 明细表 JSON
    status_detail_json = json.dumps(
        _detail_to_list_json(changes_status_df), ensure_ascii=False
    )
    building_detail_json = json.dumps(
        _detail_to_list_json(changes_bt_df), ensure_ascii=False
    )

    # 桑基图 JSON
    all_nodes = json.dumps(sankey_data["all"]["nodes"], ensure_ascii=False)
    all_links = json.dumps(sankey_data["all"]["links"], ensure_ascii=False)
    changed_nodes = json.dumps(sankey_data["changed"]["nodes"], ensure_ascii=False)
    changed_links = json.dumps(sankey_data["changed"]["links"], ensure_ascii=False)

    # 渲染模板
    template_dir = Path(__file__).parent / "templates"
    env = Environment(loader=FileSystemLoader(str(template_dir)), autoescape=False)
    template = env.get_template("report_vue.html")
    return template.render(
        date_c=date_c, date_p=date_p,
        overview_kpis_json=overview_kpis_json,
        structure_kpis_json=structure_kpis_json,
        added_count=added_count, removed_count=removed_count,
        status_changed_count=status_changed_count, area_adj_count=area_adj_count,
        status_summary_json=status_summary_json,
        building_summary_json=building_summary_json,
        status_detail_json=status_detail_json,
        building_detail_json=building_detail_json,
        all_nodes=all_nodes, all_links=all_links,
        changed_nodes=changed_nodes, changed_links=changed_links,
    )
