"""桑基图数据构建与绘制模块"""

import json
from pathlib import Path
from asset_diff import _safe_area


PALETTE = [
    "#3B6CB5", "#2E8B6E", "#C47A2B", "#B5456E", "#7B5EA7", "#2A8C8C",
    "#C4962B", "#C44B4B", "#5A8C3B", "#8B5E3B", "#5B6EB5", "#3BA5A5",
    "#A55B8B", "#6B8B2E", "#8B3B5E", "#2E6B8B",
]


def _add_flow(flows, key, area):
    if key not in flows:
        flows[key] = {"area": 0, "count": 0}
    flows[key]["area"] += area
    flows[key]["count"] += 1
    return flows[key]


def _build_sankey_from_flows(flows):
    flows = {k: v for k, v in flows.items() if v["area"] > 0}
    left_names, right_names = set(), set()
    for src, tgt in flows.keys():
        if src != "__new__":
            left_names.add(src)
        if tgt != "__del__":
            right_names.add(tgt)

    nodes, node_index = [], {}
    for n in sorted(left_names):
        node_index[f"L:{n}"] = len(nodes)
        nodes.append({"name": n, "side": "left"})
    if any(k[0] == "__new__" for k in flows):
        node_index["__new__"] = len(nodes)
        nodes.append({"name": "新增", "side": "left"})
    for n in sorted(right_names):
        node_index[f"R:{n}"] = len(nodes)
        nodes.append({"name": n, "side": "right"})
    if any(k[1] == "__del__" for k in flows):
        node_index["__del__"] = len(nodes)
        nodes.append({"name": "删除", "side": "right"})

    links = []
    for (src, tgt), v in sorted(flows.items(), key=lambda x: -x[1]["area"]):
        sk = "__new__" if src == "__new__" else f"L:{src}"
        tk = "__del__" if tgt == "__del__" else f"R:{tgt}"
        links.append({"source": node_index[sk], "target": node_index[tk],
                       "value": round(v["area"], 2), "count": v["count"]})

    connected = set()
    for l in links:
        connected.add(l["source"])
        connected.add(l["target"])
    old2new, new_nodes = {}, []
    for oi in sorted(connected):
        old2new[oi] = len(new_nodes)
        new_nodes.append(nodes[oi])
    for l in links:
        l["source"] = old2new[l["source"]]
        l["target"] = old2new[l["target"]]

    return {"nodes": [{"name": n["name"], "side": n["side"]} for n in new_nodes], "links": links}


def build_sankey_data(df_current, df_previous):
    """全量 + 变动 两个桑基图"""
    df_c = df_current[df_current["使用细分类"] != "未知"].copy()
    df_p = df_previous[df_previous["使用细分类"] != "未知"].copy()

    codes_c, codes_p = set(df_c["资产编码"]), set(df_p["资产编码"])
    added, removed, common = codes_c - codes_p, codes_p - codes_c, codes_c & codes_p

    flows_all, flows_changed = {}, {}

    for code in added:
        row = df_c[df_c["资产编码"] == code].iloc[0]
        area = _safe_area(row.get("建筑面积"))
        status = row.get("使用细分类", "") or "未知"
        if area > 0:
            _add_flow(flows_all, ("__new__", status), area)
            _add_flow(flows_changed, ("__new__", status), area)

    for code in removed:
        row = df_p[df_p["资产编码"] == code].iloc[0]
        area = _safe_area(row.get("建筑面积"))
        status = row.get("使用细分类", "") or "未知"
        if area > 0:
            _add_flow(flows_all, (status, "__del__"), area)
            _add_flow(flows_changed, (status, "__del__"), area)

    for code in common:
        rc = df_c[df_c["资产编码"] == code].iloc[0]
        rp = df_p[df_p["资产编码"] == code].iloc[0]
        ac, ap = _safe_area(rc.get("建筑面积")), _safe_area(rp.get("建筑面积"))
        sc, sp = rc.get("使用细分类", "") or "未知", rp.get("使用细分类", "") or "未知"
        if ac <= 0 and ap <= 0:
            if sp != sc:
                _add_flow(flows_all, (sp, sc), 0.01)
                _add_flow(flows_changed, (sp, sc), 0.01)
            continue
        area = max(ac, ap)
        _add_flow(flows_all, (sp, sc), area)
        if sp != sc:
            _add_flow(flows_changed, (sp, sc), area)

    return {"all": _build_sankey_from_flows(flows_all),
            "changed": _build_sankey_from_flows(flows_changed)}


def _get_color_map(nodes):
    """根据节点名称分配颜色"""
    names = sorted(set(n["name"] for n in nodes if n["name"] not in ("新增", "删除")))
    color_map = {}
    for i, name in enumerate(names):
        color_map[name] = PALETTE[i % len(PALETTE)]
    return color_map


def _hex_to_rgba(hex_color, alpha=0.25):
    """将 #RRGGBB 转为 rgba(r,g,b,a) 格式"""
    hex_color = hex_color.lstrip("#")
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


def plot_sankey(data, title="使用状态流向", mode="all", output=None, auto_open=False):
    """绘制桑基图并保存为独立 HTML 文件。

    Args:
        data: build_sankey_data() 返回的数据字典
        title: 图表标题
        mode: "all"（整体流向）或 "changed"（变动流向）
        output: 输出 HTML 文件路径，默认为 sankey_{mode}.html
        auto_open: 是否自动打开浏览器
    """
    try:
        import plotly
    except ImportError:
        raise ImportError("需要安装 plotly: pip install plotly")

    key = mode if mode in data else "all"
    sankey = data[key]
    nodes = sankey["nodes"]
    links = sankey["links"]

    color_map = _get_color_map(nodes)

    def get_color(name):
        if name in ("新增", "删除"):
            return "#9CA3AF"
        return color_map.get(name, "#999")

    node_colors = [get_color(n["name"]) for n in nodes]
    node_labels = [n["name"] for n in nodes]
    node_counts = [0] * len(nodes)
    for l in links:
        node_counts[l["source"]] += l["count"]
        node_counts[l["target"]] += l["count"]

    link_colors = []
    for l in links:
        sn = nodes[l["source"]]["name"]
        tn = nodes[l["target"]]["name"]
        if sn == "新增" or tn == "删除":
            link_colors.append("rgba(230,81,0,0.35)")
        else:
            link_colors.append(_hex_to_rgba(node_colors[l["target"]], 0.25))

    fig = plotly.graph_objects.Figure(plotly.graph_objects.Sankey(
        arrangement="fixed",
        node=dict(
            pad=20, thickness=24,
            line=dict(color="white", width=2),
            label=node_labels,
            color=node_colors,
            customdata=node_counts,
            hovertemplate="<b>%{label}</b><br>面积: %{value:,.2f} ㎡<br>宗数: %{customdata}<extra></extra>",
        ),
        link=dict(
            source=[l["source"] for l in links],
            target=[l["target"] for l in links],
            value=[l["value"] for l in links],
            color=link_colors,
            customdata=[l["count"] for l in links],
            hovertemplate="<b>%{source.label}</b> → <b>%{target.label}</b><br>面积: %{value:,.2f} ㎡<br>宗数: %{customdata}<extra></extra>",
        ),
    ))

    fig.update_layout(
        title=dict(text=title, font=dict(size=18)),
        font=dict(size=12, family="-apple-system,Microsoft YaHei,sans-serif"),
        margin=dict(l=20, r=20, t=60, b=20),
        height=700,
    )

    if output is None:
        output = f"sankey_{key}.html"

    plotly.io.write_html(fig, file=output, auto_open=auto_open)
    print(f"桑基图已保存: {output}")
    return output


if __name__ == "__main__":
    import argparse
    from asset_diff import read_asset_file, map_status_category, standardize_building_type, split_self_owned

    parser = argparse.ArgumentParser(description="独立绘制桑基图")
    parser.add_argument("-c", "--current", required=True, help="当前期 Excel")
    parser.add_argument("-p", "--previous", required=True, help="对比期 Excel")
    parser.add_argument("-o", "--output", default=None, help="输出文件名（默认 sankey_all.html）")
    parser.add_argument("--mode", default="all", choices=["all", "changed"], help="all=整体流向, changed=变动流向")
    parser.add_argument("--open", action="store_true", help="自动打开浏览器")
    args = parser.parse_args()

    print("读取数据...")
    df_c, date_c = read_asset_file(args.current)
    df_p, date_p = read_asset_file(args.previous)

    print("预处理...")
    for df in [df_c, df_p]:
        map_status_category(df)
        standardize_building_type(df)

    print("筛选自有资产...")
    df_c, _ = split_self_owned(df_c)
    df_p, _ = split_self_owned(df_p)

    print("构建桑基图数据...")
    data = build_sankey_data(df_c, df_p)

    title = "使用状态整体流向" if args.mode == "all" else "使用状态变动流向"
    plot_sankey(data, title=title, mode=args.mode, output=args.output, auto_open=args.open)
