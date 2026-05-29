"""通用工具函数"""

PALETTE = [
    "#3B6CB5", "#2E8B6E", "#C47A2B", "#B5456E", "#7B5EA7", "#2A8C8C",
    "#C4962B", "#C44B4B", "#5A8C3B", "#8B5E3B", "#5B6EB5", "#3BA5A5",
    "#A55B8B", "#6B8B2E", "#8B3B5E", "#2E6B8B",
]


def fmt(val, decimals=2):
    """格式化数字"""
    if isinstance(val, float):
        if abs(val) >= 10000:
            return f"{val:,.{decimals}f}"
        return f"{val:.{decimals}f}"
    return str(val)


def diff_color(val, threshold=0.01):
    """根据数值返回颜色：正数红色、负数绿色、零灰色"""
    if val > threshold:
        return "#c62828"
    if val < -threshold:
        return "#2e7d32"
    return "#333"


def diff_arrow(val, threshold=0.01):
    """根据数值返回箭头符号"""
    if val > threshold:
        return "↑"
    if val < -threshold:
        return "↓"
    return "→"


def badge_cls(change_type):
    """变更类型 → CSS class 映射"""
    return {"新增": "added", "删除": "removed", "状态变更": "changed",
            "面积调整": "area", "类型变更": "changed"}.get(change_type, "area")
