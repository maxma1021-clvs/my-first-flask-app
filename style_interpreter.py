# -*- coding: utf-8 -*-
# style_interpreter.py
from typing import Any, Dict, Optional, List

# ---- 小工具 ----
def _local(tag: str) -> str:
    return tag.split('}', 1)[-1] if '}' in tag else tag

def _find_first(node, localname: str):
    if node is None:
        return None
    for el in node.iter():
        if _local(el.tag) == localname:
            return el
    return None

def _find_all(node, localname: str) -> List[Any]:
    if node is None:
        return []
    return [el for el in node.iter() if _local(el.tag) == localname]

def emu_to_pt(val: Optional[str]) -> Optional[float]:
    try:
        return int(val) / 12700.0 if val is not None else None
    except Exception:
        return None

def emu_to_cm(val: Optional[str]) -> Optional[float]:
    try:
        return int(val) / 360000.0 if val is not None else None
    except Exception:
        return None

def dxa_to_pt(val: Optional[str]) -> Optional[float]:
    """twips (1/20pt) -> pt"""
    try:
        return int(val) / 20.0 if val is not None else None
    except Exception:
        return None

def perc_to_float(val: Optional[str]) -> Optional[float]:
    """OOXML 百分比，100000 = 100% -> 1.0"""
    try:
        return int(val) / 100000.0 if val is not None else None
    except Exception:
        return None

# ---- 顏色 ----
def parse_color_node(color_node) -> Optional[Dict[str, Any]]:
    """
    支援 a:srgbClr (val=RRGGBB) 及 a:schemeClr（含調整）
    回傳: {"type":"srgb","value":"#RRGGBB"} 或 {"type":"scheme","scheme":"accent1","mods":[...]}
    """
    if color_node is None:
        return None
    lname = _local(color_node.tag)
    if lname == "srgbClr":
        val = color_node.attrib.get("val")
        return {"type": "srgb", "value": f"#{val}"} if val else None
    if lname == "schemeClr":
        scheme = color_node.attrib.get("val")
        mods = []
        for ch in list(color_node):
            mods.append({"name": _local(ch.tag), "attrs": dict(ch.attrib)})
        return {"type": "scheme", "scheme": scheme, "mods": mods}
    # 其他類型回傳原 attrs
    return {"type": lname, "attrs": dict(color_node.attrib)}

# ---- 填色 ----
def interpret_fill(container_node) -> Optional[Dict[str, Any]]:
    if container_node is None:
        return None
    solid = _find_first(container_node, "solidFill")
    grad  = _find_first(container_node, "gradFill")
    patt  = _find_first(container_node, "pattFill")

    if solid is not None:
        col = _find_first(solid, "srgbClr") or _find_first(solid, "schemeClr")
        return {"類型": "單色", "顏色": parse_color_node(col)}

    if grad is not None:
        gs_list = _find_all(grad, "gs")
        stops = []
        for gs in gs_list:
            pos = gs.attrib.get("pos")
            col = _find_first(gs, "srgbClr") or _find_first(gs, "schemeClr")
            stops.append({"位置": perc_to_float(pos), "顏色": parse_color_node(col)})
        # 方向/型別
        lin = _find_first(grad, "lin")
        path = _find_first(grad, "path")
        extra = {}
        if lin is not None:
            extra["線性角度"] = lin.attrib.get("ang")
        if path is not None:
            extra["路徑類型"] = path.attrib.get("path")
        return {"類型": "漸層", "漸層點": stops, **extra}

    if patt is not None:
        fg = _find_first(patt, "fgClr")
        bg = _find_first(patt, "bgClr")
        fgcol = parse_color_node(_find_first(fg, "srgbClr") or _find_first(fg, "schemeClr")) if fg is not None else None
        bgcol = parse_color_node(_find_first(bg, "srgbClr") or _find_first(bg, "schemeClr")) if bg is not None else None
        return {"類型": "圖案", "樣式": patt.attrib.get("prst"), "前景": fgcol, "背景": bgcol}

    # 也可能直接傳進來就是 solid/grad/patt 節點
    lname = _local(container_node.tag)
    if lname in ("solidFill", "gradFill", "pattFill"):
        return interpret_fill({"dummy": "wrap", "children": [container_node]})  # 不太會走到

    return None

# ---- 外框 ----
def interpret_outline(ln_node) -> Optional[Dict[str, Any]]:
    if ln_node is None:
        return None
    w = emu_to_pt(ln_node.attrib.get("w"))
    cap = ln_node.attrib.get("cap")
    cmpd = ln_node.attrib.get("cmpd")
    algn = ln_node.attrib.get("algn")
    dash = _find_first(ln_node, "prstDash")
    col  = _find_first(ln_node, "srgbClr") or _find_first(ln_node, "schemeClr")
    return {
        "線寬_pt": w,
        "端點": cap,
        "線型": cmpd,
        "對齊": algn,
        "虛線": dash.attrib.get("val") if dash is not None else None,
        "顏色": parse_color_node(col)
    }

# ---- Glow / Reflection / Shadow / SoftEdge ----
def interpret_glow(glow_node) -> Optional[Dict[str, Any]]:
    if glow_node is None:
        return None
    col = _find_first(glow_node, "srgbClr") or _find_first(glow_node, "schemeClr")
    return {"半徑_pt": emu_to_pt(glow_node.attrib.get("rad")), "顏色": parse_color_node(col)}

def interpret_reflection(ref_node) -> Optional[Dict[str, Any]]:
    if ref_node is None:
        return None
    return {
        "距離_pt": emu_to_pt(ref_node.attrib.get("dist")),
        "模糊_pt": emu_to_pt(ref_node.attrib.get("blurRad")),
        "起始透明": perc_to_float(ref_node.attrib.get("stA")),
        "結束透明": perc_to_float(ref_node.attrib.get("endA")),
        "大小比例X": perc_to_float(ref_node.attrib.get("sx")),
        "大小比例Y": perc_to_float(ref_node.attrib.get("sy"))
    }

def interpret_shadow(shdw_node) -> Optional[Dict[str, Any]]:
    if shdw_node is None:
        return None
    col = _find_first(shdw_node, "srgbClr") or _find_first(shdw_node, "schemeClr")
    return {
        "外陰影": _local(shdw_node.tag) == "outerShdw",
        "距離_pt": emu_to_pt(shdw_node.attrib.get("dist")),
        "模糊_pt": emu_to_pt(shdw_node.attrib.get("blurRad")),
        "方向角度": shdw_node.attrib.get("dir"),
        "對齊": shdw_node.attrib.get("algn"),
        "顏色": parse_color_node(col)
    }

def interpret_softedge(se_node) -> Optional[Dict[str, Any]]:
    if se_node is None:
        return None
    return {"半徑_pt": emu_to_pt(se_node.attrib.get("rad"))}

# ---- 3D ----
def interpret_3d(sp3d_node) -> Optional[Dict[str, Any]]:
    if sp3d_node is None:
        return None
    bevel_t = _find_first(sp3d_node, "bevelT")
    bevel_b = _find_first(sp3d_node, "bevelB")
    extru   = _find_first(sp3d_node, "extrusionClr")
    cntr    = _find_first(sp3d_node, "contourClr")
    data = {
        "斜角上_pt": {
            "寬": emu_to_pt(bevel_t.attrib.get("w")) if bevel_t is not None else None,
            "高": emu_to_pt(bevel_t.attrib.get("h")) if bevel_t is not None else None,
        },
        "斜角下_pt": {
            "寬": emu_to_pt(bevel_b.attrib.get("w")) if bevel_b is not None else None,
            "高": emu_to_pt(bevel_b.attrib.get("h")) if bevel_b is not None else None,
        },
        "拉伸高度_pt": emu_to_pt(sp3d_node.attrib.get("extrusionH")),
        "輪廓寬度_pt": emu_to_pt(sp3d_node.attrib.get("contourW")),
        "拉伸顏色": parse_color_node(extru),
        "輪廓顏色": parse_color_node(cntr),
        "材質": sp3d_node.attrib.get("prstMaterial")
    }
    return data

# ---- 字型 (a:rPr / w:rPr) ----
def interpret_font(rpr_node) -> Optional[Dict[str, Any]]:
    if rpr_node is None:
        return None
    # DrawingML 的 a:rPr
    sz_raw = rpr_node.attrib.get("sz")
    font_pt = float(sz_raw) / 100.0 if sz_raw and sz_raw.isdigit() else None
    b = rpr_node.attrib.get("b")
    i = rpr_node.attrib.get("i")
    u = rpr_node.attrib.get("u")
    latin = _find_first(rpr_node, "latin")
    typeface = latin.attrib.get("typeface") if latin is not None else None
    # 色彩/外框
    fill = interpret_fill(rpr_node)
    ln = interpret_outline(_find_first(rpr_node, "ln"))
    return {
        "字型": typeface,
        "大小_pt": font_pt,
        "粗體": True if b == "1" or b == "true" else False,
        "斜體": True if i == "1" or i == "true" else False,
        "底線": u if u else None,
        "填色": fill,
        "外框": ln
    }

# ---- 文字 Warp ----
def interpret_warp(warp_node) -> Optional[Dict[str, Any]]:
    if warp_node is None:
        return None
    return {"樣式": warp_node.attrib.get("prst")}
