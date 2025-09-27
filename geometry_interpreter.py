# geometry_interpreter.py
import re
from typing import Dict, Any, Optional

EMU_PER_CM = 360000
EMU_PER_PT = 12700


def emu_to_cm(emu: Optional[str]) -> Optional[float]:
    try:
        return round(int(emu) / EMU_PER_CM, 2) if emu else None
    except Exception:
        return None


def emu_to_pt(emu: Optional[str]) -> Optional[float]:
    try:
        return round(int(emu) / EMU_PER_PT, 2) if emu else None
    except Exception:
        return None


def parse_color(node) -> Optional[str]:
    """解析 a:srgbClr / a:schemeClr → Hex"""
    if node is None:
        return None
    srgb = node.attrib.get("val")
    if srgb:
        return f"#{srgb}"
    return None


def parse_outline(node) -> Dict[str, Any]:
    """解析 a:ln 框線"""
    result = {}
    if node is None:
        return result
    if "w" in node.attrib:
        result["粗細_pt"] = emu_to_pt(node.attrib.get("w"))
    color_node = node.find(".//a:srgbClr", {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"})
    if color_node is not None:
        result["顏色"] = parse_color(color_node)
    return result


def parse_geometry(shape_xml) -> Dict[str, Any]:
    """通用圖形幾何資訊 (寬、高、位置、框線、定位點)"""
    geom = {}
    if shape_xml is None:
        return geom

    ns = {
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main"
    }

    # 尺寸
    extent = shape_xml.find(".//wp:extent", ns)
    if extent is not None:
        geom["寬_cm"] = emu_to_cm(extent.attrib.get("cx"))
        geom["高_cm"] = emu_to_cm(extent.attrib.get("cy"))

    # 定位點 (水平)
    posH = shape_xml.find(".//wp:positionH", ns)
    if posH is not None:
        geom["水平位置"] = {
            "依據": posH.attrib.get("relativeFrom"),
            "偏移_cm": None
        }
        offset = posH.findtext(".//wp:posOffset", default=None, namespaces=ns)
        if offset:
            geom["水平位置"]["偏移_cm"] = emu_to_cm(offset)

    # 定位點 (垂直)
    posV = shape_xml.find(".//wp:positionV", ns)
    if posV is not None:
        geom["垂直位置"] = {
            "依據": posV.attrib.get("relativeFrom"),
            "偏移_cm": None
        }
        offset = posV.findtext(".//wp:posOffset", default=None, namespaces=ns)
        if offset:
            geom["垂直位置"]["偏移_cm"] = emu_to_cm(offset)

    # 框線
    ln = shape_xml.find(".//a:ln", ns)
    if ln is not None:
        geom["框線"] = parse_outline(ln)

    return geom
