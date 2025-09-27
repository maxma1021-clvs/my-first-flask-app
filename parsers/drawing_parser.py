import xml.etree.ElementTree as ET
from typing import Dict, Any, Optional
from parsers.utils import NS, half_point_to_pt


class DrawingParser:

    @staticmethod
    def parse(root: ET.Element, part_name: str, out: Dict[str, Any]) -> None:
        """
        掃描 w:drawing / v:shape / w14:textEffect 等元素，解析 WordArt 與文字方塊
        """
        # --- 新版 WordArt (<w14:textEffect>) ---
        for te in root.findall(".//w14:textEffect", NS):
            wordart_obj = DrawingParser._parse_text_effect(te, part_name)
            if wordart_obj:
                out["WordArt"].append(wordart_obj)

        # --- DrawingML 文字方塊 / WordArt (<w:drawing>) ---
        for shape in root.findall(".//w:drawing", NS):
            wordart_obj = DrawingParser._parse_w_drawing(shape, part_name)
            if wordart_obj:
                out["WordArt"].append(wordart_obj)

        # --- 舊版 VML WordArt (<v:shape><v:textpath>) ---
        for vshape in root.findall(".//v:shape", NS):
            wordart_obj = DrawingParser._parse_vml_wordart(vshape, part_name)
            if wordart_obj:
                out["WordArt"].append(wordart_obj)

    # -------------------
    # 子解析器
    # -------------------
    @staticmethod
    def _parse_w_drawing(shape: ET.Element, part_name: str) -> Optional[Dict[str, Any]]:
        """
        解析 <w:drawing> → 可能包含文字方塊 (<wps:txbx>) 或 WordArt
        """
        txbx = shape.find(".//wps:txbx/w:txbxContent", NS)
        if txbx is not None:
            text = "".join([t.text for t in txbx.findall(".//w:t", NS) if t.text])
            if text.strip():
                return {
                    "來源": part_name,
                    "文字": text.strip(),
                    "字型": DrawingParser._extract_font(txbx),
                    "大小": DrawingParser._extract_size(txbx),
                    "顏色": DrawingParser._extract_color(txbx),
                    "類型": "TextBox",
                    "幾何": DrawingParser._extract_warp(shape),
                    "填色": DrawingParser._extract_fill(shape),
                    "輪廓": DrawingParser._extract_outline(shape),
                    "陰影": DrawingParser._extract_shadow(shape),
                    "Glow": DrawingParser._extract_glow(shape),
                    "Reflection": DrawingParser._extract_reflection(shape),
                }
        return None

    @staticmethod
    def _parse_vml_wordart(vshape: ET.Element, part_name: str) -> Optional[Dict[str, Any]]:
        """
        解析 VML WordArt (<v:shape><v:textpath>)
        """
        textpath = vshape.find(".//v:textpath", NS)
        if textpath is not None:
            text = textpath.attrib.get("string", "")
            style = textpath.attrib
            return {
                "來源": part_name,
                "文字": text,
                "字型": style.get("font-family"),
                "大小": DrawingParser._try_size(style.get("font-size")),
                "顏色": vshape.attrib.get("strokecolor") or vshape.attrib.get("fillcolor"),
                "類型": "VML-WordArt",
                "幾何": {"type": vshape.attrib.get("type")} if "type" in vshape.attrib else {},
            }
        return None

    @staticmethod
    def _parse_text_effect(te: ET.Element, part_name: str) -> Dict[str, Any]:
        """
        解析新版 WordArt (<w14:textEffect>)
        """
        val = te.attrib.get(f"{{{NS['w14']}}}val")
        return {
            "來源": part_name,
            "文字": None,  # w14:textEffect 本身沒有文字
            "字型": te.attrib.get("font"),
            "大小": te.attrib.get("size"),
            "顏色": te.attrib.get("fillColor") or te.attrib.get("color"),
            "類型": f"WordArt-Effect:{val}",
            "幾何": DrawingParser._extract_warp(te),
            "填色": DrawingParser._extract_fill(te),
            "輪廓": DrawingParser._extract_outline(te),
            "陰影": DrawingParser._extract_shadow(te),
            "Glow": DrawingParser._extract_glow(te),
            "Reflection": DrawingParser._extract_reflection(te),
        }

    # -------------------
    # 輔助方法
    # -------------------
    @staticmethod
    def _extract_font(node: ET.Element) -> Optional[str]:
        rFonts = node.find(".//w:rFonts", NS)
        if rFonts is not None:
            return rFonts.attrib.get(f"{{{NS['w']}}}ascii") or rFonts.attrib.get(f"{{{NS['w']}}}hAnsi")
        return None

    @staticmethod
    def _extract_size(node: ET.Element) -> Optional[float]:
        sz = node.find(".//w:sz", NS)
        if sz is not None:
            return half_point_to_pt(sz.attrib.get(f"{{{NS['w']}}}val"))
        return None

    @staticmethod
    def _extract_color(node: ET.Element) -> Optional[str]:
        color = node.find(".//w:color", NS)
        if color is not None and color.attrib.get(f"{{{NS['w']}}}val"):
            return "#" + color.attrib.get(f"{{{NS['w']}}}val")
        return None

    @staticmethod
    def _try_size(sz_str: str) -> Optional[float]:
        if not sz_str:
            return None
        try:
            if sz_str.endswith("pt"):
                return float(sz_str.replace("pt", ""))
            return float(sz_str)
        except Exception:
            return None

    # ---- 幾何與樣式 ----
    @staticmethod
    def _extract_warp(node: ET.Element) -> Dict[str, Any]:
        prst = node.find(".//a:prstTxWarp", NS)
        return {"prstTxWarp": prst.attrib.get("prst")} if prst is not None else {}

    @staticmethod
    def _extract_fill(node: ET.Element) -> Optional[str]:
        solid = node.find(".//a:solidFill/a:srgbClr", NS)
        if solid is not None and "val" in solid.attrib:
            return "#" + solid.attrib["val"]
        return None

    @staticmethod
    def _extract_outline(node: ET.Element) -> Optional[str]:
        line = node.find(".//a:ln/a:solidFill/a:srgbClr", NS)
        if line is not None and "val" in line.attrib:
            return "#" + line.attrib["val"]
        return None

    @staticmethod
    def _extract_shadow(node: ET.Element) -> Optional[Dict[str, Any]]:
        shadow = node.find(".//a:effectLst/a:outerShdw", NS)
        if shadow is not None:
            return {
                "blurRad": shadow.attrib.get("blurRad"),
                "dist": shadow.attrib.get("dist"),
                "dir": shadow.attrib.get("dir"),
            }
        return None

    @staticmethod
    def _extract_glow(node: ET.Element) -> Optional[Dict[str, Any]]:
        glow = node.find(".//a:effectLst/a:glow", NS)
        if glow is not None:
            clr = glow.find("a:srgbClr", NS)
            return {
                "半徑": glow.attrib.get("rad"),
                "顏色": "#" + clr.attrib["val"] if clr is not None and "val" in clr.attrib else None,
            }
        return None

    @staticmethod
    def _extract_reflection(node: ET.Element) -> Optional[Dict[str, Any]]:
        refl = node.find(".//a:effectLst/a:reflection", NS)
        if refl is not None:
            return {k: v for k, v in refl.attrib.items()}
        return None
