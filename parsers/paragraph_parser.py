# parsers/paragraph_parser.py
import xml.etree.ElementTree as ET
from typing import Dict, Any, List
from .utils import NS, half_point_to_pt


# -------- run style --------
def parse_run_props(r: ET.Element) -> Dict[str, Any]:
    rPr = r.find("w:rPr", NS)
    if rPr is None:
        return {}
    out: Dict[str, Any] = {}

    # 字元樣式名稱 (rStyle)
    rStyle = rPr.find("w:rStyle", NS)
    if rStyle is not None:
        style_name = rStyle.attrib.get(f"{{{NS['w']}}}val")
        if style_name:
            out["styleName"] = style_name

    # 字型
    rFonts = rPr.find("w:rFonts", NS)
    if rFonts is not None:
        out["fonts"] = {
            "ascii": rFonts.attrib.get(f"{{{NS['w']}}}ascii"),
            "eastAsia": rFonts.attrib.get(f"{{{NS['w']}}}eastAsia"),
            "cs": rFonts.attrib.get(f"{{{NS['w']}}}cs"),
            "hAnsi": rFonts.attrib.get(f"{{{NS['w']}}}hAnsi"),
        }

    # 字號
    sz = rPr.find("w:sz", NS)
    if sz is not None:
        out["size_pt"] = half_point_to_pt(sz.attrib.get(f"{{{NS['w']}}}val"))

    # 粗體、斜體
    if rPr.find("w:b", NS) is not None:
        out["bold"] = True
    if rPr.find("w:i", NS) is not None:
        out["italic"] = True

    # 文字顏色
    color = rPr.find("w:color", NS)
    if color is not None and color.attrib.get(f"{{{NS['w']}}}val"):
        out["color"] = "#" + color.attrib.get(f"{{{NS['w']}}}val")

    # ➤ 字距 (spacing, 單位 1/20 pt → cm)
    spacing = rPr.find("w:spacing", NS)
    if spacing is not None and spacing.attrib.get(f"{{{NS['w']}}}val"):
        try:
            twips = int(spacing.attrib.get(f"{{{NS['w']}}}val"))  # 1/20 pt
            pt = twips / 20.0
            out["spacing_cm"] = round(pt * 0.0352778, 3)
        except Exception:
            pass

    # ➤ 底紋 (背景顏色)
    shd = rPr.find("w:shd", NS)
    if shd is not None and shd.attrib.get(f"{{{NS['w']}}}fill"):
        out["background_color"] = "#" + shd.attrib.get(f"{{{NS['w']}}}fill")

    # 底線
    underline = rPr.find("w:u", NS)
    if underline is not None:
        out["underline"] = True

    # 刪除線
    strike = rPr.find("w:strike", NS)
    if strike is not None:
        out["strike"] = True

    return out


# -------- 合併相鄰 run --------
def merge_runs(runs: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    合併相鄰且樣式相同的 runs
    """
    if not runs:
        return []

    merged: List[Dict[str, Any]] = []

    def styles_equal(style1: Dict, style2: Dict) -> bool:
        """比較兩個樣式是否相同"""
        if not style1 and not style2:
            return True
        if not style1 or not style2:
            return False

        # 比較基本屬性 (包含新的 styleName)
        basic_attrs = ['styleName', 'bold', 'italic', 'underline', 'strike', 'size_pt', 'color', 'background_color',
                       'spacing_cm']
        for attr in basic_attrs:
            if style1.get(attr) != style2.get(attr):
                return False

        # 比較字型
        fonts1 = style1.get('fonts', {})
        fonts2 = style2.get('fonts', {})
        if fonts1 != fonts2:
            return False

        return True

    for r in runs:
        if not r.get("text"):
            continue

        if merged and styles_equal(r["style"], merged[-1]["style"]):
            merged[-1]["text"] += r["text"]
        else:
            merged.append(r)

    return merged


# -------- 段落屬性 --------
def parse_paragraph_props(p: ET.Element) -> Dict[str, Any]:
    pPr = p.find("w:pPr", NS)
    if pPr is None:
        return {}
    out: Dict[str, Any] = {}

    # 段落樣式名稱 (pStyle)
    pStyle = pPr.find("w:pStyle", NS)
    if pStyle is not None:
        style_name = pStyle.attrib.get(f"{{{NS['w']}}}val")
        if style_name:
            out["styleName"] = style_name

    # 對齊
    jc = pPr.find("w:jc", NS)
    if jc is not None:
        out["align"] = jc.attrib.get(f"{{{NS['w']}}}val")

    # 行距 - 轉換單位
    sp = pPr.find("w:spacing", NS)
    if sp is not None:
        spacing_info = {}
        for k, v in sp.attrib.items():
            attr_name = k.split('}')[1] if '}' in k else k
            spacing_info[attr_name] = v
            # 轉換 twips 到 cm (如果需要)
            if attr_name in ["line", "before", "after"]:
                try:
                    twips = int(v)
                    cm = round(twips / 20.0 * 0.0352778, 3)
                    spacing_info[f"{attr_name}_cm"] = cm
                except (ValueError, TypeError):
                    pass
        out["spacing"] = spacing_info

    # 縮排 - 轉換單位
    ind = pPr.find("w:ind", NS)
    if ind is not None:
        indent_info = {}
        # 獲取所有屬性
        for k, v in ind.attrib.items():
            attr_name = k.split('}')[1] if '}' in k else k
            indent_info[attr_name] = v
            # 轉換數值屬性到 cm
            if attr_name in ["left", "right", "firstLine", "hanging", "start", "end"]:
                try:
                    twips = int(v)
                    cm = round(twips / 20.0 * 0.0352778, 3)
                    indent_info[f"{attr_name}_cm"] = cm
                except (ValueError, TypeError):
                    pass  # 保持原始值
        out["indent"] = indent_info

    # ➤ 定位點 (tabs)
    tabs = []
    for tab in pPr.findall("w:tabs/w:tab", NS):
        pos_twips = float(tab.attrib.get(f"{{{NS['w']}}}pos", "0"))
        pos_pt = pos_twips / 20.0
        pos_cm = round(pos_pt * 0.0352778, 3)

        tab_info = {
            "位置_pt": pos_pt,
            "位置_cm": pos_cm,
            "對齊": tab.attrib.get(f"{{{NS['w']}}}val", "left"),
            "前置符號": tab.attrib.get(f"{{{NS['w']}}}leader", "none")
        }
        tabs.append(tab_info)
    if tabs:
        out["tabs"] = tabs

    # ➤ 列表編號 (numPr)
    numPr = pPr.find("w:numPr", NS)
    if numPr is not None:
        num_info = {}
        # 編號 ID
        numId = numPr.find("w:numId", NS)
        if numId is not None:
            num_info["numId"] = numId.attrib.get(f"{{{NS['w']}}}val")

        # 級別
        ilvl = numPr.find("w:ilvl", NS)
        if ilvl is not None:
            num_info["level"] = ilvl.attrib.get(f"{{{NS['w']}}}val")

        # 是否為列表項
        num_info["isListItem"] = True
        out["numbering"] = num_info

    return out


# -------- 段落 --------
def extract_paragraph_text(p: ET.Element) -> str:
    return "".join([t.text for t in p.findall(".//w:t", NS) if t.text]).strip()


def parse_paragraphs(root: ET.Element, part_name: str, out: Dict[str, Any]) -> None:
    for idx, p in enumerate(root.findall(".//w:p", NS), 1):

        # 👉 檢查此段是否包含 WordArt (避免當作一般段落)
        if p.find(".//w14:textEffect", NS) is not None \
                or p.find(".//wps:wsp//w14:textEffect", NS) is not None \
                or p.find(".//mc:AlternateContent//w14:textEffect", NS) is not None \
                or p.find(".//a:graphic/a:graphicData/a:txBody/a:bodyPr/a:prstTxWarp", NS) is not None \
                or p.find(".//v:textpath", NS) is not None:
            continue

        text = extract_paragraph_text(p)

        runs: List[Dict[str, Any]] = []
        for r in p.findall("./w:r", NS):
            t = r.find("w:t", NS)
            txt = t.text if t is not None else ""
            runs.append({"text": txt, "style": parse_run_props(r)})

        # 合併相同樣式的 runs
        runs = merge_runs(runs)

        props = parse_paragraph_props(p)

        out["段落"].append({
            "來源": part_name,
            "編號": idx,
            "文字": text,
            "runs": runs,
            "屬性": props
        })
