import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, Any, Optional, List

# ------- Namespaces -------
NS = {
    "w":   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp":  "http://schemas.openxmlformats.org/wordprocessingDrawing",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "v":   "urn:schemas-microsoft-com:vml",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",  # Office 2019+ WordArt
    "mc":  "http://schemas.openxmlformats.org/markup-compatibility/2006"
}

# ------- 單位轉換 -------
HALF_POINT_TO_PT = 0.5        # Word sz 單位 → pt
EMU_PER_CM = 360000.0         # EMU → cm
PT_TO_CM = 0.0352778          # pt → cm


def get_xml(z: zipfile.ZipFile, name: str) -> Optional[ET.Element]:
    """讀取並解析 XML 檔案"""
    try:
        return ET.fromstring(z.read(name))
    except Exception:
        return None


def half_point_to_pt(v: Optional[str]) -> Optional[float]:
    try:
        return round(float(v) * HALF_POINT_TO_PT, 2)
    except Exception:
        return None


def emu_to_cm(v: Optional[str]) -> Optional[float]:
    try:
        return round(int(v) / EMU_PER_CM, 3)
    except Exception:
        return None


# -------- Drawing 幾何 --------
def parse_wrapper_geom(wrapper: ET.Element) -> Dict[str, Any]:
    geom: Dict[str, Any] = {}
    ext = wrapper.find("wp:extent", NS)
    if ext is not None:
        geom["寬_cm"] = emu_to_cm(ext.attrib.get("cx"))
        geom["高_cm"] = emu_to_cm(ext.attrib.get("cy"))
    off = wrapper.find("wp:positionH/wp:posOffset", NS)
    if off is not None:
        geom["x_cm"] = emu_to_cm(off.text or "0")
    offv = wrapper.find("wp:positionV/wp:posOffset", NS)
    if offv is not None:
        geom["y_cm"] = emu_to_cm(offv.text or "0")
    rot = wrapper.find(".//a:xfrm", NS)
    if rot is not None and "rot" in rot.attrib:
        try:
            geom["旋轉度"] = round(int(rot.attrib.get("rot")) / 60000.0, 2)
        except Exception:
            pass
    return geom


# -------- WordArt 處理 --------
def parse_wordart(gdata: ET.Element, part_name: str, geom: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    info: Dict[str, Any] = {
        "來源": part_name,
        "文字": text_from_txbx(gdata),
        "幾何": geom,
        "版本": None,
        "style": {},
        "isWordArt": True
    }

    # w14 (Office 2019+ WordArt)
    te = gdata.find(".//w14:textEffect", NS)
    if te is not None:
        info["版本"] = "w14:textEffect (2019+)"
        if te.attrib.get("kern"):
            try:
                kern_pt = int(te.attrib["kern"]) / 100.0
                info["style"]["kern_pt"] = kern_pt
            except:
                info["style"]["kern_raw"] = te.attrib["kern"]
        info["style"]["size_pt"] = half_point_to_pt(te.attrib.get("fontSize"))
        info["style"]["fonts"] = {"name": te.attrib.get("font")}
        fill = te.find(".//a:solidFill/a:srgbClr", NS)
        if fill is not None:
            info["style"]["fill_color"] = "#" + fill.attrib.get("val", "")
        outline = te.find(".//a:ln/a:solidFill/a:srgbClr", NS)
        if outline is not None:
            info["style"]["outline_color"] = "#" + outline.attrib.get("val", "")
        return info

    # DrawingML (2010~2016)
    if gdata.find(".//a:bodyPr/a:prstTxWarp", NS) is not None:
        info["版本"] = "DrawingML (2010~2016)"
        font = gdata.find(".//a:latin", NS)
        if font is not None and font.attrib.get("typeface"):
            info["style"]["fonts"] = {"latin": font.attrib.get("typeface")}
        fill = gdata.find(".//a:solidFill/a:srgbClr", NS)
        if fill is not None:
            info["style"]["fill_color"] = "#" + fill.attrib.get("val", "")
        return info

    # VML (2003~2007)
    tp = gdata.find(".//v:textpath", NS)
    if tp is not None:
        info["版本"] = "VML (2003~2007)"
        style_str = tp.attrib.get("style", "")
        styles = dict(s.split(":") for s in style_str.split(";") if ":" in s)
        if "font-family" in styles:
            info["style"]["fonts"] = {"name": styles["font-family"].strip()}
        if "font-size" in styles:
            try:
                info["style"]["size_pt"] = float(styles["font-size"].replace("pt", "").strip())
            except:
                pass
        return info

    return None


def text_from_txbx(wrapper: ET.Element) -> str:
    """抽取文字方塊/WordArt中的文字"""
    texts: List[str] = []
    for t in wrapper.findall(".//w:txbxContent//w:t", NS):
        if t.text:
            texts.append(t.text)
    for t in wrapper.findall(".//a:txBody//a:t", NS):
        if t.text:
            texts.append(t.text)
    return "".join([x for x in texts if x.strip()])


# -------- 整合文字樣式 --------
def normalize_spacing(values: List[Any]) -> List[Any]:
    """將字距單位標準化：kern 轉 pt，段落 spacing 保持 cm"""
    result = []
    for v in values:
        if isinstance(v, str):
            try:
                result.append(int(v) / 100.0)  # kern 單位轉換 (1/100 pt)
            except:
                result.append(v)
        else:
            result.append(v)
    return result


def merge_across_sources(summary: Dict[str, Any]) -> Dict[str, Any]:
    """跨來源合併文字樣式，生成整合摘要"""
    para_map: Dict[str, Dict[str, Any]] = {}

    # 段落
    for para in summary.get("段落", []):
        txt = para.get("文字", "").strip()
        if not txt:
            continue
        if txt not in para_map:
            para_map[txt] = {
                "文字": txt,
                "來源列表": set(),
                "fonts": set(),
                "size_pt": set(),
                "color": set(),
                "isWordArt": False,
                "幾何": None,
                "段落屬性": {},
                "字距": set(),
                "outline_color": set(),
                "background_color": set(),
            }
        entry = para_map[txt]
        entry["來源列表"].add(para["來源"])
        for r in para.get("runs", []):
            st = r.get("style", {})
            if "fonts" in st and st["fonts"]:
                for v in st["fonts"].values():
                    if v:
                        entry["fonts"].add(v)
            if "size_pt" in st and st["size_pt"]:
                entry["size_pt"].add(st["size_pt"])
            if "color" in st and st["color"]:
                entry["color"].add(st["color"])
            if "spacing_cm" in st:
                entry["字距"].add(st["spacing_cm"])
            if "background_color" in st:
                entry["background_color"].add(st["background_color"])
        entry["段落屬性"] = para.get("屬性", {})

    # WordArt
    for wa in summary.get("WordArt", []):
        txt = wa.get("文字", "").strip()
        if not txt:
            continue
        if txt not in para_map:
            para_map[txt] = {
                "文字": txt,
                "來源列表": set(),
                "fonts": set(),
                "size_pt": set(),
                "color": set(),
                "isWordArt": True,
                "幾何": wa.get("幾何", {}),
                "段落屬性": {},
                "字距": set(),
                "outline_color": set(),
                "background_color": set(),
            }
        entry = para_map[txt]
        entry["來源列表"].add(wa.get("來源", ""))
        entry["isWordArt"] = True
        entry["幾何"] = wa.get("幾何", {})

        style = wa.get("style", {})
        if "fonts" in style and style["fonts"]:
            for v in style["fonts"].values():
                if v:
                    entry["fonts"].add(v)
        if "size_pt" in style and style["size_pt"]:
            entry["size_pt"].add(style["size_pt"])
        if "fill_color" in style and style["fill_color"]:
            entry["color"].add(style["fill_color"])
        if "kern_pt" in style:
            entry["字距"].add(style["kern_pt"])
        if "outline_color" in style and style["outline_color"]:
            entry["outline_color"].add(style["outline_color"])
        if "background_color" in style and style["background_color"]:
            entry["background_color"].add(style["background_color"])

    # 標準化字距
    for data in para_map.values():
        data["字距"] = normalize_spacing(list(data["字距"]))

    return {
        "段落完整文字": [
            {
                "文字": data["文字"],
                "來源列表": list(data["來源列表"]),
                "fonts": list(data["fonts"]),
                "size_pt": list(data["size_pt"]),
                "color": list(data["color"]),
                "isWordArt": data["isWordArt"],
                "幾何": data["幾何"],
                "段落屬性": data["段落屬性"],
                "字距": list(data["字距"]),
                "outline_color": list(data["outline_color"]),
                "background_color": list(data["background_color"]),
            }
            for data in para_map.values()
        ]
    }
