# parsers/image_parser.py
import xml.etree.ElementTree as ET
from typing import Dict, Any, List, Optional
from .utils import NS, emu_to_cm, half_point_to_pt


class ImageParser:
    """Word 圖片解析器"""

    @staticmethod
    def parse_images(root: ET.Element, part_name: str) -> List[Dict[str, Any]]:
        """
        解析文件中的所有圖片

        Args:
            root: XML 根元素
            part_name: 部件名稱

        Returns:
            List[Dict[str, Any]]: 圖片資訊列表
        """
        images = []

        # 查找所有圖片元素
        # 1. 直接在段落中的圖片
        for idx, drawing in enumerate(root.findall(".//w:drawing", NS), 1):
            try:
                image_info = ImageParser._parse_drawing_element(drawing, part_name, idx)
                if image_info:
                    images.append(image_info)
            except Exception as e:
                print(f"解析圖片 {idx} 時發生錯誤: {e}")
                continue

        # 2. 舊版 VML 圖片
        for idx, pict in enumerate(root.findall(".//w:pict", NS), len(images) + 1):
            try:
                image_info = ImageParser._parse_vml_element(pict, part_name, idx)
                if image_info:
                    images.append(image_info)
            except Exception as e:
                print(f"解析 VML 圖片 {idx} 時發生錯誤: {e}")
                continue

        return images

    @staticmethod
    def _parse_drawing_element(drawing: ET.Element, part_name: str, img_index: int) -> Optional[Dict[str, Any]]:
        """
        解析 DrawingML 圖片元素

        Args:
            drawing: Drawing XML 元素
            part_name: 部件名稱
            img_index: 圖片索引

        Returns:
            Dict[str, Any]: 圖片資訊
        """
        # 獲取圖片關係 ID
        blip = drawing.find(".//a:blip", NS)
        if blip is None:
            return None

        embed_id = blip.attrib.get(f"{{{NS['r']}}}embed")
        if not embed_id:
            return None

        # 獲取幾何資訊
        geom = ImageParser._extract_geometry_info(drawing)

        # 獲取圖片屬性
        img_props = ImageParser._parse_image_properties(drawing)

        # 獲取圖片效果
        effects = ImageParser._parse_image_effects(drawing)

        # 獲取位置資訊
        position = ImageParser._extract_position_info(drawing)

        return {
            "來源": part_name,
            "圖片編號": img_index,
            "類型": "DrawingML",
            "關係ID": embed_id,
            "幾何資訊": geom,
            "位置資訊": position,
            "圖片屬性": img_props,
            "視覺效果": effects,
            "版本": "Modern (2007+)",
            "特殊屬性": ImageParser._extract_special_properties(drawing)
        }

    @staticmethod
    def _parse_vml_element(pict: ET.Element, part_name: str, img_index: int) -> Optional[Dict[str, Any]]:
        """
        解析 VML 圖片元素（舊版）

        Args:
            pict: VML 圖片 XML 元素
            part_name: 部件名稱
            img_index: 圖片索引

        Returns:
            Dict[str, Any]: 圖片資訊
        """
        # 查找 VML shape 元素
        shape = pict.find(".//v:shape", NS)
        if shape is None:
            return None

        # 獲取圖片關係 ID
        imagedata = shape.find(".//v:imagedata", NS)
        if imagedata is None:
            return None

        embed_id = imagedata.attrib.get(f"{{{NS['r']}}}id")
        if not embed_id:
            return None

        # 獲取 VML 幾何資訊
        geom = ImageParser._extract_vml_geometry(shape)

        # 獲取 VML 樣式
        vml_style = ImageParser._parse_vml_style(shape)

        return {
            "來源": part_name,
            "圖片編號": img_index,
            "類型": "VML",
            "關係ID": embed_id,
            "幾何資訊": geom,
            "VML樣式": vml_style,
            "版本": "Legacy (2003-2007)",
            "特殊屬性": ImageParser._extract_vml_special_properties(shape)
        }

    @staticmethod
    def _extract_geometry_info(drawing: ET.Element) -> Dict[str, Any]:
        """
        提取圖片幾何資訊

        Args:
            drawing: Drawing XML 元素

        Returns:
            Dict[str, Any]: 幾何資訊
        """
        geom = {}

        # 獲取尺寸
        extent = drawing.find(".//wp:extent", NS)
        if extent is not None:
            cx = extent.attrib.get("cx")
            cy = extent.attrib.get("cy")
            if cx:
                geom["寬_cm"] = emu_to_cm(cx)
            if cy:
                geom["高_cm"] = emu_to_cm(cy)

        # 獲取旋轉
        xfrm = drawing.find(".//a:xfrm", NS)
        if xfrm is not None:
            rot = xfrm.attrib.get("rot")
            if rot:
                try:
                    # 旋轉單位是 1/60000 度
                    geom["旋轉度"] = round(int(rot) / 60000.0, 2)
                except:
                    pass

        # 獲取縮放
        xfrm = drawing.find(".//a:xfrm", NS)
        if xfrm is not None:
            sx = xfrm.attrib.get("sx")
            sy = xfrm.attrib.get("sy")
            if sx:
                try:
                    geom["水平縮放"] = round(int(sx) / 100000.0, 2)
                except:
                    pass
            if sy:
                try:
                    geom["垂直縮放"] = round(int(sy) / 100000.0, 2)
                except:
                    pass

        return geom

    @staticmethod
    def _extract_vml_geometry(shape: ET.Element) -> Dict[str, Any]:
        """
        提取 VML 圖片幾何資訊

        Args:
            shape: VML shape 元素

        Returns:
            Dict[str, Any]: VML 幾何資訊
        """
        geom = {}

        # 獲取尺寸（從 style 屬性）
        style_attr = shape.attrib.get("style", "")
        if style_attr:
            styles = dict(s.split(":") for s in style_attr.split(";") if ":" in s)
            if "width" in styles:
                geom["寬度"] = styles["width"]
            if "height" in styles:
                geom["高度"] = styles["height"]

        # 獲取座標
        coord_attr = shape.attrib.get("coordsize", "")
        if coord_attr:
            geom["座標大小"] = coord_attr

        return geom

    @staticmethod
    def _extract_position_info(drawing: ET.Element) -> Dict[str, Any]:
        """
        提取圖片位置資訊

        Args:
            drawing: Drawing XML 元素

        Returns:
            Dict[str, Any]: 位置資訊
        """
        position = {}

        # 水平位置
        pos_h = drawing.find(".//wp:positionH", NS)
        if pos_h is not None:
            align = pos_h.find(".//wp:align", NS)
            if align is not None:
                position["水平對齊"] = align.text
            offset = pos_h.find(".//wp:posOffset", NS)
            if offset is not None and offset.text:
                position["水平偏移_cm"] = emu_to_cm(offset.text)

        # 垂直位置
        pos_v = drawing.find(".//wp:positionV", NS)
        if pos_v is not None:
            align = pos_v.find(".//wp:align", NS)
            if align is not None:
                position["垂直對齊"] = align.text
            offset = pos_v.find(".//wp:posOffset", NS)
            if offset is not None and offset.text:
                position["垂直偏移_cm"] = emu_to_cm(offset.text)

        # 錨定方式
        anchor = drawing.find(".//wp:anchor", NS)
        if anchor is not None:
            position["錨定方式"] = anchor.attrib.get("relativeHeight", "未知")

        return position

    @staticmethod
    def _parse_image_properties(drawing: ET.Element) -> Dict[str, Any]:
        """
        解析圖片屬性

        Args:
            drawing: Drawing XML 元素

        Returns:
            Dict[str, Any]: 圖片屬性
        """
        props = {}

        # 圖片格式
        pic = drawing.find(".//pic:pic", NS)
        if pic is not None:
            # 獲取圖片名稱
            nv_pic_pr = pic.find(".//pic:nvPicPr", NS)
            if nv_pic_pr is not None:
                c_nv_pic_pr = nv_pic_pr.find(".//pic:cNvPicPr", NS)
                if c_nv_pic_pr is not None:
                    props["圖片名稱"] = c_nv_pic_pr.attrib.get("name", "未命名")

            # 獲取圖片描述
            doc_pr = nv_pic_pr.find(".//pic:cNvPr", NS)
            if doc_pr is not None:
                props["圖片描述"] = doc_pr.attrib.get("descr", "")
                props["標題"] = doc_pr.attrib.get("title", "")

        # 圖片填充
        blip_fill = pic.find(".//pic:blipFill", NS)
        if blip_fill is not None:
            # 透明度
            alpha_mod = blip_fill.find(".//a:alphaMod", NS)
            if alpha_mod is not None:
                props["透明度"] = alpha_mod.attrib.get("amt", "100000")

        return props

    @staticmethod
    def _parse_image_effects(drawing: ET.Element) -> Dict[str, Any]:
        """
        解析圖片視覺效果

        Args:
            drawing: Drawing XML 元素

        Returns:
            Dict[str, Any]: 視覺效果
        """
        effects = {}

        effect_lst = drawing.find(".//a:effectLst", NS)
        if effect_lst is None:
            return effects

        # 陰影效果
        shadow = effect_lst.find(".//a:outerShdw", NS)
        if shadow is not None:
            effects["陰影"] = {
                "模糊半徑": shadow.attrib.get("blurRad", "0"),
                "距離": shadow.attrib.get("dist", "0"),
                "方向": shadow.attrib.get("dir", "0")
            }

        # 發光效果
        glow = effect_lst.find(".//a:glow", NS)
        if glow is not None:
            srgb_clr = glow.find(".//a:srgbClr", NS)
            effects["發光"] = {
                "半徑": glow.attrib.get("rad", "0"),
                "顏色": "#" + srgb_clr.attrib.get("val", "FFFFFF") if srgb_clr is not None else "FFFFFF"
            }

        # 軟邊效果
        soft_edge = effect_lst.find(".//a:softEdge", NS)
        if soft_edge is not None:
            effects["軟邊"] = {
                "半徑": soft_edge.attrib.get("rad", "0")
            }

        # 倒影效果
        reflection = effect_lst.find(".//a:reflection", NS)
        if reflection is not None:
            effects["倒影"] = dict(reflection.attrib)

        return effects

    @staticmethod
    def _parse_vml_style(shape: ET.Element) -> Dict[str, Any]:
        """
        解析 VML 樣式

        Args:
            shape: VML shape 元素

        Returns:
            Dict[str, Any]: VML 樣式
        """
        style = {}

        style_attr = shape.attrib.get("style", "")
        if style_attr:
            styles = dict(s.split(":") for s in style_attr.split(";") if ":" in s)
            style.update(styles)

        # 填充顏色
        fill_color = shape.attrib.get("fillcolor")
        if fill_color:
            style["填充顏色"] = fill_color

        # 邊框顏色
        stroke_color = shape.attrib.get("strokecolor")
        if stroke_color:
            style["邊框顏色"] = stroke_color

        # 邊框粗細
        stroke_weight = shape.attrib.get("strokeweight")
        if stroke_weight:
            style["邊框粗細"] = stroke_weight

        return style

    @staticmethod
    def _extract_special_properties(drawing: ET.Element) -> List[str]:
        """
        提取圖片特殊屬性

        Args:
            drawing: Drawing XML 元素

        Returns:
            List[str]: 特殊屬性列表
        """
        special_props = []

        # 檢查是否有裁剪
        if drawing.find(".//a:srcRect", NS) is not None:
            special_props.append("圖片裁剪")

        # 檢查是否有旋轉
        xfrm = drawing.find(".//a:xfrm", NS)
        if xfrm is not None and "rot" in xfrm.attrib:
            special_props.append("圖片旋轉")

        # 檢查是否有特效
        if drawing.find(".//a:effectLst", NS) is not None:
            special_props.append("視覺特效")

        # 檢查是否有文字框
        if drawing.find(".//wp:cNvGraphicFramePr", NS) is not None:
            special_props.append("文字框")

        return special_props

    @staticmethod
    def _extract_vml_special_properties(shape: ET.Element) -> List[str]:
        """
        提取 VML 圖片特殊屬性

        Args:
            shape: VML shape 元素

        Returns:
            List[str]: 特殊屬性列表
        """
        special_props = []

        # 檢查是否有旋轉
        if "rotation" in shape.attrib:
            special_props.append("VML旋轉")

        # 檢查是否有填充
        if "fill" in shape.attrib:
            special_props.append("VML填充")

        # 檢查是否有邊框
        if "stroke" in shape.attrib:
            special_props.append("VML邊框")

        return special_props


# 便利函數
def parse_all_images_in_document(root: ET.Element, part_name: str, out: Dict[str, Any]) -> None:
    """
    解析文件中的所有圖片並添加到輸出中

    Args:
        root: XML 根元素
        part_name: 部件名稱
        out: 輸出字典
    """
    parser = ImageParser()
    images = parser.parse_images(root, part_name)

    if images:
        out["圖片"] = images
        out["提示"].append(f"檢測到 {len(images)} 張圖片")


# 整合到主解析器的使用範例
def integrate_with_main_parser():
    """
    在主解析器中整合圖片解析的範例
    """
    # 在 doc_parser.py 的 summarize 方法中添加：
    """
    # 解析圖片
    from .image_parser import parse_all_images_in_document
    parse_all_images_in_document(root, name, summary)
    """
    pass