# parsers/table_parser.py
import xml.etree.ElementTree as ET
from typing import Dict, Any, List, Optional
from .utils import NS, half_point_to_pt, emu_to_cm


class TableParser:
    """Word 表格解析器"""

    @staticmethod
    def parse_tables(root: ET.Element, part_name: str) -> List[Dict[str, Any]]:
        """
        解析文件中的所有表格

        Args:
            root: XML 根元素
            part_name: 部件名稱

        Returns:
            List[Dict[str, Any]]: 表格資訊列表
        """
        tables = []

        # 查找所有表格元素
        for idx, tbl in enumerate(root.findall(".//w:tbl", NS), 1):
            try:
                table_info = TableParser._parse_single_table(tbl, part_name, idx)
                if table_info:
                    tables.append(table_info)
            except Exception as e:
                # 記錄錯誤但不中斷其他表格解析
                print(f"解析表格 {idx} 時發生錯誤: {e}")
                continue

        return tables

    @staticmethod
    def _parse_single_table(tbl: ET.Element, part_name: str, table_index: int) -> Optional[Dict[str, Any]]:
        """
        解析單一表格

        Args:
            tbl: 表格 XML 元素
            part_name: 部件名稱
            table_index: 表格索引

        Returns:
            Dict[str, Any]: 表格資訊
        """
        # 獲取表格基本資訊
        rows = tbl.findall(".//w:tr", NS)
        if not rows:
            return None

        row_count = len(rows)
        col_count = TableParser._get_column_count(rows)

        # 解析表格內容
        table_data = TableParser._extract_table_data(rows)

        # 解析表格屬性
        table_props = TableParser._parse_table_properties(tbl)

        # 解析表格樣式
        table_style = TableParser._parse_table_style(tbl)

        return {
            "來源": part_name,
            "表格編號": table_index,
            "行列結構": {
                "行數": row_count,
                "列數": col_count,
                "總儲存格數": row_count * col_count
            },
            "表格資料": table_data,
            "表格屬性": table_props,
            "表格樣式": table_style,
            "內容統計": TableParser._analyze_table_content(table_data),
            "特殊元素": TableParser._find_special_elements(tbl)
        }

    @staticmethod
    def _get_column_count(rows: List[ET.Element]) -> int:
        """
        計算表格列數（以第一行為準）

        Args:
            rows: 表格行元素列表

        Returns:
            int: 列數
        """
        if not rows:
            return 0

        first_row = rows[0]
        cells = first_row.findall(".//w:tc", NS)
        return len(cells)

    @staticmethod
    def _extract_table_data(rows: List[ET.Element]) -> List[List[Dict[str, Any]]]:
        """
        提取表格資料

        Args:
            rows: 表格行元素列表

        Returns:
            List[List[Dict[str, Any]]]: 二維表格資料
        """
        table_data = []

        for row_idx, row in enumerate(rows):
            row_data = []
            cells = row.findall(".//w:tc", NS)

            for cell_idx, cell in enumerate(cells):
                cell_info = TableParser._parse_cell_content(cell, row_idx, cell_idx)
                row_data.append(cell_info)

            table_data.append(row_data)

        return table_data

    @staticmethod
    def _parse_cell_content(cell: ET.Element, row_idx: int, cell_idx: int) -> Dict[str, Any]:
        """
        解析單一儲存格內容

        Args:
            cell: 儲存格 XML 元素
            row_idx: 行索引
            cell_idx: 列索引

        Returns:
            Dict[str, Any]: 儲存格資訊
        """
        # 提取文字內容
        paragraphs = cell.findall(".//w:p", NS)
        cell_text = ""
        cell_runs = []

        for para in paragraphs:
            # 提取段落文字
            para_text = "".join([t.text for t in para.findall(".//w:t", NS) if t.text])
            cell_text += para_text + "\n"

            # 提取段落樣式
            for run in para.findall(".//w:r", NS):
                run_text_elem = run.find(".//w:t", NS)
                run_text = run_text_elem.text if run_text_elem is not None else ""

                if run_text:
                    run_style = TableParser._parse_run_style(run)
                    cell_runs.append({
                        "文字": run_text,
                        "樣式": run_style
                    })

        cell_text = cell_text.strip()

        # 解析儲存格屬性
        cell_props = TableParser._parse_cell_properties(cell)

        return {
            "位置": {
                "行": row_idx + 1,
                "列": cell_idx + 1
            },
            "文字": cell_text,
            "文字長度": len(cell_text),
            "樣式": cell_runs,
            "屬性": cell_props,
            "是否為空": not bool(cell_text.strip())
        }

    @staticmethod
    def _parse_run_style(run: ET.Element) -> Dict[str, Any]:
        """
        解析文字樣式（簡化版本，可與 paragraph_parser 共用）

        Args:
            run: 文字運行元素

        Returns:
            Dict[str, Any]: 樣式資訊
        """
        rPr = run.find(".//w:rPr", NS)
        if rPr is None:
            return {}

        style = {}

        # 字型
        rFonts = rPr.find(".//w:rFonts", NS)
        if rFonts is not None:
            style["字型"] = {
                "ascii": rFonts.attrib.get(f"{{{NS['w']}}}ascii"),
                "hAnsi": rFonts.attrib.get(f"{{{NS['w']}}}hAnsi")
            }

        # 字號
        sz = rPr.find(".//w:sz", NS)
        if sz is not None:
            style["字號_pt"] = half_point_to_pt(sz.attrib.get(f"{{{NS['w']}}}val"))

        # 粗體、斜體
        if rPr.find(".//w:b", NS) is not None:
            style["粗體"] = True
        if rPr.find(".//w:i", NS) is not None:
            style["斜體"] = True

        # 文字顏色
        color = rPr.find(".//w:color", NS)
        if color is not None and color.attrib.get(f"{{{NS['w']}}}val"):
            style["顏色"] = "#" + color.attrib.get(f"{{{NS['w']}}}val")

        return style

    @staticmethod
    def _parse_cell_properties(cell: ET.Element) -> Dict[str, Any]:
        """
        解析儲存格屬性

        Args:
            cell: 儲存格 XML 元素

        Returns:
            Dict[str, Any]: 儲存格屬性
        """
        tcPr = cell.find(".//w:tcPr", NS)
        if tcPr is None:
            return {}

        props = {}

        # 儲存格寬度
        tcW = tcPr.find(".//w:tcW", NS)
        if tcW is not None:
            width_val = tcW.attrib.get(f"{{{NS['w']}}}w")
            width_type = tcW.attrib.get(f"{{{NS['w']}}}type")
            if width_val and width_type == "dxa":
                try:
                    props["寬度_pt"] = half_point_to_pt(width_val)
                except:
                    pass

        # 垂直對齊
        vAlign = tcPr.find(".//w:vAlign", NS)
        if vAlign is not None:
            props["垂直對齊"] = vAlign.attrib.get(f"{{{NS['w']}}}val")

        # 背景顏色
        shd = tcPr.find(".//w:shd", NS)
        if shd is not None and shd.attrib.get(f"{{{NS['w']}}}fill"):
            props["背景顏色"] = "#" + shd.attrib.get(f"{{{NS['w']}}}fill")

        # 合併資訊
        vMerge = tcPr.find(".//w:vMerge", NS)
        if vMerge is not None:
            props["垂直合併"] = vMerge.attrib.get(f"{{{NS['w']}}}val", "continue")

        hMerge = tcPr.find(".//w:hMerge", NS)
        if hMerge is not None:
            props["水平合併"] = hMerge.attrib.get(f"{{{NS['w']}}}val", "continue")

        return props

    @staticmethod
    def _parse_table_properties(tbl: ET.Element) -> Dict[str, Any]:
        """
        解析表格屬性

        Args:
            tbl: 表格 XML 元素

        Returns:
            Dict[str, Any]: 表格屬性
        """
        tblPr = tbl.find(".//w:tblPr", NS)
        if tblPr is None:
            return {}

        props = {}

        # 表格寬度
        tblW = tblPr.find(".//w:tblW", NS)
        if tblW is not None:
            width_val = tblW.attrib.get(f"{{{NS['w']}}}w")
            width_type = tblW.attrib.get(f"{{{NS['w']}}}type")
            if width_val and width_type == "pct":
                try:
                    props["寬度百分比"] = int(width_val) / 50  # 5000 = 100%
                except:
                    pass

        # 表格對齊
        jc = tblPr.find(".//w:jc", NS)
        if jc is not None:
            props["對齊方式"] = jc.attrib.get(f"{{{NS['w']}}}val")

        # 表格邊框
        tblBorders = tblPr.find(".//w:tblBorders", NS)
        if tblBorders is not None:
            props["邊框"] = TableParser._parse_borders(tblBorders)

        # 表格間距
        tblCellSpacing = tblPr.find(".//w:tblCellSpacing", NS)
        if tblCellSpacing is not None:
            spacing_val = tblCellSpacing.attrib.get(f"{{{NS['w']}}}w")
            if spacing_val:
                try:
                    props["儲存格間距_pt"] = half_point_to_pt(spacing_val)
                except:
                    pass

        return props

    @staticmethod
    def _parse_table_style(tbl: ET.Element) -> Dict[str, Any]:
        """
        解析表格樣式

        Args:
            tbl: 表格 XML 元素

        Returns:
            Dict[str, Any]: 表格樣式
        """
        style = {}

        # 表格樣式名稱
        tblStyle = tbl.find(".//w:tblPr/w:tblStyle", NS)
        if tblStyle is not None:
            style["樣式名稱"] = tblStyle.attrib.get(f"{{{NS['w']}}}val")

        # 表格條件格式
        tblLook = tbl.find(".//w:tblPr/w:tblLook", NS)
        if tblLook is not None:
            style["條件格式"] = {
                "首行": tblLook.attrib.get(f"{{{NS['w']}}}firstRow", "0") == "1",
                "末行": tblLook.attrib.get(f"{{{NS['w']}}}lastRow", "0") == "1",
                "首列": tblLook.attrib.get(f"{{{NS['w']}}}firstColumn", "0") == "1",
                "末列": tblLook.attrib.get(f"{{{NS['w']}}}lastColumn", "0") == "1"
            }

        return style

    @staticmethod
    def _parse_borders(borders_elem: ET.Element) -> Dict[str, Any]:
        """
        解析邊框設定

        Args:
            borders_elem: 邊框 XML 元素

        Returns:
            Dict[str, Any]: 邊框資訊
        """
        borders = {}
        border_types = ["top", "left", "bottom", "right", "insideH", "insideV"]

        for border_type in border_types:
            border = borders_elem.find(f".//w:{border_type}", NS)
            if border is not None:
                borders[border_type] = {
                    "顏色": "#" + border.attrib.get(f"{{{NS['w']}}}color", "000000"),
                    "大小_pt": half_point_to_pt(border.attrib.get(f"{{{NS['w']}}}sz")),
                    "樣式": border.attrib.get(f"{{{NS['w']}}}val", "single")
                }

        return borders

    @staticmethod
    def _analyze_table_content(table_data: List[List[Dict[str, Any]]]) -> Dict[str, Any]:
        """
        分析表格內容統計

        Args:
            table_data: 表格資料

        Returns:
            Dict[str, Any]: 內容統計
        """
        total_cells = 0
        empty_cells = 0
        total_text_length = 0
        text_cells = 0

        for row in table_data:
            for cell in row:
                total_cells += 1
                if cell["是否為空"]:
                    empty_cells += 1
                else:
                    total_text_length += cell["文字長度"]
                    text_cells += 1

        return {
            "總儲存格數": total_cells,
            "空儲存格數": empty_cells,
            "有文字儲存格數": text_cells,
            "平均文字長度": round(total_text_length / text_cells, 2) if text_cells > 0 else 0,
            "填滿率": round((text_cells / total_cells) * 100, 2) if total_cells > 0 else 0
        }

    @staticmethod
    def _find_special_elements(tbl: ET.Element) -> List[str]:
        """
        尋找表格中的特殊元素

        Args:
            tbl: 表格 XML 元素

        Returns:
            List[str]: 特殊元素列表
        """
        special_elements = []

        # 檢查是否有圖片
        if tbl.find(".//w:drawing", NS) is not None:
            special_elements.append("圖片")

        # 檢查是否有 WordArt
        if (tbl.find(".//w14:textEffect", NS) is not None or
                tbl.find(".//v:textpath", NS) is not None):
            special_elements.append("WordArt")

        # 檢查是否有特殊格式
        if tbl.find(".//w:shd[@w:fill]", NS) is not None:
            special_elements.append("背景顏色")

        return special_elements


# 便利函數
def parse_all_tables_in_document(root: ET.Element, part_name: str, out: Dict[str, Any]) -> None:
    """
    解析文件中的所有表格並添加到輸出中

    Args:
        root: XML 根元素
        part_name: 部件名稱
        out: 輸出字典
    """
    parser = TableParser()
    tables = parser.parse_tables(root, part_name)

    if tables:
        out["表格"] = tables
        out["提示"].append(f"檢測到 {len(tables)} 個表格")