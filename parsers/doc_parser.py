# parsers/doc_parser.py
import os
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, Any

from .utils import get_xml, merge_across_sources
from .paragraph_parser import parse_paragraphs
from .drawing_parser import DrawingParser
from .table_parser import parse_all_tables_in_document
from .image_parser import parse_all_images_in_document


class DocParser:
    @staticmethod
    def summarize(file_path: str) -> Dict[str, Any]:
        summary: Dict[str, Any] = {
            "段落": [],
            "WordArt": [],
            "提示": [],
            "整合文字樣式": {}
        }

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"❌ 找不到檔案: {file_path}")

        try:
            with zipfile.ZipFile(file_path, "r") as z:
                # 檢查 ZIP 檔案完整性
                bad_file = z.testzip()
                if bad_file:
                    raise zipfile.BadZipFile(f"ZIP 檔案損毀: {bad_file}")

                for name in z.namelist():
                    # 只處理 word/ 底下的 xml（排除 _rels）
                    if not name.startswith("word/") or not name.endswith(".xml") or "/_rels/" in name:
                        continue

                    root = get_xml(z, name)
                    if root is None:
                        continue

                    # 解析段落
                    parse_paragraphs(root, name, summary)

                    # 解析圖形 / WordArt
                    DrawingParser.parse(root, name, summary)

                    # 解析表格
                    parse_all_tables_in_document(root, name, summary)

                    # 解析圖片
                    parse_all_images_in_document(root, name, summary)

            # 合併相同樣式的 runs
            DocParser._merge_similar_runs(summary)

            # 整合樣式
            summary["整合文字樣式"] = merge_across_sources(summary)

            return summary

        except zipfile.BadZipFile as e:
            raise ValueError(f"無效的 Word 文件: {e}")
        except Exception as e:
            raise Exception(f"文件解析錯誤: {str(e)}")

    @staticmethod
    def _merge_similar_runs(summary: Dict[str, Any]) -> None:
        """
        合併段落中具有相同樣式的 runs
        """
        for para in summary.get("段落", []):
            runs = para.get("runs", [])
            if not runs:
                continue

            # 合併相同樣式的 runs
            merged_runs = DocParser._merge_runs_by_style(runs)
            para["runs"] = merged_runs

    @staticmethod
    def _merge_runs_by_style(runs: list) -> list:
        """
        根據相同樣式合併 runs

        Args:
            runs: 原始 runs 列表

        Returns:
            list: 合併後的 runs 列表
        """
        if not runs:
            return []

        merged = []
        current_group = {
            'text': runs[0].get('text', ''),
            'style': runs[0].get('style', {})
        }

        # 創建樣式比較鍵
        def create_style_key(style):
            """創建用於比較樣式的鍵"""
            if not style:
                return ""

            # 將樣式字典轉換為可比較的字符串
            key_parts = []
            # 按固定順序添加樣式屬性
            for attr in ['bold', 'italic', 'underline', 'strike', 'size_pt', 'color',
                         'background_color', 'spacing_cm']:
                if attr in style:
                    key_parts.append(f"{attr}:{style[attr]}")

            # 處理字型
            if 'fonts' in style and style['fonts']:
                fonts_str = ",".join([f"{k}:{v}" for k, v in sorted(style['fonts'].items())])
                key_parts.append(f"fonts:{fonts_str}")

            return "|".join(key_parts)

        current_style_key = create_style_key(runs[0].get('style', {}))

        for run in runs[1:]:
            style = run.get('style', {})
            text = run.get('text', '')
            style_key = create_style_key(style)

            if style_key == current_style_key and text:
                # 樣式相同且有文字，合併文字
                current_group['text'] += text
            else:
                # 樣式不同或沒有文字，保存當前組，開始新組
                if current_group['text']:  # 只有當有文字時才添加
                    merged.append(current_group)

                current_group = {
                    'text': text,
                    'style': style
                }
                current_style_key = style_key

        # 添加最後一組（如果有文字）
        if current_group['text']:
            merged.append(current_group)

        return merged


# 便利函數
def parse_document(file_path: str) -> Dict[str, Any]:
    """
    解析文件的便利函數

    Args:
        file_path: 文件路徑

    Returns:
        Dict[str, Any]: 解析結果
    """
    parser = DocParser()
    return parser.summarize(file_path)