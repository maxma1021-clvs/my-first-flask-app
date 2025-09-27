# analysis/document_analyzer.py
import json
from typing import Dict, Any, List
from collections import Counter


class DocumentAnalyzer:
    """文件整合分析器"""

    def __init__(self):
        pass

    def analyze_document(self, file_path: str) -> Dict[str, Any]:
        """完整分析文件並生成整合報告"""
        try:
            # 由於我們無法直接訪問 DocParser，這裡假設 raw_data 已經被傳入
            # 在實際使用中，這個方法會在 Flask 路由中被調用
            pass

        except Exception as e:
            return {"error": f"文件分析失敗: {str(e)}"}

    def analyze_from_summary(self, summary_dict: Dict[str, Any]) -> Dict[str, Any]:
        """從摘要數據分析文件"""
        try:
            # 1. 生成各元素的整合敘述
            analysis_result = {
                "文件概覽": self._generate_document_overview(summary_dict),
                "段落分析": self._analyze_paragraphs(summary_dict),
                "圖片分析": self._analyze_images(summary_dict),
                "表格分析": self._analyze_tables(summary_dict),
                "樣式分析": self._analyze_styles(summary_dict),
                "WordArt分析": self._analyze_wordart(summary_dict),
                "綜合評估": self._generate_comprehensive_assessment(summary_dict)
            }

            return analysis_result

        except Exception as e:
            return {"error": f"文件分析失敗: {str(e)}"}

    def _generate_document_overview(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """生成文件概覽"""
        return {
            "總段落數": len(data.get("段落", [])),
            "WordArt數量": len(data.get("WordArt", [])),
            "圖片數量": len(data.get("圖片", [])) if "圖片" in data else 0,
            "表格數量": len(data.get("表格", [])) if "表格" in data else 0,
            "主要字型": self._extract_main_fonts(data),
            "平均字號": self._calculate_average_font_size(data),
            "文件複雜度": self._assess_complexity(data)
        }

    def _analyze_paragraphs(self, data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """分析所有段落並生成詳細敘述"""
        paragraphs = data.get("段落", [])
        analysis = []

        for i, para in enumerate(paragraphs, 1):
            # 詳細分析段落樣式
            style_analysis = self._analyze_paragraph_styles(para)

            desc = {
                "段落編號": i,
                "內容": para.get("文字", ""),
                "內容預覽": para.get("文字", "")[:100] + "..." if len(para.get("文字", "")) > 100 else para.get("文字",
                                                                                                                ""),
                "文字長度": len(para.get("文字", "")),
                "樣式分析": style_analysis,
                "格式設定": para.get("屬性", {}),
                "特殊元素": self._identify_special_elements(para),
                "runs數量": len(para.get("runs", [])),
                "平均字號": self._calculate_average_font_size_in_paragraph(para),
                "主要字型": self._extract_main_font_in_paragraph(para)
            }
            analysis.append(desc)

        return analysis

    def _analyze_paragraph_styles(self, para: Dict[str, Any]) -> List[Dict[str, Any]]:
        """詳細分析段落樣式"""
        runs = para.get("runs", [])
        styles = []

        for i, run in enumerate(runs):
            style = run.get("style", {})
            if style:
                style_desc = {
                    "文字片段": run.get("text", "")[:20] + "..." if len(run.get("text", "")) > 20 else run.get("text",
                                                                                                               ""),
                    "文字長度": len(run.get("text", "")),
                    "粗體": style.get("bold", False),
                    "斜體": style.get("italic", False),
                    "底線": style.get("underline", False),
                    "刪除線": style.get("strike", False),
                    "字型": style.get("fonts", {}),
                    "字號_pt": style.get("size_pt", 0),
                    "顏色": style.get("color", "預設"),
                    "背景色": style.get("background_color", "無"),
                    "字距_cm": style.get("spacing_cm", 0)
                }
                styles.append(style_desc)
        return styles

    def _calculate_average_font_size_in_paragraph(self, para: Dict[str, Any]) -> float:
        """計算段落平均字號"""
        sizes = []
        for run in para.get("runs", []):
            style = run.get("style", {})
            size = style.get("size_pt")
            if size:
                sizes.append(size)
        return round(sum(sizes) / len(sizes), 2) if sizes else 0

    def _extract_main_font_in_paragraph(self, para: Dict[str, Any]) -> str:
        """提取段落主要字型"""
        fonts = []
        for run in para.get("runs", []):
            style = run.get("style", {})
            font_info = style.get("fonts", {})
            # 優先順序：ascii > hAnsi > eastAsia
            main_font = font_info.get("ascii") or font_info.get("hAnsi") or font_info.get("eastAsia") or "預設"
            if main_font:
                fonts.append(main_font)

        # 返回最常見的字型
        if fonts:
            font_counter = Counter(fonts)
            return font_counter.most_common(1)[0][0]
        return "預設"

    def _analyze_images(self, data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """分析圖片元素"""
        images = data.get("圖片", [])
        analysis = []

        for i, img in enumerate(images, 1):
            desc = {
                "圖片編號": i,
                "類型": img.get("類型", "未知"),
                "尺寸": f"{img.get('幾何資訊', {}).get('寬_cm', 0)}cm x {img.get('幾何資訊', {}).get('高_cm', 0)}cm",
                "位置": img.get("位置資訊", {}),
                "視覺效果": list(img.get("視覺效果", {}).keys()),
                "特殊屬性": img.get("特殊屬性", [])
            }
            analysis.append(desc)

        return analysis

    def _analyze_tables(self, data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """分析表格元素"""
        tables = data.get("表格", [])
        analysis = []

        for i, table in enumerate(tables, 1):
            structure = table.get("行列結構", {})
            desc = {
                "表格編號": i,
                "行列結構": f"{structure.get('行數', 0)}行 x {structure.get('列數', 0)}列",
                "填滿率": f"{table.get('內容統計', {}).get('填滿率', 0)}%",
                "特殊元素": table.get("特殊元素", []),
                "背景顏色": any("背景顏色" in str(cell.get("屬性", {}))
                                for row in table.get("表格資料", [])
                                for cell in row)
            }
            analysis.append(desc)

        return analysis

    def _analyze_wordart(self, data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """分析 WordArt 元素"""
        wordarts = data.get("WordArt", [])
        analysis = []

        for i, wa in enumerate(wordarts, 1):
            desc = {
                "WordArt編號": i,
                "文字內容": wa.get("文字", ""),
                "版本類型": wa.get("版本", "未知"),
                "視覺效果": self._describe_wordart_effects(wa),
                "尺寸": f"{wa.get('幾何', {}).get('寬_cm', 0)}cm x {wa.get('幾何', {}).get('高_cm', 0)}cm",
                "樣式設定": wa.get("style", {})
            }
            analysis.append(desc)

        return analysis

    def _analyze_styles(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """分析整體樣式使用情況"""
        return {
            "字型使用統計": self._analyze_font_usage(data),
            "顏色使用統計": self._analyze_color_usage(data),
            "格式多樣性": self._analyze_format_diversity(data)
        }

    def _generate_comprehensive_assessment(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """生成綜合評估報告"""
        return {
            "文件品質評分": self._calculate_document_quality(data),
            "創意元素": self._identify_creative_elements(data),
            "技術複雜度": self._assess_technical_complexity(data)
        }

    # 輔助方法
    def _describe_wordart_effects(self, wordart: Dict[str, Any]) -> List[str]:
        """描述 WordArt 視覺效果"""
        effects = []
        style = wordart.get("style", {})

        if style.get("fill_color"):
            effects.append("填色效果")
        if style.get("outline_color"):
            effects.append("輪廓效果")
        if wordart.get("幾何", {}).get("旋轉度"):
            effects.append("旋轉效果")
        if style.get("kern_pt"):
            effects.append("字距調整")

        return effects if effects else ["基本效果"]

    def _extract_main_fonts(self, data: Dict[str, Any]) -> List[str]:
        """提取主要使用字型"""
        fonts = set()
        for para in data.get("段落", []):
            for run in para.get("runs", []):
                style = run.get("style", {})
                font_info = style.get("fonts", {})
                fonts.update(font_info.values())
        return list(filter(None, fonts))

    def _calculate_average_font_size(self, data: Dict[str, Any]) -> float:
        """計算平均字號"""
        sizes = []
        for para in data.get("段落", []):
            for run in para.get("runs", []):
                style = run.get("style", {})
                size = style.get("size_pt")
                if size:
                    sizes.append(size)
        return round(sum(sizes) / len(sizes), 2) if sizes else 0

    def _assess_complexity(self, data: Dict[str, Any]) -> str:
        """評估文件複雜度"""
        elements_count = (len(data.get("段落", [])) +
                          len(data.get("WordArt", [])) +
                          len(data.get("圖片", [])) if "圖片" in data else 0 +
                                                                           len(data.get("表格",
                                                                                        [])) if "表格" in data else 0)
        if elements_count > 50:
            return "高複雜度"
        elif elements_count > 20:
            return "中等複雜度"
        else:
            return "低複雜度"

    def _calculate_document_quality(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """計算文件品質評分"""
        paragraph_count = len(data.get("段落", []))
        wordart_count = len(data.get("WordArt", []))
        image_count = len(data.get("圖片", [])) if "圖片" in data else 0
        table_count = len(data.get("表格", [])) if "表格" in data else 0
        style_variety = len(self._extract_main_fonts(data))

        score = min(100, (paragraph_count * 2) + (wordart_count * 10) +
                    (image_count * 5) + (table_count * 8) + (style_variety * 3))

        return {
            "總分": score,
            "評分維度": {
                "內容豐富度": min(30, paragraph_count * 2),
                "創意元素": min(40, wordart_count * 10),
                "視覺元素": min(20, (image_count + table_count) * 5),
                "樣式多樣性": min(10, style_variety * 3)
            }
        }

    def _identify_creative_elements(self, data: Dict[str, Any]) -> List[str]:
        """識別創意元素"""
        creative = []
        if data.get("WordArt", []):
            creative.append("藝術字")
        if "圖片" in data and data.get("圖片", []):
            creative.append("圖片")
        if "表格" in data and data.get("表格", []):
            creative.append("表格")
        return creative

    def _assess_technical_complexity(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """評估技術複雜度"""
        return {
            "格式複雜度": len(self._extract_main_fonts(data)),
            "特效使用": sum(len(img.get("視覺效果", {})) for img in data.get("圖片", [])) if "圖片" in data else 0,
            "樣式變化": sum(len(para.get("runs", [])) for para in data.get("段落", []))
        }

    def _analyze_font_usage(self, data: Dict[str, Any]) -> Dict[str, int]:
        """分析字型使用情況"""
        font_count = {}
        for para in data.get("段落", []):
            for run in para.get("runs", []):
                style = run.get("style", {})
                fonts = style.get("fonts", {})
                for font in fonts.values():
                    if font:
                        font_count[font] = font_count.get(font, 0) + 1
        return font_count

    def _analyze_color_usage(self, data: Dict[str, Any]) -> Dict[str, int]:
        """分析顏色使用情況"""
        color_count = {}
        # 段落顏色
        for para in data.get("段落", []):
            for run in para.get("runs", []):
                style = run.get("style", {})
                color = style.get("color")
                if color:
                    color_count[color] = color_count.get(color, 0) + 1
        # WordArt 顏色
        for wa in data.get("WordArt", []):
            style = wa.get("style", {})
            fill_color = style.get("fill_color")
            if fill_color:
                color_count[fill_color] = color_count.get(fill_color, 0) + 1
        return color_count

    def _analyze_format_diversity(self, data: Dict[str, Any]) -> int:
        """分析格式多樣性"""
        formats = set()
        # 段落格式
        for para in data.get("段落", []):
            for run in para.get("runs", []):
                style = run.get("style", {})
                if style.get("bold"):
                    formats.add("粗體")
                if style.get("italic"):
                    formats.add("斜體")
                if style.get("underline"):
                    formats.add("底線")
                if style.get("strike"):
                    formats.add("刪除線")
                if style.get("color") and style.get("color") != "#000000":
                    formats.add("彩色文字")
        # WordArt 格式
        for wa in data.get("WordArt", []):
            effects = self._describe_wordart_effects(wa)
            formats.update(effects)
        return len(formats)

    def _identify_special_elements(self, para: Dict[str, Any]) -> List[str]:
        """識別段落中的特殊元素"""
        special = []
        props = para.get("屬性", {})

        # 對齊方式
        align = props.get("align")
        if align:
            if align == "center":
                special.append("置中對齊")
            elif align == "right":
                special.append("右對齊")
            elif align == "both":
                special.append("兩端對齊")

        # 特殊格式
        runs = para.get("runs", [])
        if any(run.get("style", {}).get("bold") for run in runs):
            special.append("包含粗體")
        if any(run.get("style", {}).get("italic") for run in runs):
            special.append("包含斜體")
        if any(run.get("style", {}).get("underline") for run in runs):
            special.append("包含底線")
        if any(run.get("style", {}).get("strike") for run in runs):
            special.append("包含刪除線")
        if any(run.get("style", {}).get("color") and run.get("style", {}).get("color") != "#000000" for run in runs):
            special.append("包含彩色文字")
        if any(run.get("style", {}).get("background_color") for run in runs):
            special.append("包含背景色")

        # 行距和縮排
        if props.get("spacing"):
            special.append("自訂行距")
        if props.get("indent"):
            special.append("自訂縮排")

        return special


# 生成可讀報告的函數
def generate_document_report(analysis_result: Dict[str, Any]) -> str:
    """生成完整文件分析報告"""
    if "error" in analysis_result:
        return f"分析錯誤: {analysis_result['error']}"

    if not analysis_result:
        return "無分析結果"

    overview = analysis_result.get('文件概覽', {})
    report = f"""
=== Word 文件分析報告 ===

【文件概覽】
- 總段落數: {overview.get('總段落數', 0)}
- WordArt 數量: {overview.get('WordArt數量', 0)}
- 圖片數量: {overview.get('圖片數量', 0)}
- 表格數量: {overview.get('表格數量', 0)}
- 主要字型: {', '.join(overview.get('主要字型', []))}
- 平均字號: {overview.get('平均字號', 0)} pt
- 文件複雜度: {overview.get('文件複雜度', '未知')}

【段落分析】
"""

    for para in analysis_result.get('段落分析', []):
        report += f"- 段落 {para['段落編號']}: {para['內容預覽']}\n"
        report += f"  長度: {para['文字長度']} 字符"
        if para.get('平均字號'):
            report += f", 平均字號: {para['平均字號']}pt"
        if para.get('主要字型') and para['主要字型'] != '預設':
            report += f", 主要字型: {para['主要字型']}"
        report += "\n"
        if para.get('特殊元素'):
            report += f"  特殊元素: {', '.join(para['特殊元素'])}\n"

    report += f"""
【圖片分析】
"""
    for img in analysis_result.get('圖片分析', []):
        report += f"- 圖片 {img['圖片編號']}: {img['尺寸']}\n"
        if img['視覺效果']:
            report += f"  效果: {', '.join(img['視覺效果'])}\n"

    report += f"""
【表格分析】
"""
    for table in analysis_result.get('表格分析', []):
        report += f"- 表格 {table['表格編號']}: {table['行列結構']}\n"
        report += f"  填滿率: {table['填滿率']}\n"

    report += f"""
【WordArt分析】
"""
    for wa in analysis_result.get('WordArt分析', []):
        report += f"- WordArt {wa['WordArt編號']}: {wa['文字內容'][:20]}...\n"
        report += f"  效果: {', '.join(wa['視覺效果'])}\n"

    assessment = analysis_result.get('綜合評估', {}).get('文件品質評分', {})
    report += f"""
【綜合評估】
總評分: {assessment.get('總分', 0)}/100
"""

    return report


# 確保 __init__.py 存在
def create_init_file():
    """創建 __init__.py 文件"""
    import os
    init_path = os.path.join(os.path.dirname(__file__), '__init__.py')
    if not os.path.exists(init_path):
        with open(init_path, 'w') as f:
            f.write('# Analysis module\n')