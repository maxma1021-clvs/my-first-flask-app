import os
import pandas as pd
from typing import Dict, Any, List


class RulesLoader:
    def __init__(self):
        # 欄位對照表（可擴充）
        self.col_map = {
            "检查项目": "檢查項目",
            "項目": "檢查項目",
            "points": "配分",
            "score": "配分",
            "指导": "教學",
            "指導": "教學",
            "guide": "教學",
        }

    def load_rules(self, xlsx_path: str) -> List[Dict[str, Any]]:
        if not os.path.exists(xlsx_path):
            return []

        try:
            df = pd.read_excel(xlsx_path)
        except Exception as e:
            return [{
                "檢查項目": "規則表載入失敗",
                "配分": 0,
                "教學": f"無法讀取 {os.path.basename(xlsx_path)}：{type(e).__name__} {e}"
            }]

        # 標準化欄位名稱（轉繁體 + 去空白）
        new_cols = []
        for c in df.columns:
            col = str(c).strip()
            col = self.col_map.get(col, col)
            new_cols.append(col)
        df.columns = new_cols

        required = {"檢查項目", "配分", "教學"}
        missing = required - set(df.columns)
        if missing:
            return [{
                "檢查項目": "規則表欄位錯誤",
                "配分": 0,
                "教學": f"缺少欄位: {', '.join(sorted(missing))}"
            }]

        if df.empty:
            return [{
                "檢查項目": "規則表為空",
                "配分": 0,
                "教學": "請在 score_rules.xlsx 填入規則（檢查項目／配分／教學）。"
            }]

        def to_number_safe(v):
            try:
                return float(v) if pd.notna(v) else 0
            except Exception:
                return 0

        rules: List[Dict[str, Any]] = []
        for _, row in df.iterrows():
            item = str(row.get("檢查項目", "")).strip()
            score = to_number_safe(row.get("配分", 0))
            guide = str(row.get("教學", "")).strip() or "未提供教學"
            if item:
                rules.append({
                    "檢查項目": item,
                    "配分": score,
                    "教學": guide
                })

        return rules
