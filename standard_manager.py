import os
import json
import pandas as pd
from typing import Any
from filelock import FileLock  # pip install filelock


class StandardAnswerManager:
    def __init__(self, excel_path: str):
        """
        初始化標準答案管理器
        :param excel_path: 標準答案 Excel 的路徑
        """
        self.excel_path = excel_path
        self.columns = ["項目", "JSON", "評語", "學生作答檔名", "上傳時間"]
        self.lock_path = f"{self.excel_path}.lock"

        # 如果檔案不存在，建立空檔案
        if not os.path.exists(self.excel_path):
            df = pd.DataFrame(columns=self.columns)
            df.to_excel(self.excel_path, index=False)

    def _normalize_json(self, json_data: str) -> str:
        """確保 JSON 內容一致性（避免縮排/順序不同造成重複紀錄）"""
        try:
            obj = json.loads(json_data)
            return json.dumps(obj, ensure_ascii=False, sort_keys=True)
        except Exception:
            return str(json_data).strip()

    def check_and_insert(self, item: str, json_data: str, comment: str,
                         filename: str, upload_time: Any) -> bool:
        """
        檢查 JSON 是否已存在於標準答案，若沒有則新增
        :param item: 規則項目名稱（例如「全部」）
        :param json_data: 學生文件的 JSON 摘要
        :param comment: 評語（通常是 "完全正確"）
        :param filename: 學生上傳的檔案名稱
        :param upload_time: 上傳時間 (ISO 字串)
        :return: 是否有新增紀錄
        """
        norm_json = self._normalize_json(json_data)

        with FileLock(self.lock_path):  # 避免併發寫入
            try:
                df = pd.read_excel(self.excel_path)
            except Exception:
                df = pd.DataFrame(columns=self.columns)

            # 檢查是否已存在相同 JSON
            if not df.empty and norm_json in df["JSON"].astype(str).tolist():
                return False

            new_row = {
                "項目": item,
                "JSON": norm_json,
                "評語": comment,
                "學生作答檔名": filename,
                "上傳時間": str(upload_time),
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

            df.to_excel(self.excel_path, index=False)

        return True
