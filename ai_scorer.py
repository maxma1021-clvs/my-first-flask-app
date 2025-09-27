import os
import re
import json
from typing import Dict, Any, List, Optional

from openai import OpenAI


class AIScorer:
    def __init__(self):
        self.api_key = (os.environ.get("DEEPSEEK_API_KEY") or "").strip()
        self.base_url = os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com").strip()
        self.model = os.environ.get("DEEPSEEK_MODEL", "deepseek-chat").strip()
        self.client: Optional[OpenAI] = None

        if self.api_key:
            self.client = OpenAI(api_key=self.api_key, base_url=self.base_url)

    def score_with_ai(self, summary_json: str, rules: List[Dict[str, Any]]) -> dict:
        if not self.api_key or not self.client:
            return {
                "總分": 0,
                "逐項評語": [
                    {"項目": "系統設定", "得分": 0,
                     "說明": "DEEPSEEK_API_KEY 未設定或為空。",
                     "教學": "請設定環境變數後重試。"}
                ],
                "_raw_ai_first": "[無法呼叫 API：缺少 DEEPSEEK_API_KEY]",
                "_raw_ai_fixed": ""
            }

        try:
            sj = json.loads(summary_json)
        except Exception:
            sj = {}

        prompt = f"""
你是一位嚴謹的文件評分老師。你會收到兩份資訊：
1) 學生文件的完整設定（JSON，可能已截斷）
2) 規則清單（每條包含：檢查項目、配分、教學）

【重要要求】
- 僅針對「規則清單中的每一條」輸出一筆評語，順序一致。
- 每筆格式：項目、得分（0~配分）、說明、教學。
- 總分 = 所有得分加總。
- 必須輸出 **有效 JSON**（開頭 {{，結尾 }}），不得有額外說明文字。

【規則清單】
{json.dumps(rules, ensure_ascii=False, indent=2)}

【學生文件設定】（可能因字數限制已截斷）
{json.dumps(sj, ensure_ascii=False)[:3000]}
"""

        try:
            resp = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0,
                timeout=60  # 加 timeout，避免卡住
            )
            raw_first = (resp.choices[0].message.content or "").strip()
        except Exception as e:
            return {
                "總分": 0,
                "逐項評語": [{"項目": "API失敗", "得分": 0, "說明": str(e), "教學": ""}],
                "_raw_ai_first": "",
                "_raw_ai_fixed": ""
            }

        parsed, raw_fixed = self._repair_json(raw_first)

        if parsed:
            # 強制轉型，避免 AI 回傳字串數字
            try:
                parsed["總分"] = int(parsed.get("總分", 0))
            except Exception:
                parsed["總分"] = 0
            parsed["_raw_ai_first"] = raw_first
            parsed["_raw_ai_fixed"] = raw_fixed
            return parsed

        return {
            "總分": 0,
            "逐項評語": [{"項目": "解析失敗", "得分": 0, "說明": raw_first[:200], "教學": ""}],
            "_raw_ai_first": raw_first,
            "_raw_ai_fixed": raw_fixed
        }

    # ---------------- JSON 修復 ----------------
    def _repair_json(self, raw_text: str):
        parsed, fixed = None, ""

        def clean_json_text(text: str) -> str:
            text = (text.replace("“", "\"").replace("”", "\"")
                        .replace("‘", "'").replace("’", "'"))
            text = re.sub(r",(\s*[}\]])", r"\1", text.strip())
            return text

        def extract_json_object(text: str) -> str:
            start = text.find("{")
            end = text.rfind("}")
            return text[start:end+1] if (start != -1 and end != -1 and end > start) else ""

        try:
            parsed = json.loads(raw_text)
        except Exception:
            candidate = clean_json_text(extract_json_object(raw_text))
            try:
                parsed = json.loads(candidate)
            except Exception:
                if not self.client:
                    return None, ""
                fix_prompt = f"下列內容應該是一個 JSON，但不是有效 JSON，請修正為有效 JSON：\n{raw_text}"
                try:
                    resp2 = self.client.chat.completions.create(
                        model=self.model,
                        messages=[{"role": "user", "content": fix_prompt}],
                        temperature=0,
                        timeout=30
                    )
                    fixed = (resp2.choices[0].message.content or "").strip()
                    parsed = json.loads(clean_json_text(extract_json_object(fixed)))
                except Exception:
                    parsed = None
        return parsed, fixed
