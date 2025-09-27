# utils.py
import re
import unicodedata
import os
from typing import Any, Optional
from flask import abort
from werkzeug.utils import secure_filename
from docx import Document  # ← 改用 python-docx

# ---------------- 文字處理 ----------------
def normalize_text(s: str) -> str:
    return " ".join((s or "").replace("\r", "").split())

def _normalize_for_match(s: str) -> str:
    if not s:
        return ""
    s = s.replace("\u3000", " ")
    s = re.sub(r"[\uF000-\uF8FF\u25A0-\u25FF\u2700-\u27BF\u2000-\u206F\u2460-\u24FF\u2600-\u26FF\u2190-\u21FF]", "", s, flags=re.UNICODE)
    s = re.sub(r"[^\w\u4e00-\u9fff]+", "", s, flags=re.UNICODE)
    return s.lower()

def norm_key(s: str) -> str:
    s = unicodedata.normalize('NFKC', str(s or '')).strip().lower()
    s = re.sub(r'\s+', '', s)
    s = re.sub(r'[^\w\u4e00-\u9fff]+', '', s, flags=re.UNICODE)
    return s

def pt_to_cm(pt: float) -> float:
    return float(pt) / 28.3465

def cm_to_pt(cm: float) -> float:
    return float(cm) * 28.3465

def extract_json_object(text: str) -> str:
    if not text:
        return ""
    text = re.sub(r"```(?:json)?\s*", "", text)
    text = text.replace("```", "")
    start = text.find("{")
    end = text.rfind("}")
    return text[start:end+1] if (start != -1 and end != -1 and end > start) else ""

def clean_json_text(text: str) -> str:
    if not text:
        return text
    text = (text.replace("“", "\"").replace("”", "\"")
                .replace("‘", "'").replace("’", "'"))
    text = re.sub(r",(\s*[}\]])", r"\1", text.strip())
    return text

# ---------------- 檔案處理 ----------------
def allowed_file(filename: str, allowed_exts: set) -> bool:
    """
    檢查檔案副檔名是否允許
    :param filename: 檔案名稱
    :param allowed_exts: 允許的副檔名集合，例如 {".docx", ".docm", ".doc", ".rtf"}
    """
    if not filename:
        return False
    ext = os.path.splitext(filename)[1].lower()
    return ext in allowed_exts

def open_doc_or_error(path: str) -> Document:
    """
    開啟 Word 檔案，失敗則回傳 400 錯誤
    :param path: 檔案路徑
    """
    try:
        return Document(path)  # ← 改成 python-docx 的 Document
    except Exception as e:
        abort(400, f"文件解析失敗：{type(e).__name__} {e}")
