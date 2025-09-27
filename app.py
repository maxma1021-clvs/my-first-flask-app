# app.py
import os
import json
import time
from datetime import datetime
from flask import Flask, request, render_template, send_from_directory, redirect, url_for, abort
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge

from parsers.doc_parser import DocParser
from ai_scorer import AIScorer
from rules_loader import RulesLoader
from standard_manager import StandardAnswerManager
from utils import allowed_file, open_doc_or_error

# ---------------- Flask 基本設定 ----------------
app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
SCORE_RULES_XLSX = os.path.join(UPLOAD_FOLDER, "score_rules.xlsx")
STANDARD_ANSWERS_XLSX = os.path.join(UPLOAD_FOLDER, "standard_answers.xlsx")
ALLOWED_DOC_EXTS = {".docx"}  # ⚠️ 建議先限制 docx，避免 .doc 出問題
MAX_CONTENT_LENGTH_MB = 25

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH_MB * 1024 * 1024
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ---------------- 模組初始化 ----------------
doc_parser = DocParser()
ai_scorer = AIScorer()
rules_loader = RulesLoader()
standard_manager = StandardAnswerManager(STANDARD_ANSWERS_XLSX)


# ---------------- 錯誤處理 ----------------
@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(e):
    return render_template(
        "score_result.html",
        error=f"上傳檔案超過 {MAX_CONTENT_LENGTH_MB} MB，請壓縮或分割後再試。",
        summary_json="{}",
        summary_dict={},
        rules=[],
        scored={},
        rules_count=0,
        raw_ai_first="",
        raw_ai_fixed=""
    ), 413


@app.errorhandler(400)
def handle_bad_request(e):
    msg = getattr(e, "description", "Bad Request")
    return {"ok": False, "error": msg}, 400


# ---------------- 路由 ----------------
@app.route("/health")
def health():
    return {"ok": True}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload_rules", methods=["POST"])
def api_upload_rules():
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith(".xlsx"):
        abort(400, "請上傳 .xlsx 規則表")

    path = SCORE_RULES_XLSX
    f.save(path)
    return {"ok": True, "path": path, "updated_at": datetime.now().isoformat(timespec="seconds")}


@app.route("/score", methods=["POST"])
def score():
    file = request.files.get("file")
    if not file or not file.filename:
        return redirect(url_for("index"))

    filename = secure_filename(file.filename)
    if not allowed_file(filename, ALLOWED_DOC_EXTS):
        abort(400, f"不支援的檔案格式：{filename}")

    safe_name = f"{os.path.splitext(filename)[0]}_{int(time.time())}{os.path.splitext(filename)[1]}"
    path = os.path.join(app.config["UPLOAD_FOLDER"], safe_name)
    file.save(path)

    error = None
    summary_dict = {}
    summary_json = "{}"
    scored = {}
    rules = []

    # ---------- 解析文件 ----------
    try:
        summary_dict = doc_parser.summarize(file_path=path) or {}
        summary_json = json.dumps(summary_dict, ensure_ascii=False, indent=2)
    except Exception as e:
        error = f"文件解析失敗：{type(e).__name__} {e}"
        summary_dict = {"提示": [error]}
        summary_json = json.dumps(summary_dict, ensure_ascii=False, indent=2)

    # ---------- 載入規則 ----------
    try:
        rules = rules_loader.load_rules(SCORE_RULES_XLSX) or []
        if rules and "規則表" in rules[0]["檢查項目"]:
            return render_template("score_result.html",
                                   error=rules[0]["教學"],
                                   summary_json=summary_json,
                                   summary_dict=summary_dict,
                                   rules=[],
                                   scored={},
                                   rules_count=0,
                                   raw_ai_first="",
                                   raw_ai_fixed=""
                                   )
    except Exception as e:
        rules = [{"檢查項目": "規則表錯誤", "配分": 0, "教學": str(e)}]

    # ---------- AI 評分 ----------
    try:
        scored = ai_scorer.score_with_ai(summary_json, rules) or {}
    except Exception as e:
        error = f"AI 評分失敗：{type(e).__name__} {e}"
        scored = {"總分": 0, "逐項評語": []}

    # ---------- 標準答案檢查 ----------
    try:
        total_points = float(sum(float(r.get("配分", 0) or 0) for r in rules))
        got_points = float(scored.get("總分", 0) or 0)
        if abs(got_points - total_points) < 1e-6 and total_points > 0:
            upload_iso = datetime.fromtimestamp(os.path.getmtime(path)).isoformat(timespec="seconds")
            standard_manager.check_and_insert(
                item="全部",
                json_data=summary_json,
                comment="完全正確",
                filename=safe_name,
                upload_time=upload_iso
            )
    except Exception:
        pass

    return render_template(
        "score_result.html",
        error=error,
        summary_json=summary_json,
        summary_dict=summary_dict,  # ✅ 傳進模板
        rules=rules,
        scored=scored,
        rules_count=len(rules),
        raw_ai_first=scored.get("_raw_ai_first", ""),
        raw_ai_fixed=scored.get("_raw_ai_fixed", "")
    )


@app.route("/uploads/<path:filename>")
def uploaded_file(filename):
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)


# ==================== 新增的分析路由 ====================
@app.route("/analyze", methods=["GET"])
def analyze_page():
    """顯示分析頁面"""
    return render_template("analyze.html")


@app.route("/api/analyze", methods=["POST"])
def api_analyze():
    """API: 分析上傳的文件"""
    file = request.files.get("file")
    if not file or not file.filename:
        return {"error": "請上傳文件"}, 400

    filename = secure_filename(file.filename)
    if not allowed_file(filename, ALLOWED_DOC_EXTS):
        return {"error": f"不支援的檔案格式：{filename}"}, 400

    safe_name = f"{os.path.splitext(filename)[0]}_{int(time.time())}{os.path.splitext(filename)[1]}"
    path = os.path.join(app.config["UPLOAD_FOLDER"], safe_name)
    file.save(path)

    try:
        # 解析文件
        summary_dict = doc_parser.summarize(file_path=path) or {}
        summary_json = json.dumps(summary_dict, ensure_ascii=False, indent=2)

        return {
            "ok": True,
            "summary_dict": summary_dict,
            "summary_json": summary_json,
            "filename": safe_name
        }
    except Exception as e:
        return {"error": f"分析失敗: {str(e)}"}, 500


@app.route("/analysis/report/<filename>")
def analysis_report(filename):
    """顯示分析報告"""
    try:
        path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        if not os.path.exists(path):
            return "文件不存在", 404

        # 解析文件
        summary_dict = doc_parser.summarize(file_path=path) or {}
        summary_json = json.dumps(summary_dict, ensure_ascii=False, indent=2)

        return render_template("analysis_report.html",
                               filename=filename,
                               summary_dict=summary_dict,
                               summary_json=summary_json)
    except Exception as e:
        return f"生成報告失敗: {str(e)}", 500


# ==================== 輔助函數 ====================
def _identify_special_elements(para):
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


if __name__ == "__main__":
    app.run(debug=True)