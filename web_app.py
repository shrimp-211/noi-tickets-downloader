#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NOI 准考证批量下载 —— 本地网页版

在浏览器中管理多个账号（手动添加 / 上传 CSV、Excel），自定义下载目录，
一键批量登录并下载各账号的活动准考证、获奖/等级证书，并查询历史活动成绩、导出 Excel。

启动：
    python web_app.py
然后浏览器打开 http://127.0.0.1:8765 （脚本会自动打开）
"""

from __future__ import annotations

import argparse
import csv
import io
import os
import re
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from copy import deepcopy
from typing import Dict, List, Optional, Tuple

from flask import Flask, jsonify, render_template, request, send_file
from playwright.sync_api import sync_playwright

import download_tickets as core

# ---------------------------------------------------------------- 应用与状态

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024  # 上传表格最大 8MB
app.json.ensure_ascii = False

ACCOUNT_LOCK = threading.RLock()
ACCOUNTS: List[dict] = []
NEXT_ACCOUNT_ID = 1

JOB_LOCK = threading.RLock()
JOB: Dict = {
    "running": False,
    "total": 0,
    "done": 0,
    "current": "",
    "out_dir": "",
    "layout": "",
    "filename_template": "",
    "started_at": None,
    "finished_at": None,
    "results": [],
    "log": [],
}

MAX_LOG_LINES = 300

SCORE_RESP_CLASS = 1558219344   # 成绩信息响应类
SCORE_ERROR_CLASS = 619155620   # 通用错误消息类

SCORE_LOCK = threading.RLock()
SCORE_JOB: Dict = {
    "running": False,
    "total": 0,
    "done": 0,
    "current": "",
    "started_at": None,
    "finished_at": None,
    "results": [],
    "log": [],
}

CERTIFICATE_TYPES = {
    "prize": ("获奖证书", "download_prize_certificate"),
    "grade": ("等级证书", "download_grade_certificate"),
}

CERT_LOCK = threading.RLock()
CERT_JOB: Dict = {
    "running": False,
    "total": 0,
    "done": 0,
    "current": "",
    "out_dir": "",
    "layout": "",
    "started_at": None,
    "finished_at": None,
    "results": [],
    "log": [],
}

# 表头别名（小写比较）
USER_HEADERS = {"账号", "用户名", "手机号", "用户", "邮箱", "user", "username", "account", "login"}
PASS_HEADERS = {"密码", "password", "passwd", "pwd", "pass"}
NOTE_HEADERS = {"备注", "姓名", "说明", "name", "note", "remark"}


# ---------------------------------------------------------------- 工具函数


def now_ts() -> float:
    return time.time()


def concise_error(exc: Exception) -> str:
    """异常信息只保留首行，避免 Playwright 调用日志撑满界面。"""
    text = str(exc).strip()
    first_line = text.splitlines()[0] if text else ""
    return f"{type(exc).__name__}: {first_line}" if first_line else type(exc).__name__


def log_line(text: str, level: str = "info") -> None:
    with JOB_LOCK:
        JOB["log"].append(
            {"time": time.strftime("%H:%M:%S"), "level": level, "text": text}
        )
        if len(JOB["log"]) > MAX_LOG_LINES:
            del JOB["log"][: len(JOB["log"]) - MAX_LOG_LINES]


def score_log(text: str, level: str = "info") -> None:
    with SCORE_LOCK:
        SCORE_JOB["log"].append(
            {"time": time.strftime("%H:%M:%S"), "level": level, "text": text}
        )
        if len(SCORE_JOB["log"]) > MAX_LOG_LINES:
            del SCORE_JOB["log"][: len(SCORE_JOB["log"]) - MAX_LOG_LINES]


def cert_log(text: str, level: str = "info") -> None:
    with CERT_LOCK:
        CERT_JOB["log"].append(
            {"time": time.strftime("%H:%M:%S"), "level": level, "text": text}
        )
        if len(CERT_JOB["log"]) > MAX_LOG_LINES:
            del CERT_JOB["log"][: len(CERT_JOB["log"]) - MAX_LOG_LINES]


PHONE_PATTERN = re.compile(r"(\+?86[\s-]?)?(1[3-9]\d{9})")


def mask_account(username: str) -> str:
    """手机号脱敏：将中间四位替换为 ****；其他账号原样返回。"""
    text = (username or "").strip()
    match = PHONE_PATTERN.fullmatch(text)
    if match:
        prefix = "+86 " if match.group(1) else ""
        number = match.group(2)
        return f"{prefix}{number[:3]}****{number[-4:]}"
    # 非手机号的纯数字账号也做同样处理，避免意外泄露
    if len(text) >= 7 and text.isdigit():
        return text[:3] + "****" + text[-4:]
    return text


def folder_account(username: str) -> str:
    """用于文件夹名的脱敏形式（* 在 Windows 文件名中非法，替换为 x）。"""
    return mask_account(username).replace("*", "x")


def safe_folder(name: str) -> str:
    return core.sanitize_filename(name)


def account_folder(account: dict) -> str:
    note = (account.get("note") or "").strip()
    masked = folder_account(account.get("username", ""))
    return safe_folder(f"{note}_{masked}" if note else masked)


LAYOUT_ACTIVITY = "activity"  # 按活动区分（默认）
LAYOUT_USER = "user"          # 按用户区分
LAYOUT_NONE = "none"          # 不区分
LAYOUTS = {LAYOUT_ACTIVITY, LAYOUT_USER, LAYOUT_NONE}


def activity_folder_name(activity) -> str:
    return safe_folder(activity.name) or f"活动_{activity.id}"


def resolve_target(layout: str, out_dir: str, user_tag: str, activity) -> str:
    """根据目录结构返回准考证的保存目录。"""
    if layout == LAYOUT_USER:
        return os.path.join(out_dir, user_tag)
    if layout == LAYOUT_NONE:
        return out_dir
    return os.path.join(out_dir, activity_folder_name(activity))


DEFAULT_FILENAME_TEMPLATE = "{original}"
FILENAME_PLACEHOLDERS = {
    "original": "服务端原始文件名（不含 .pdf）",
    "activity": "活动名称",
    "user": "备注 + 脱敏账号（与下载目录名一致）",
    "account": "脱敏账号",
    "note": "备注",
    "id": "活动 ID",
    "date": "下载日期（YYYYMMDD）",
}


def normalize_filename_template(template: str) -> str:
    template = (template or "").strip()
    if not template:
        return DEFAULT_FILENAME_TEMPLATE
    if len(template) > 200:
        raise ValueError("文件名格式过长（最多 200 个字符）")
    unknown = sorted(
        {name for name in re.findall(r"\{([^{}]+)\}", template) if name not in FILENAME_PLACEHOLDERS}
    )
    if unknown:
        raise ValueError("不支持的文件名占位符：" + "、".join("{" + name + "}" for name in unknown))
    return template


def make_filename_builder(template: str, user_tag: str, account: dict):
    """按模板渲染文件名；{original} 保持服务端原始文件名。"""
    note = (account.get("note") or "").strip()
    display = mask_account(account.get("username", ""))

    def builder(activity, original_filename: Optional[str]) -> str:
        base = original_filename or f"准考证_{activity.id}"
        if base.lower().endswith(".pdf"):
            base = base[:-4]
        values = {
            "original": base,
            "activity": activity.name,
            "user": user_tag,
            "account": display,
            "note": note,
            "id": str(activity.id),
            "date": time.strftime("%Y%m%d"),
        }
        return re.sub(
            r"\{([^{}]+)\}", lambda match: values.get(match.group(1), match.group(0)), template
        )

    return builder


# ---------------------------------------------------------------- 成绩解析


def format_score(raw) -> str:
    """站点分数以 100 为系数存储，转换为可读分数。"""
    try:
        value = int(raw)
    except (TypeError, ValueError):
        return str(raw)
    if value % 100 == 0:
        return str(value // 100)
    if value % 10 == 0:
        return f"{value / 100:.1f}"
    return f"{value / 100:.2f}"


def parse_score_response(data: bytes) -> dict:
    """解析成绩查询响应。"""
    reader = core.Reader(data)
    magic = reader.varint()

    if magic == SCORE_RESP_CLASS:
        try:
            activity_title = reader.string()
            ticket_no = reader.string()
            name = reader.string()
            count = reader.varint()
            problems = []
            for _ in range(count):
                problem_name = reader.string()
                score = reader.varint()
                reader.varint()  # 申诉相关标记，暂不使用
                problems.append({"name": problem_name, "score": score})
            total = reader.varint()
            remark = reader.string()
        except Exception as exc:
            return {"available": False, "message": f"成绩数据解析失败：{exc}"}
        return {
            "available": True,
            "activity_title": activity_title,
            "name": name,
            "ticket_no": ticket_no,
            "problems": problems,
            "total": total,
            "remark": remark,
        }

    if magic == SCORE_ERROR_CLASS:
        try:
            message = reader.string()
        except Exception:
            message = ""
        return {"available": False, "message": message or "暂无成绩信息"}

    return {"available": False, "message": f"未知成绩响应（{magic}）"}


# ---------------------------------------------------------------- 证书下载


def certificate_filename(activity, label: str, original_filename: Optional[str]) -> str:
    activity_part = core.sanitize_filename(activity.name)
    prefix = f"{activity_part}_{label}"
    if original_filename:
        original = core.sanitize_filename(original_filename)
        filename = f"{prefix}_{original}"
    else:
        filename = f"{prefix}_{activity.id}.pdf"
    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"
    return filename


def download_certificate(
    request, base_url: str, activity, cert_key: str, out_dir: str
) -> Tuple[Optional[str], Optional[str]]:
    """下载单个活动的证书，返回 (保存路径, 失败原因)。"""
    label, endpoint = CERTIFICATE_TYPES[cert_key]
    response = request.get(
        f"{base_url}/api/contestant/activity/{endpoint}?id={activity.id}",
        timeout=60_000,
    )
    if response.status != 200:
        return None, f"HTTP {response.status}"

    data = response.body()
    content_type = (response.headers.get("content-type") or "").split(";")[0].strip().lower()
    if data[:5] == b"%PDF-" or content_type == "application/pdf":
        original = core.filename_from_disposition(response.headers.get("content-disposition"))
        filename = certificate_filename(activity, label, original)
        path = core.unique_path(out_dir, filename)
        with open(path, "wb") as fh:
            fh.write(data)
        return path, None

    message = core.decode_message(data)
    if message:
        return None, message
    return None, f"返回内容不是 PDF（content-type={content_type or '未知'}，{len(data)} 字节）"


def cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


# ---------------------------------------------------------------- 表格解析


def decode_text(data: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "gb18030", "gbk", "utf-16"):
        try:
            return data.decode(encoding)
        except (UnicodeDecodeError, LookupError):
            continue
    return data.decode("utf-8", errors="replace")


def read_table_rows(filename: str, data: bytes) -> List[List[str]]:
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext in ("csv", "txt"):
        text = decode_text(data)
        sample = text[:4096]
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
            reader = csv.reader(io.StringIO(text), dialect)
        except csv.Error:
            reader = csv.reader(io.StringIO(text))
        return [[cell.strip() for cell in row] for row in reader]

    if ext in ("xlsx", "xlsm"):
        try:
            import openpyxl
        except ImportError as exc:  # pragma: no cover
            raise ValueError("缺少 openpyxl，请执行 pip install openpyxl") from exc
        workbook = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        sheet = workbook.active
        rows = []
        for row in sheet.iter_rows(values_only=True):
            rows.append([cell_text(value) for value in row])
        workbook.close()
        return rows

    if ext == "xls":
        raise ValueError("不支持旧版 .xls，请在 Excel 中另存为 .xlsx 或 .csv")
    raise ValueError("仅支持 .csv、.txt、.xlsx 表格")


def find_column(header: List[str], names: set) -> Optional[int]:
    for index, cell in enumerate(header):
        if cell.strip().lower() in names:
            return index
    return None


def parse_accounts_table(filename: str, data: bytes) -> Tuple[List[dict], List[dict]]:
    """返回 (待添加账号, 跳过明细)。"""
    rows = [row for row in read_table_rows(filename, data) if any(cell for cell in row)]
    if not rows:
        raise ValueError("表格内容为空")

    header = rows[0]
    user_col = find_column(header, USER_HEADERS)
    pass_col = find_column(header, PASS_HEADERS)
    note_col = find_column(header, NOTE_HEADERS)

    has_header = user_col is not None or pass_col is not None
    if not has_header:
        user_col, pass_col, note_col = 0, 1, 2
        data_rows = rows
    else:
        if user_col is None:
            user_col = 0
        if pass_col is None:
            pass_col = user_col + 1
        data_rows = rows[1:]

    parsed: List[dict] = []
    skipped: List[dict] = []
    seen = set()

    for line_no, row in enumerate(data_rows, start=2 if has_header else 1):
        def get(index: Optional[int]) -> str:
            if index is None or index >= len(row):
                return ""
            return row[index].strip()

        username = get(user_col)
        password = get(pass_col)
        note = get(note_col)

        if not username and not password:
            continue
        if username.startswith("#"):
            continue
        if not username:
            skipped.append({"row": line_no, "username": "", "reason": "账号为空"})
            continue
        if not password:
            skipped.append({"row": line_no, "username": mask_account(username), "reason": "密码为空"})
            continue
        if len(username) > 100 or len(password) > 200:
            skipped.append({"row": line_no, "username": mask_account(username[:20]), "reason": "内容过长"})
            continue
        if username in seen:
            skipped.append({"row": line_no, "username": mask_account(username), "reason": "表格内重复"})
            continue

        seen.add(username)
        parsed.append({"username": username, "password": password, "note": note[:60]})

    return parsed, skipped


# ---------------------------------------------------------------- 账号 API


def public_account(account: dict) -> dict:
    return {
        "id": account["id"],
        "display": mask_account(account["username"]),
        "note": account.get("note", ""),
    }


def add_account(username: str, password: str, note: str = "") -> Tuple[bool, str]:
    global NEXT_ACCOUNT_ID
    username = username.strip()
    password = password.strip()
    note = note.strip()[:60]
    if not username:
        return False, "账号不能为空"
    if not password:
        return False, "密码不能为空"
    if len(username) > 100 or len(password) > 200:
        return False, "账号或密码过长"
    with ACCOUNT_LOCK:
        if any(acc["username"] == username for acc in ACCOUNTS):
            return False, "该账号已存在"
        ACCOUNTS.append(
            {
                "id": NEXT_ACCOUNT_ID,
                "username": username,
                "password": password,
                "note": note,
            }
        )
        NEXT_ACCOUNT_ID += 1
    return True, "已添加"


@app.get("/api/accounts")
def api_accounts():
    with ACCOUNT_LOCK:
        return jsonify({"ok": True, "accounts": [public_account(a) for a in ACCOUNTS]})


@app.post("/api/accounts")
def api_add_account():
    payload = request.get_json(silent=True) or {}
    ok, message = add_account(
        str(payload.get("username", "")),
        str(payload.get("password", "")),
        str(payload.get("note", "")),
    )
    return jsonify({"ok": ok, "message": message}), (200 if ok else 400)


@app.delete("/api/accounts/<int:account_id>")
def api_delete_account(account_id: int):
    with ACCOUNT_LOCK:
        for index, account in enumerate(ACCOUNTS):
            if account["id"] == account_id:
                ACCOUNTS.pop(index)
                return jsonify({"ok": True})
    return jsonify({"ok": False, "message": "账号不存在"}), 404


@app.post("/api/accounts/clear")
def api_clear_accounts():
    with ACCOUNT_LOCK:
        ACCOUNTS.clear()
    return jsonify({"ok": True, "message": "已清空账号列表"})


@app.post("/api/accounts/upload")
def api_upload_accounts():
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"ok": False, "message": "没有收到文件"}), 400

    try:
        parsed, skipped = parse_accounts_table(file.filename, file.read())
    except ValueError as exc:
        return jsonify({"ok": False, "message": str(exc)}), 400

    added, duplicates = 0, []
    for item in parsed:
        ok, message = add_account(item["username"], item["password"], item["note"])
        if ok:
            added += 1
        else:
            duplicates.append({"username": mask_account(item["username"]), "reason": message})

    return jsonify(
        {
            "ok": True,
            "added": added,
            "skipped": skipped,
            "duplicates": duplicates,
            "message": f"导入完成：新增 {added} 个，跳过 {len(skipped) + len(duplicates)} 个",
        }
    )


@app.get("/api/template.csv")
def api_template():
    content = (
        "账号,密码,备注\n"
        "# 示例（号码已脱敏）：138****0000,你的密码,张三（以 # 开头的行会被忽略，可直接删除）\n"
    )
    buffer = io.BytesIO(("﻿" + content).encode("utf-8"))
    return send_file(
        buffer,
        mimetype="text/csv; charset=utf-8",
        as_attachment=True,
        download_name="账号导入模板.csv",
    )


# ---------------------------------------------------------------- 下载任务


def reset_job(total: int, out_dir: str, layout: str, filename_template: str) -> None:
    with JOB_LOCK:
        JOB.update(
            {
                "running": True,
                "total": total,
                "done": 0,
                "current": "",
                "out_dir": out_dir,
                "layout": layout,
                "filename_template": filename_template,
                "started_at": now_ts(),
                "finished_at": None,
                "results": [],
                "log": [],
            }
        )


def run_job(config: dict, accounts: List[dict]) -> None:
    base_url = config["base_url"]
    out_dir = config["out_dir"]
    delay = config["delay"]

    log_line(f"任务开始，共 {len(accounts)} 个账号")
    log_line(f"下载目录：{out_dir}")

    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=not config["headed"])
            log_line("浏览器已启动" + ("（显示窗口）" if config["headed"] else "（后台运行）"))
            try:
                for index, account in enumerate(accounts, 1):
                    username = account["username"]
                    label = account.get("note") or mask_account(username)
                    # 结果先在本地组装，完成后整体替换占位项，避免轮询读取时发生并发写入
                    result = {
                        "account": mask_account(username),
                        "note": account.get("note", ""),
                        "status": "running",
                        "message": "",
                        "files": [],
                        "skipped": [],
                    }
                    with JOB_LOCK:
                        JOB["results"].append(deepcopy(result))
                        result_index = len(JOB["results"]) - 1
                        JOB["current"] = label

                    log_line(f"[{index}/{len(accounts)}] {label} 开始处理")
                    context = None
                    try:
                        context = browser.new_context(viewport={"width": 1280, "height": 900})
                        page = context.new_page()
                        page.set_default_timeout(30_000)

                        core.login(
                            page,
                            base_url,
                            username,
                            account["password"],
                            config["login_timeout"],
                        )
                        log_line(f"{mask_account(username)} 登录成功")

                        activities = core.fetch_all_activities(
                            context.request, base_url, core.TAB_CURRENT, False, 50
                        )
                        if config["history"]:
                            activities += core.fetch_all_activities(
                                context.request, base_url, core.TAB_HISTORY, False, 50
                            )

                        if not activities:
                            result["status"] = "empty"
                            result["message"] = "没有找到任何活动"
                            log_line(f"{mask_account(username)} 没有活动")
                        else:
                            user_tag = account_folder(account)
                            filename_builder = make_filename_builder(
                                config["filename_template"], user_tag, account
                            )
                            for activity in activities:
                                folder = resolve_target(
                                    config["layout"], out_dir, user_tag, activity
                                )
                                try:
                                    os.makedirs(folder, exist_ok=True)
                                except OSError as exc:
                                    result["skipped"].append(
                                        {"activity": activity.name, "reason": f"无法创建目录：{exc}"}
                                    )
                                    log_line(
                                        f"{mask_account(username)} 跳过「{activity.name}」："
                                        f"无法创建目录 {exc}",
                                        "error",
                                    )
                                    continue

                                path, reason = core.download_ticket(
                                    context.request,
                                    base_url,
                                    activity,
                                    folder,
                                    filename_builder=filename_builder,
                                )
                                if path:
                                    result["files"].append(
                                        {
                                            "name": os.path.basename(path),
                                            "path": path,
                                            "activity": activity.name,
                                        }
                                    )
                                    log_line(
                                        f"{mask_account(username)} 已保存 {os.path.basename(path)}"
                                    )
                                else:
                                    result["skipped"].append(
                                        {"activity": activity.name, "reason": reason}
                                    )
                                    log_line(
                                        f"{mask_account(username)} 跳过「{activity.name}」：{reason}",
                                        "warn",
                                    )
                                    # 清理因本次跳过而新建的空目录
                                    if folder != out_dir:
                                        try:
                                            if not os.listdir(folder):
                                                os.rmdir(folder)
                                        except OSError:
                                            pass

                            if result["files"] and result["skipped"]:
                                result["status"] = "partial"
                            elif result["files"]:
                                result["status"] = "success"
                            else:
                                result["status"] = "skipped"
                                result["message"] = "所有活动均未开放准考证下载"
                    except core.LoginError as exc:
                        message = str(exc).strip()
                        result["status"] = "failed"
                        result["message"] = message if message.startswith("登录失败") else f"登录失败：{message}"
                        log_line(f"{mask_account(username)} 登录失败：{message}", "error")
                    except Exception as exc:  # 单账号异常不影响后续账号
                        result["status"] = "failed"
                        result["message"] = concise_error(exc)
                        log_line(f"{mask_account(username)} 处理失败：{exc}", "error")
                    finally:
                        if context is not None:
                            try:
                                context.close()
                            except Exception:
                                pass

                    with JOB_LOCK:
                        JOB["results"][result_index] = result
                        JOB["done"] = index
                        JOB["current"] = ""

                    if index < len(accounts) and delay > 0:
                        time.sleep(delay)
            finally:
                try:
                    browser.close()
                except Exception:
                    pass
    except Exception as exc:
        log_line(f"任务异常终止：{exc}", "error")
        with JOB_LOCK:
            for result in JOB["results"]:
                if result["status"] == "running":
                    result["status"] = "failed"
                    result["message"] = f"任务中断：{exc}"
    finally:
        with JOB_LOCK:
            JOB["running"] = False
            JOB["current"] = ""
            JOB["finished_at"] = now_ts()
        log_line("任务结束")


# ---------------------------------------------------------------- 成绩查询任务


def reset_score_job(total: int) -> None:
    with SCORE_LOCK:
        SCORE_JOB.update(
            {
                "running": True,
                "total": total,
                "done": 0,
                "current": "",
                "started_at": now_ts(),
                "finished_at": None,
                "results": [],
                "log": [],
            }
        )


def run_score_job(config: dict, accounts: List[dict]) -> None:
    base_url = config["base_url"]
    delay = config["delay"]
    score_log(f"成绩查询开始，共 {len(accounts)} 个账号")

    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=not config["headed"])
            try:
                for index, account in enumerate(accounts, 1):
                    username = account["username"]
                    label = account.get("note") or mask_account(username)
                    result = {
                        "account": mask_account(username),
                        "note": account.get("note", ""),
                        "status": "running",
                        "message": "",
                        "records": [],
                    }
                    with SCORE_LOCK:
                        SCORE_JOB["results"].append(deepcopy(result))
                        result_index = len(SCORE_JOB["results"]) - 1
                        SCORE_JOB["current"] = label

                    score_log(f"[{index}/{len(accounts)}] {label} 开始查询")
                    context = None
                    try:
                        context = browser.new_context(viewport={"width": 1280, "height": 900})
                        page = context.new_page()
                        page.set_default_timeout(30_000)
                        core.login(
                            page,
                            base_url,
                            username,
                            account["password"],
                            config["login_timeout"],
                        )

                        activities = core.fetch_all_activities(
                            context.request, base_url, core.TAB_HISTORY, False, 50
                        )
                        if not activities:
                            result["status"] = "empty"
                            result["message"] = "没有历史活动"
                            score_log(f"{mask_account(username)} 没有历史活动")
                        else:
                            for activity in activities:
                                record = {
                                    "activity": activity.name,
                                    "activity_id": activity.id,
                                }
                                try:
                                    response = context.request.post(
                                        f"{base_url}/api/contestant/activity/"
                                        f"get_score_info?id={activity.id}",
                                        timeout=30_000,
                                    )
                                    if response.status != 200:
                                        record.update(
                                            {
                                                "available": False,
                                                "message": f"HTTP {response.status}",
                                            }
                                        )
                                    else:
                                        record.update(parse_score_response(response.body()))
                                except Exception as exc:
                                    record.update(
                                        {"available": False, "message": f"查询失败：{exc}"}
                                    )
                                result["records"].append(record)

                                if record.get("available"):
                                    score_log(
                                        f"{mask_account(username)} {activity.name}："
                                        f"总分 {format_score(record.get('total', 0))}"
                                    )
                                else:
                                    score_log(
                                        f"{mask_account(username)} {activity.name}："
                                        f"{record.get('message', '暂无成绩信息')}",
                                        "warn",
                                    )

                            if any(r.get("available") for r in result["records"]):
                                result["status"] = "success"
                            else:
                                result["status"] = "empty"
                                result["message"] = "所有历史活动暂无成绩信息"
                    except core.LoginError as exc:
                        message = str(exc).strip()
                        result["status"] = "failed"
                        result["message"] = (
                            message
                            if message.startswith("登录失败")
                            else f"登录失败：{message}"
                        )
                        score_log(f"{mask_account(username)} 登录失败：{message}", "error")
                    except Exception as exc:
                        result["status"] = "failed"
                        result["message"] = concise_error(exc)
                        score_log(f"{mask_account(username)} 查询失败：{exc}", "error")
                    finally:
                        if context is not None:
                            try:
                                context.close()
                            except Exception:
                                pass

                    with SCORE_LOCK:
                        SCORE_JOB["results"][result_index] = result
                        SCORE_JOB["done"] = index
                        SCORE_JOB["current"] = ""

                    if index < len(accounts) and delay > 0:
                        time.sleep(delay)
            finally:
                try:
                    browser.close()
                except Exception:
                    pass
    except Exception as exc:
        score_log(f"成绩查询异常终止：{exc}", "error")
        with SCORE_LOCK:
            for result in SCORE_JOB["results"]:
                if result["status"] == "running":
                    result["status"] = "failed"
                    result["message"] = f"任务中断：{exc}"
    finally:
        with SCORE_LOCK:
            SCORE_JOB["running"] = False
            SCORE_JOB["current"] = ""
            SCORE_JOB["finished_at"] = now_ts()
        score_log("成绩查询结束")


def _score_groups(job: dict) -> List[dict]:
    groups = {}
    order = []
    for result in job.get("results", []):
        for record in result.get("records", []):
            key = record.get("activity_id") or record.get("activity")
            if key not in groups:
                groups[key] = {"activity": record.get("activity", ""), "records": []}
                order.append(key)
            groups[key]["records"].append(
                {
                    "account": result.get("account", ""),
                    "note": result.get("note", ""),
                    **record,
                }
            )
    return [groups[key] for key in order]


def _unique_sheet_name(workbook, base: str) -> str:
    base = re.sub(r"[\[\]:*?/\\]", "_", base or "").strip() or "活动"
    base = base[:31]
    name = base
    index = 2
    while name in workbook.sheetnames:
        suffix = f"({index})"
        name = base[: 31 - len(suffix)] + suffix
        index += 1
    return name


def _style_score_sheet(ws, problem_count: int) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="F2F4F7")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"
    widths = [20, 14, 12, 16] + [12] * problem_count + [10, 24]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width


def build_scores_workbook(job: dict):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    workbook = Workbook()
    summary = workbook.active
    summary.title = "成绩汇总"
    summary.append(
        ["账号", "备注", "活动", "姓名", "准考证号", "题目得分", "总成绩", "成绩备注"]
    )

    groups = _score_groups(job)
    for result in job.get("results", []):
        records = result.get("records") or []
        if not records:
            summary.append(
                [
                    result.get("account", ""),
                    result.get("note", ""),
                    "",
                    "",
                    "",
                    "",
                    "",
                    result.get("message", ""),
                ]
            )
            continue
        for record in records:
            if record.get("available"):
                detail = "；".join(
                    f"{item.get('name', '')} {format_score(item.get('score', 0))}"
                    for item in record.get("problems", [])
                )
                summary.append(
                    [
                        result.get("account", ""),
                        result.get("note", ""),
                        record.get("activity", ""),
                        record.get("name", ""),
                        record.get("ticket_no", ""),
                        detail,
                        record.get("total", 0) / 100,
                        record.get("remark") or "无",
                    ]
                )
            else:
                summary.append(
                    [
                        result.get("account", ""),
                        result.get("note", ""),
                        record.get("activity", ""),
                        "",
                        "",
                        "",
                        "",
                        record.get("message", "暂无成绩信息"),
                    ]
                )

    for row in summary.iter_rows(min_row=2, min_col=7, max_col=7):
        row[0].number_format = "0.##"
    widths = {"A": 20, "B": 14, "C": 26, "D": 12, "E": 16, "F": 40, "G": 10, "H": 24}
    for column, width in widths.items():
        summary.column_dimensions[column].width = width
    summary.freeze_panes = "A2"
    header_fill = PatternFill("solid", fgColor="F2F4F7")
    for cell in summary[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True)

    detail_rows = []
    for result in job.get("results", []):
        for record in result.get("records", []):
            if not record.get("available"):
                continue
            for problem in record.get("problems", []) or []:
                detail_rows.append(
                    [
                        result.get("account", ""),
                        result.get("note", ""),
                        record.get("activity", ""),
                        record.get("name", ""),
                        record.get("ticket_no", ""),
                        problem.get("name", ""),
                        problem.get("score", 0) / 100,
                    ]
                )

    if detail_rows:
        detail_sheet = workbook.create_sheet("题目得分明细")
        detail_sheet.append(
            ["账号", "备注", "活动", "姓名", "准考证号", "题目名称", "得分"]
        )
        for row in detail_rows:
            detail_sheet.append(row)
        for row in detail_sheet.iter_rows(min_row=2, min_col=7, max_col=7):
            row[0].number_format = "0.##"
        widths = {"A": 20, "B": 14, "C": 26, "D": 12, "E": 16, "F": 22, "G": 10}
        for column, width in widths.items():
            detail_sheet.column_dimensions[column].width = width
        detail_sheet.freeze_panes = "A2"
        for cell in detail_sheet[1]:
            cell.fill = header_fill
            cell.font = Font(bold=True)

    for group in groups:
        ws = workbook.create_sheet(_unique_sheet_name(workbook, group["activity"]))
        problems: List[str] = []
        for record in group["records"]:
            for item in record.get("problems", []) or []:
                if item.get("name") not in problems:
                    problems.append(item.get("name"))

        ws.append(["账号", "备注", "姓名", "准考证号"] + problems + ["总成绩", "成绩备注"])
        for record in group["records"]:
            if record.get("available"):
                score_map = {
                    item.get("name"): item.get("score", 0)
                    for item in record.get("problems", [])
                }
                row = [
                    record.get("account", ""),
                    record.get("note", ""),
                    record.get("name", ""),
                    record.get("ticket_no", ""),
                ]
                row.extend(
                    score_map[name] / 100 if name in score_map else "—" for name in problems
                )
                row.append(record.get("total", 0) / 100)
                row.append(record.get("remark") or "无")
            else:
                row = [
                    record.get("account", ""),
                    record.get("note", ""),
                    "",
                    "",
                ] + [""] * len(problems) + ["", record.get("message", "暂无成绩信息")]
            ws.append(row)

        total_column = 5 + len(problems)
        for row in ws.iter_rows(
            min_row=2, min_col=total_column, max_col=total_column
        ):
            row[0].number_format = "0.##"
        _style_score_sheet(ws, len(problems))

    return workbook


# ---------------------------------------------------------------- 证书下载任务


def reset_cert_job(total: int, out_dir: str, layout: str) -> None:
    with CERT_LOCK:
        CERT_JOB.update(
            {
                "running": True,
                "total": total,
                "done": 0,
                "current": "",
                "out_dir": out_dir,
                "layout": layout,
                "started_at": now_ts(),
                "finished_at": None,
                "results": [],
                "log": [],
            }
        )


def run_cert_job(config: dict, accounts: List[dict]) -> None:
    base_url = config["base_url"]
    out_dir = config["out_dir"]
    layout = config["layout"]
    delay = config["delay"]
    cert_root = os.path.join(out_dir, "证书")
    cert_log(f"证书下载开始，共 {len(accounts)} 个账号")

    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=not config["headed"])
            try:
                for index, account in enumerate(accounts, 1):
                    username = account["username"]
                    label = account.get("note") or mask_account(username)
                    result = {
                        "account": mask_account(username),
                        "note": account.get("note", ""),
                        "status": "running",
                        "message": "",
                        "files": [],
                        "skipped": [],
                    }
                    with CERT_LOCK:
                        CERT_JOB["results"].append(deepcopy(result))
                        result_index = len(CERT_JOB["results"]) - 1
                        CERT_JOB["current"] = label

                    cert_log(f"[{index}/{len(accounts)}] {label} 开始下载")
                    context = None
                    try:
                        context = browser.new_context(viewport={"width": 1280, "height": 900})
                        page = context.new_page()
                        page.set_default_timeout(30_000)
                        core.login(
                            page,
                            base_url,
                            username,
                            account["password"],
                            config["login_timeout"],
                        )

                        current = core.fetch_all_activities(
                            context.request, base_url, core.TAB_CURRENT, False, 50
                        )
                        history = core.fetch_all_activities(
                            context.request, base_url, core.TAB_HISTORY, False, 50
                        )
                        activities = []
                        seen = set()
                        for activity in current + history:
                            if activity.id not in seen:
                                seen.add(activity.id)
                                activities.append(activity)

                        if not activities:
                            result["status"] = "empty"
                            result["message"] = "没有找到活动"
                            cert_log(f"{mask_account(username)} 没有活动")
                        else:
                            user_tag = account_folder(account)
                            for activity in activities:
                                for cert_key in config["types"]:
                                    cert_label, _ = CERTIFICATE_TYPES[cert_key]
                                    folder = resolve_target(
                                        layout, cert_root, user_tag, activity
                                    )
                                    try:
                                        os.makedirs(folder, exist_ok=True)
                                    except OSError as exc:
                                        result["skipped"].append(
                                            {
                                                "activity": activity.name,
                                                "type": cert_label,
                                                "reason": f"无法创建目录：{exc}",
                                            }
                                        )
                                        continue

                                    path, reason = download_certificate(
                                        context.request,
                                        base_url,
                                        activity,
                                        cert_key,
                                        folder,
                                    )
                                    if path:
                                        result["files"].append(
                                            {
                                                "name": os.path.basename(path),
                                                "path": path,
                                                "activity": activity.name,
                                                "type": cert_label,
                                            }
                                        )
                                        cert_log(
                                            f"{mask_account(username)} 已保存 "
                                            f"{os.path.basename(path)}"
                                        )
                                    else:
                                        result["skipped"].append(
                                            {
                                                "activity": activity.name,
                                                "type": cert_label,
                                                "reason": reason,
                                            }
                                        )
                                        cert_log(
                                            f"{mask_account(username)} 跳过"
                                            f"「{activity.name}」{cert_label}：{reason}",
                                            "warn",
                                        )
                                        if folder != cert_root:
                                            try:
                                                if not os.listdir(folder):
                                                    os.rmdir(folder)
                                            except OSError:
                                                pass

                            if result["files"] and result["skipped"]:
                                result["status"] = "partial"
                            elif result["files"]:
                                result["status"] = "success"
                            else:
                                result["status"] = "skipped"
                                result["message"] = "没有可下载的证书"
                    except core.LoginError as exc:
                        message = str(exc).strip()
                        result["status"] = "failed"
                        result["message"] = (
                            message
                            if message.startswith("登录失败")
                            else f"登录失败：{message}"
                        )
                        cert_log(f"{mask_account(username)} 登录失败：{message}", "error")
                    except Exception as exc:
                        result["status"] = "failed"
                        result["message"] = concise_error(exc)
                        cert_log(f"{mask_account(username)} 下载失败：{exc}", "error")
                    finally:
                        if context is not None:
                            try:
                                context.close()
                            except Exception:
                                pass

                    with CERT_LOCK:
                        CERT_JOB["results"][result_index] = result
                        CERT_JOB["done"] = index
                        CERT_JOB["current"] = ""

                    if index < len(accounts) and delay > 0:
                        time.sleep(delay)
            finally:
                try:
                    browser.close()
                except Exception:
                    pass
    except Exception as exc:
        cert_log(f"证书下载异常终止：{exc}", "error")
        with CERT_LOCK:
            for result in CERT_JOB["results"]:
                if result["status"] == "running":
                    result["status"] = "failed"
                    result["message"] = f"任务中断：{exc}"
    finally:
        with CERT_LOCK:
            CERT_JOB["running"] = False
            CERT_JOB["current"] = ""
            CERT_JOB["finished_at"] = now_ts()
        cert_log("证书下载结束")


@app.post("/api/start")
def api_start():
    payload = request.get_json(silent=True) or {}

    with JOB_LOCK:
        if JOB["running"]:
            return jsonify({"ok": False, "message": "已有任务正在执行，请等待完成"}), 409
    with SCORE_LOCK:
        if SCORE_JOB["running"]:
            return jsonify({"ok": False, "message": "成绩查询正在执行，请等待完成"}), 409
    with CERT_LOCK:
        if CERT_JOB["running"]:
            return jsonify({"ok": False, "message": "证书下载正在执行，请等待完成"}), 409

    with ACCOUNT_LOCK:
        accounts = [deepcopy(account) for account in ACCOUNTS]
    if not accounts:
        return jsonify({"ok": False, "message": "账号列表为空，请先添加账号"}), 400

    raw_out = str(payload.get("out_dir") or "tickets").strip() or "tickets"
    out_dir = os.path.abspath(os.path.expanduser(raw_out))
    try:
        os.makedirs(out_dir, exist_ok=True)
        probe = os.path.join(out_dir, ".write_test")
        with open(probe, "w", encoding="utf-8") as fh:
            fh.write("ok")
        os.remove(probe)
    except OSError as exc:
        return jsonify({"ok": False, "message": f"下载目录不可用：{exc}"}), 400

    base_url = str(payload.get("base_url") or core.DEFAULT_BASE_URL).strip().rstrip("/")
    if not re.match(r"^https?://[^\s/]+", base_url):
        return jsonify({"ok": False, "message": "站点地址格式不正确（需以 http:// 或 https:// 开头）"}), 400

    layout = str(payload.get("layout") or LAYOUT_ACTIVITY).strip().lower()
    if layout not in LAYOUTS:
        return jsonify({"ok": False, "message": "目录结构参数不正确"}), 400

    try:
        filename_template = normalize_filename_template(str(payload.get("filename_template") or ""))
    except ValueError as exc:
        return jsonify({"ok": False, "message": str(exc)}), 400

    try:
        login_timeout = max(10, min(300, int(payload.get("login_timeout") or 60)))
    except (TypeError, ValueError):
        login_timeout = 60
    try:
        delay = max(0.0, min(10.0, float(payload.get("delay", 1))))
    except (TypeError, ValueError):
        delay = 1.0

    config = {
        "base_url": base_url,
        "out_dir": out_dir,
        "layout": layout,
        "filename_template": filename_template,
        "headed": bool(payload.get("headed")),
        "history": bool(payload.get("history")),
        "login_timeout": login_timeout,
        "delay": delay,
    }

    reset_job(len(accounts), out_dir, layout, filename_template)
    thread = threading.Thread(target=run_job, args=(config, accounts), daemon=True)
    thread.start()

    return jsonify({"ok": True, "message": "任务已开始", "total": len(accounts)})


@app.get("/api/job")
def api_job():
    with JOB_LOCK:
        return jsonify({"ok": True, "job": deepcopy(JOB)})


@app.post("/api/scores/start")
def api_scores_start():
    payload = request.get_json(silent=True) or {}

    with JOB_LOCK:
        if JOB["running"]:
            return jsonify({"ok": False, "message": "下载任务正在执行，请等待完成"}), 409
    with SCORE_LOCK:
        if SCORE_JOB["running"]:
            return jsonify({"ok": False, "message": "成绩查询正在执行，请等待完成"}), 409
    with CERT_LOCK:
        if CERT_JOB["running"]:
            return jsonify({"ok": False, "message": "证书下载正在执行，请等待完成"}), 409

    with ACCOUNT_LOCK:
        accounts = [deepcopy(account) for account in ACCOUNTS]
    if not accounts:
        return jsonify({"ok": False, "message": "账号列表为空，请先添加账号"}), 400

    base_url = str(payload.get("base_url") or core.DEFAULT_BASE_URL).strip().rstrip("/")
    if not re.match(r"^https?://[^\s/]+", base_url):
        return jsonify(
            {"ok": False, "message": "站点地址格式不正确（需以 http:// 或 https:// 开头）"}
        ), 400

    try:
        login_timeout = max(10, min(300, int(payload.get("login_timeout") or 60)))
    except (TypeError, ValueError):
        login_timeout = 60
    try:
        delay = max(0.0, min(10.0, float(payload.get("delay", 1))))
    except (TypeError, ValueError):
        delay = 1.0

    config = {
        "base_url": base_url,
        "headed": bool(payload.get("headed")),
        "login_timeout": login_timeout,
        "delay": delay,
    }

    reset_score_job(len(accounts))
    thread = threading.Thread(target=run_score_job, args=(config, accounts), daemon=True)
    thread.start()
    return jsonify({"ok": True, "message": "成绩查询已开始", "total": len(accounts)})


@app.get("/api/scores")
def api_scores():
    with SCORE_LOCK:
        return jsonify({"ok": True, "job": deepcopy(SCORE_JOB)})


@app.get("/api/scores/export.xlsx")
def api_scores_export():
    with SCORE_LOCK:
        job = deepcopy(SCORE_JOB)
    record_count = sum(len(result.get("records") or []) for result in job.get("results", []))
    if not record_count:
        return jsonify({"ok": False, "message": "暂无可导出的成绩数据，请先查询"}), 404

    workbook = build_scores_workbook(job)
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    filename = f"历史成绩_{time.strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


@app.post("/api/certificates/start")
def api_certificates_start():
    payload = request.get_json(silent=True) or {}

    with JOB_LOCK:
        if JOB["running"]:
            return jsonify({"ok": False, "message": "下载任务正在执行，请等待完成"}), 409
    with SCORE_LOCK:
        if SCORE_JOB["running"]:
            return jsonify({"ok": False, "message": "成绩查询正在执行，请等待完成"}), 409
    with CERT_LOCK:
        if CERT_JOB["running"]:
            return jsonify({"ok": False, "message": "证书下载正在执行，请等待完成"}), 409

    with ACCOUNT_LOCK:
        accounts = [deepcopy(account) for account in ACCOUNTS]
    if not accounts:
        return jsonify({"ok": False, "message": "账号列表为空，请先添加账号"}), 400

    raw_out = str(payload.get("out_dir") or "tickets").strip() or "tickets"
    out_dir = os.path.abspath(os.path.expanduser(raw_out))
    try:
        os.makedirs(out_dir, exist_ok=True)
        probe = os.path.join(out_dir, ".write_test")
        with open(probe, "w", encoding="utf-8") as fh:
            fh.write("ok")
        os.remove(probe)
    except OSError as exc:
        return jsonify({"ok": False, "message": f"下载目录不可用：{exc}"}), 400

    try:
        os.makedirs(os.path.join(out_dir, "证书"), exist_ok=True)
    except OSError:
        pass

    base_url = str(payload.get("base_url") or core.DEFAULT_BASE_URL).strip().rstrip("/")
    if not re.match(r"^https?://[^\s/]+", base_url):
        return jsonify(
            {"ok": False, "message": "站点地址格式不正确（需以 http:// 或 https:// 开头）"}
        ), 400

    layout = str(payload.get("layout") or LAYOUT_ACTIVITY).strip().lower()
    if layout not in LAYOUTS:
        layout = LAYOUT_ACTIVITY

    raw_types = payload.get("types")
    if isinstance(raw_types, str):
        raw_types = [item.strip() for item in raw_types.split(",") if item.strip()]
    if not isinstance(raw_types, list):
        raw_types = ["prize"]
    types = []
    for item in raw_types:
        key = str(item).strip().lower()
        if key in CERTIFICATE_TYPES and key not in types:
            types.append(key)
    if not types:
        return jsonify({"ok": False, "message": "请至少选择一种证书类型"}), 400

    try:
        login_timeout = max(10, min(300, int(payload.get("login_timeout") or 60)))
    except (TypeError, ValueError):
        login_timeout = 60
    try:
        delay = max(0.0, min(10.0, float(payload.get("delay", 1))))
    except (TypeError, ValueError):
        delay = 1.0

    config = {
        "base_url": base_url,
        "out_dir": out_dir,
        "layout": layout,
        "types": types,
        "headed": bool(payload.get("headed")),
        "login_timeout": login_timeout,
        "delay": delay,
    }

    reset_cert_job(len(accounts), out_dir, layout)
    thread = threading.Thread(target=run_cert_job, args=(config, accounts), daemon=True)
    thread.start()
    return jsonify({"ok": True, "message": "证书下载已开始", "total": len(accounts)})


@app.get("/api/certificates")
def api_certificates():
    with CERT_LOCK:
        return jsonify({"ok": True, "job": deepcopy(CERT_JOB)})


# ---------------------------------------------------------------- 目录操作


@app.post("/api/open_dir")
def api_open_dir():
    payload = request.get_json(silent=True) or {}
    path = str(payload.get("path") or "").strip()
    if not path:
        with JOB_LOCK:
            path = JOB.get("out_dir", "")
    path = os.path.abspath(os.path.expanduser(path))
    if not path or not os.path.isdir(path):
        return jsonify({"ok": False, "message": "目录不存在"}), 404
    try:
        if os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as exc:
        return jsonify({"ok": False, "message": f"无法打开目录：{exc}"}), 500
    return jsonify({"ok": True, "path": path})


@app.post("/api/pick_dir")
def api_pick_dir():
    """弹出系统文件夹选择框（仅 Windows / macOS 支持）。"""
    if os.name == "nt":
        script = (
            "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
            "Add-Type -AssemblyName System.Windows.Forms; "
            "$dialog = New-Object System.Windows.Forms.FolderBrowserDialog; "
            "$dialog.Description = '选择准考证下载目录'; "
            "$dialog.ShowNewFolderButton = $true; "
            "if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) "
            "{ [Console]::Out.Write($dialog.SelectedPath) }"
        )
        try:
            completed = subprocess.run(
                ["powershell", "-NoProfile", "-STA", "-Command", script],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=300,
            )
            path = (completed.stdout or "").strip()
            if path:
                return jsonify({"ok": True, "path": path})
            return jsonify({"ok": False, "message": "未选择目录"})
        except Exception as exc:
            return jsonify({"ok": False, "message": f"无法打开目录选择框：{exc}"}), 500

    return jsonify({"ok": False, "message": "当前系统不支持弹出选择框，请手动输入路径"}), 400


# ---------------------------------------------------------------- 页面


@app.get("/")
def index():
    return render_template("index.html")


@app.errorhandler(413)
def too_large(_):
    return jsonify({"ok": False, "message": "文件过大（上限 8MB）"}), 413


def find_free_port(preferred: int) -> int:
    for port in range(preferred, preferred + 30):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise OSError("找不到可用端口")


def main() -> None:
    parser = argparse.ArgumentParser(description="NOI 准考证批量下载 —— 网页版")
    parser.add_argument("--host", default="127.0.0.1", help="监听地址，默认 127.0.0.1")
    parser.add_argument("--port", type=int, default=8765, help="端口，默认 8765（被占用时自动顺延）")
    parser.add_argument("--no-browser", action="store_true", help="启动时不自动打开浏览器")
    parser.add_argument("--debug", action="store_true", help="启用 Flask 调试模式")
    args = parser.parse_args()

    if args.host in ("127.0.0.1", "localhost") and not args.debug:
        port = find_free_port(args.port)
    else:
        port = args.port

    url = f"http://127.0.0.1:{port}/"
    print("=" * 56)
    print("  NOI 准考证批量下载 · 网页版")
    print(f"  访问地址：{url}")
    print("  按 Ctrl+C 停止服务")
    print("=" * 56)

    if not args.no_browser:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    app.run(host=args.host, port=port, debug=args.debug, use_reloader=False, threaded=True)


if __name__ == "__main__":
    main()
