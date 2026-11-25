#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NOI 在线报名系统 —— 登录并批量下载准考证

流程：
  1. 用 Playwright 打开登录页，填写账号密码并登录（获取会话 Cookie）；
  2. 调用活动列表接口，遍历「当前活动」（可选「历史活动」）；
  3. 对每个活动调用 download_ticket 接口，把准考证 PDF 保存到本地。

用法：
  python download_tickets.py -u 你的账号 -p '密码'
  python download_tickets.py            # 交互式输入账号密码
  python download_tickets.py --history  # 同时遍历历史活动
  python download_tickets.py --headed   # 显示浏览器窗口（出现验证码时可手动完成）

依赖：
  pip install playwright
  python -m playwright install chromium
"""

from __future__ import annotations

import argparse
import email.header
import getpass
import os
import re
import sys
import time
import urllib.parse
from dataclasses import dataclass, field
from typing import Callable, List, Optional, Tuple

from playwright.sync_api import APIRequestContext, Page, sync_playwright

# ---------------------------------------------------------------- 常量与协议

DEFAULT_BASE_URL = "https://cspsjtest.noi.cn"

LOGIN_URL_PATH = "/"
LIST_API_PATH = "/api/contestant/activity/list"
TICKET_API_PATH = "/api/contestant/activity/download_ticket"

# 站点使用自定义二进制协议：请求体以「类哈希」开头，随后按字段顺序写入
# 无符号 varint / 长度前缀字符串。
LIST_REQUEST_CLASS = 397453382      # 活动列表请求类
ACTIVITY_LIST_RESP_CLASS = 877769905  # 活动列表响应类
SESSION_EXPIRED_CLASS = 1811961344  # 会话过期响应类

TAB_CURRENT = 0   # 当前活动
TAB_HISTORY = 1   # 历史活动

MAX_PAGES = 100


class ProtocolError(RuntimeError):
    """接口返回的数据无法按预期协议解析。"""


class LoginError(RuntimeError):
    """登录失败。"""


# ---------------------------------------------------------------- varint 编解码


def encode_varint(value: int) -> bytes:
    if value < 0:
        raise ValueError("varint 不支持负数")
    out = bytearray()
    while True:
        byte = value & 0x7F
        value >>= 7
        if value:
            out.append(byte | 0x80)
        else:
            out.append(byte)
            return bytes(out)


class Reader:
    """读取站点二进制协议的游标。"""

    def __init__(self, data: bytes):
        self.data = data
        self.pos = 0

    def varint(self) -> int:
        value = 0
        shift = 0
        while True:
            if self.pos >= len(self.data):
                raise ProtocolError("数据在 varint 中途结束")
            byte = self.data[self.pos]
            self.pos += 1
            value |= (byte & 0x7F) << shift
            if not byte & 0x80:
                return value
            shift += 7
            if shift > 63:
                raise ProtocolError("varint 过长")

    def string(self) -> str:
        length = self.varint()
        end = self.pos + length
        if end > len(self.data):
            raise ProtocolError("字符串长度越界")
        text = self.data[self.pos:end].decode("utf-8")
        self.pos = end
        return text

    def at_end(self) -> bool:
        return self.pos >= len(self.data)


# ---------------------------------------------------------------- 数据结构


@dataclass
class Activity:
    id: int
    name: str
    code: str
    start: str
    end: str
    enroll_start: str
    enroll_end: str
    kind: int
    tab: str = "当前活动"


# ---------------------------------------------------------------- 协议解析


def parse_activity_list(data: bytes) -> Tuple[int, List[Activity]]:
    """解析活动列表响应，返回 (总数, 本页活动)。"""
    reader = Reader(data)
    magic = reader.varint()
    if magic == SESSION_EXPIRED_CLASS:
        raise LoginError("会话已过期，请重新登录")
    if magic != ACTIVITY_LIST_RESP_CLASS:
        raise ProtocolError(f"活动列表响应头不匹配：{magic}")

    total = reader.varint()
    count = reader.varint()

    activities: List[Activity] = []
    for _ in range(count):
        activity = Activity(
            id=reader.varint(),
            name=reader.string(),
            code=reader.string(),
            start=reader.string(),
            end=reader.string(),
            enroll_start=reader.string(),
            enroll_end=reader.string(),
            kind=reader.varint(),
        )
        activities.append(activity)
    return total, activities


def decode_message(data: bytes) -> Optional[str]:
    """解析站点协议返回的错误消息（类哈希 + 长度 + UTF-8 文本）。"""
    try:
        reader = Reader(data)
        reader.varint()
        return reader.string()
    except Exception:
        return None


# ---------------------------------------------------------------- 网络操作


def build_list_request(tab: int, only_enrolled: bool, page_size: int, page_index: int) -> bytes:
    return (
        encode_varint(LIST_REQUEST_CLASS)
        + bytes([tab, 1 if only_enrolled else 0])
        + encode_varint(page_size)
        + encode_varint(page_index)
    )


def fetch_activity_page(
    request: APIRequestContext,
    base_url: str,
    tab: int,
    only_enrolled: bool,
    page_size: int,
    page_index: int,
) -> Tuple[int, List[Activity]]:
    body = build_list_request(tab, only_enrolled, page_size, page_index)
    response = request.post(
        base_url + LIST_API_PATH,
        data=body,
        headers={"Content-Type": "application/octet-stream"},
        timeout=30_000,
    )
    if response.status != 200:
        raise ProtocolError(f"活动列表接口返回 HTTP {response.status}")
    return parse_activity_list(response.body())


def fetch_all_activities(
    request: APIRequestContext,
    base_url: str,
    tab: int,
    only_enrolled: bool,
    page_size: int,
) -> List[Activity]:
    collected: List[Activity] = []
    seen = set()
    page_index = 0

    while page_index < MAX_PAGES:
        total, batch = fetch_activity_page(
            request, base_url, tab, only_enrolled, page_size, page_index
        )
        new_items = [a for a in batch if a.id not in seen]
        for activity in new_items:
            seen.add(activity.id)
            collected.append(activity)

        if not batch:
            break
        if total is not None and len(collected) >= total:
            break
        if not new_items:
            break  # 服务端忽略了分页参数，避免死循环
        page_index += 1

    return collected


# ---------------------------------------------------------------- 文件名处理


def sanitize_filename(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", name)
    name = name.strip().strip(".")
    return name[:120] or "未命名"


def filename_from_disposition(disposition: Optional[str]) -> Optional[str]:
    if not disposition:
        return None

    match = re.search(r"filename\*\s*=\s*UTF-8''([^;]+)", disposition, re.IGNORECASE)
    if match:
        return urllib.parse.unquote(match.group(1).strip().strip('"'))

    match = re.search(r'filename\s*=\s*("?)(.+?)\1(?:\s*;|$)', disposition, re.IGNORECASE)
    if match:
        raw = match.group(2)
        try:
            return str(email.header.make_header(email.header.decode_header(raw)))
        except Exception:
            return raw
    return None


def unique_path(directory: str, filename: str) -> str:
    path = os.path.join(directory, filename)
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(filename)
    index = 2
    while True:
        candidate = os.path.join(directory, f"{stem} ({index}){ext}")
        if not os.path.exists(candidate):
            return candidate
        index += 1


# ---------------------------------------------------------------- 登录


def login(page: Page, base_url: str, user: str, password: str, timeout_seconds: int) -> None:
    page.goto(base_url + LOGIN_URL_PATH, wait_until="domcontentloaded", timeout=60_000)

    # 若已有会话，直接进入报名系统
    if "contestant.html" in page.url:
        return

    page.wait_for_selector('input[placeholder="用户名/邮箱/手机号"]', timeout=30_000)

    login_errors: List[str] = []

    def on_dialog(dialog) -> None:
        login_errors.append(dialog.message)
        try:
            dialog.accept()
        except Exception:
            pass

    page.on("dialog", on_dialog)
    try:
        page.fill('input[placeholder="用户名/邮箱/手机号"]', user)
        page.fill('input[placeholder="密码"]', password)
        page.locator('button.button-primary:has-text("登录")').last.click()

        deadline = time.time() + timeout_seconds
        while time.time() < deadline:
            if login_errors:
                raise LoginError(login_errors[0])
            if "contestant.html" in page.url:
                try:
                    page.wait_for_load_state("networkidle", timeout=10_000)
                except Exception:
                    pass
                return
            page.wait_for_timeout(200)

        raise LoginError(
            "登录超时。请确认账号密码正确、网络正常；"
            "若页面出现图形/滑块验证码，请使用 --headed 参数在浏览器窗口中手动完成。"
        )
    finally:
        page.remove_listener("dialog", on_dialog)


# ---------------------------------------------------------------- 下载


def build_ticket_filename(
    activity: Activity,
    original_filename: Optional[str],
    filename_prefix: Optional[str] = None,
) -> str:
    """生成准考证文件名。

    filename_prefix 为 None 时默认使用活动名称作前缀；
    传入空字符串则不加前缀。
    """
    if filename_prefix is None:
        prefix = sanitize_filename(activity.name)
    elif filename_prefix == "":
        prefix = ""
    else:
        prefix = sanitize_filename(filename_prefix)

    if original_filename:
        original = sanitize_filename(original_filename)
        filename = f"{prefix}_{original}" if prefix else original
    else:
        filename = f"{prefix}_准考证_{activity.id}.pdf" if prefix else f"准考证_{activity.id}.pdf"

    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"
    return filename


def download_ticket(
    request: APIRequestContext,
    base_url: str,
    activity: Activity,
    out_dir: str,
    filename_prefix: Optional[str] = None,
    filename_builder: Optional[Callable[[Activity, Optional[str]], str]] = None,
) -> Tuple[Optional[str], Optional[str]]:
    """下载单个活动的准考证，返回 (保存路径, 失败原因)。

    filename_prefix 控制文件名前缀，默认使用活动名称。
    filename_builder 优先于 filename_prefix，用于自定义文件名（网页版模板）。
    """
    url = f"{base_url}{TICKET_API_PATH}?id={activity.id}"
    response = request.get(url, timeout=60_000)

    if response.status != 200:
        return None, f"HTTP {response.status}"

    data = response.body()
    content_type = (response.headers.get("content-type") or "").split(";")[0].strip().lower()

    if data[:5] == b"%PDF-" or content_type == "application/pdf":
        original = filename_from_disposition(response.headers.get("content-disposition"))
        if filename_builder is not None:
            filename = sanitize_filename(filename_builder(activity, original))
            if not filename.lower().endswith(".pdf"):
                filename += ".pdf"
        else:
            filename = build_ticket_filename(activity, original, filename_prefix)

        path = unique_path(out_dir, filename)
        with open(path, "wb") as fh:
            fh.write(data)
        return path, None

    message = decode_message(data)
    if message:
        return None, message
    return None, f"返回内容不是 PDF（content-type={content_type or '未知'}，{len(data)} 字节）"


# ---------------------------------------------------------------- 主流程


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="登录 NOI 在线报名系统，遍历活动列表并下载准考证 PDF。",
    )
    parser.add_argument("-u", "--user", help="用户名 / 手机号 / 邮箱")
    parser.add_argument("-p", "--password", help="密码（不填则交互式输入）")
    parser.add_argument(
        "--base-url",
        default=DEFAULT_BASE_URL,
        help=f"站点地址，默认 {DEFAULT_BASE_URL}",
    )
    parser.add_argument("-o", "--out", default="tickets", help="准考证保存目录，默认 ./tickets")
    parser.add_argument("--history", action="store_true", help="同时遍历「历史活动」")
    parser.add_argument("--only-enrolled", action="store_true", help="仅处理已报名的活动")
    parser.add_argument("--page-size", type=int, default=50, help="每页活动数，默认 50")
    parser.add_argument("--headed", action="store_true", help="显示浏览器窗口，便于手动完成验证码")
    parser.add_argument("--login-timeout", type=int, default=60, help="登录等待秒数，默认 60")
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv)
    base_url = args.base_url.rstrip("/")

    user = (args.user or input("账号（用户名/手机号/邮箱）：")).strip()
    password = args.password or getpass.getpass("密码：")
    if not user or not password:
        print("错误：账号或密码为空。", file=sys.stderr)
        return 1

    out_dir = os.path.abspath(args.out)
    os.makedirs(out_dir, exist_ok=True)

    print(f"站点：{base_url}")
    print(f"保存目录：{out_dir}")

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=not args.headed)
        context = browser.new_context(viewport={"width": 1280, "height": 900})
        page = context.new_page()
        page.set_default_timeout(30_000)

        try:
            print("正在登录 ……", flush=True)
            login(page, base_url, user, password, args.login_timeout)
            print("登录成功。", flush=True)

            tabs = [(TAB_CURRENT, "当前活动")]
            if args.history:
                tabs.append((TAB_HISTORY, "历史活动"))

            activities: List[Activity] = []
            for tab, label in tabs:
                print(f"读取{label}列表 ……", flush=True)
                batch = fetch_all_activities(
                    context.request,
                    base_url,
                    tab,
                    args.only_enrolled,
                    args.page_size,
                )
                for activity in batch:
                    activity.tab = label
                activities.extend(batch)
                print(f"  {label}：{len(batch)} 个", flush=True)

            if not activities:
                print("没有找到任何活动。")
                return 0

            print("\n活动列表：")
            print(f"  {'序号':<4}{'活动名称':<24}{'活动编码':<18}{'开始时间':<21}编号")
            for index, activity in enumerate(activities, 1):
                print(
                    f"  {index:<4}{activity.name:<24}{activity.code:<18}"
                    f"{activity.start:<21}{activity.id}"
                )
            print()

            downloaded: List[str] = []
            skipped: List[Tuple[Activity, str]] = []

            for index, activity in enumerate(activities, 1):
                print(f"[{index}/{len(activities)}] {activity.name}（{activity.code}）…… ", end="", flush=True)
                try:
                    path, reason = download_ticket(context.request, base_url, activity, out_dir)
                except Exception as exc:  # 网络等异常不应中断整批任务
                    path, reason = None, f"下载异常：{exc}"

                if path:
                    downloaded.append(path)
                    print(f"已保存 -> {os.path.basename(path)}", flush=True)
                else:
                    skipped.append((activity, reason or "未知原因"))
                    print(f"跳过（{reason}）", flush=True)

            print("\n========== 结果 ==========")
            print(f"成功下载：{len(downloaded)} 个")
            print(f"跳过：{len(skipped)} 个")
            if skipped:
                print("\n跳过明细：")
                for activity, reason in skipped:
                    print(f"  - {activity.name}（{activity.code}，id={activity.id}）：{reason}")
            return 0

        except LoginError as exc:
            print(f"\n登录失败：{exc}", file=sys.stderr)
            return 2
        except ProtocolError as exc:
            print(f"\n接口解析失败：{exc}", file=sys.stderr)
            return 3
        except Exception as exc:
            print(f"\n执行失败：{exc}", file=sys.stderr)
            return 1
        finally:
            browser.close()


if __name__ == "__main__":
    sys.exit(main())
