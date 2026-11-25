#!/usr/bin/env sh
# NOI 准考证批量下载 · 网页版启动脚本（macOS / Linux，Windows 的 Git Bash 也可用）
set -u

cd "$(dirname "$0")" || exit 1

echo "========================================"
echo "  NOI 准考证批量下载 · 网页版"
echo "========================================"

# ---------- 选择 Python（优先选已安装依赖的解释器） ----------
PY=""
FALLBACK=""
for candidate in python3 python; do
    command -v "$candidate" >/dev/null 2>&1 || continue
    if ! "$candidate" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 8) else 1)' 2>/dev/null; then
        continue
    fi
    [ -z "$FALLBACK" ] && FALLBACK="$candidate"
    if "$candidate" -c 'import flask, openpyxl, playwright' >/dev/null 2>&1; then
        PY="$candidate"
        break
    fi
done

if [ -z "$PY" ]; then
    if [ -n "$FALLBACK" ]; then
        PY="$FALLBACK"
    else
        echo "错误：未找到 Python 3.8 及以上版本，请先安装。" >&2
        echo "下载地址：https://www.python.org/downloads/" >&2
        exit 1
    fi
fi

echo "使用 Python：$PY（$("$PY" -V 2>&1)）"

install_deps() {
    echo "正在安装 Python 依赖……"
    "$PY" -m pip install -r requirements.txt || return 1
    echo "正在安装 Chromium 内核（首次约需下载 150MB）……"
    "$PY" -m playwright install chromium || return 1
    return 0
}

# ---------- 检查依赖 ----------
if ! "$PY" -c 'import flask, openpyxl, playwright' >/dev/null 2>&1; then
    echo "缺少依赖：flask / openpyxl / playwright"
    if [ -t 0 ]; then
        printf "是否现在安装？[y/N] "
        read -r answer
        case "$answer" in
            y|Y)
                if ! install_deps; then
                    echo "依赖安装失败，请检查网络后重试。" >&2
                    exit 1
                fi
                ;;
            *)
                echo "已取消。请手动执行："
                echo "  $PY -m pip install -r requirements.txt"
                echo "  $PY -m playwright install chromium"
                exit 1
                ;;
        esac
    else
        echo "请先执行："
        echo "  $PY -m pip install -r requirements.txt"
        echo "  $PY -m playwright install chromium"
        exit 1
    fi
fi

# ---------- 检查 Chromium 内核 ----------
if ! "$PY" -c '
import os
import sys

from playwright.sync_api import sync_playwright

try:
    with sync_playwright() as p:
        path = p.chromium.executable_path
except Exception:
    sys.exit(1)

sys.exit(0 if path and os.path.exists(path) else 1)
' >/dev/null 2>&1; then
    echo "未检测到 Playwright Chromium 内核。"
    if [ -t 0 ]; then
        printf "是否现在安装？[y/N] "
        read -r answer
        case "$answer" in
            y|Y)
                if ! "$PY" -m playwright install chromium; then
                    echo "Chromium 内核安装失败。" >&2
                    exit 1
                fi
                ;;
            *)
                echo "已取消。请手动执行：$PY -m playwright install chromium"
                exit 1
                ;;
        esac
    else
        echo "请先执行：$PY -m playwright install chromium"
        exit 1
    fi
fi

# ---------- 启动 ----------
# 统一 UTF-8 输出，避免中文日志在终端中乱码
PYTHONIOENCODING=utf-8
export PYTHONIOENCODING

exec "$PY" web_app.py "$@"
