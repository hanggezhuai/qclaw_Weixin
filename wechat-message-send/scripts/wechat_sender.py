#!/usr/bin/env python3
"""Stable WeChat sender via Windows UI automation.

This script intentionally avoids WeChat injection and instead automates:
- focusing the main WeChat window
- optionally searching a conversation by a provided keyword
- clicking the message input area
- pasting the message from clipboard
- pressing Enter to send

It is best suited for:
1. sending to the currently opened chat (`--current`)
2. batch sending with a manual id -> search keyword mapping file
"""

from __future__ import annotations

import argparse
from difflib import SequenceMatcher
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import time
from typing import Dict, List, Optional, Sequence

import ctypes as ct
from ctypes import wintypes as wt


SW_RESTORE = 9
VK_CONTROL = 0x11
VK_C = 0x43
VK_F = 0x46
VK_V = 0x56
VK_A = 0x41
VK_BACK = 0x08
VK_DOWN = 0x28
VK_RETURN = 0x0D
VK_MENU = 0x12
VK_ESCAPE = 0x1B
KEYEVENTF_KEYUP = 0x0002

MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004

CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002
SRCCOPY = 0x00CC0020
BI_RGB = 0
DIB_RGB_COLORS = 0
PW_RENDERFULLCONTENT = 0x00000002
MIN_HEADER_CHANGE = 6.0
MIN_BODY_CHANGE = 5.0
SEARCH_BOX_X_RATIO = 0.215
SEARCH_BOX_Y_RATIO = 0.08
SEARCH_BOX_REGION = (0.05, 0.02, 0.36, 0.12)
SEARCH_RESULT_X_RATIO = 0.26
SEARCH_RESULT_Y_RATIO = 0.20
# Personal WeChat: search hits render in the LEFT column under the search box.
# A single region starting at ~0.37 width was missing the list (caused not_found).
# Wide enough to include middle-column search hits on some WeChat 3.x layouts; dedupe happens after OCR.
SEARCH_RESULTS_REGIONS = ((0.02, 0.07, 0.52, 0.85),)
# Title bar can sit over the chat pane; try two horizontal bands.
TITLE_REGIONS = (
    (0.28, 0.00, 0.38, 0.14),
    (0.50, 0.00, 0.46, 0.14),
)
# Merge OCR boxes into one list row if vertical gap from previous box bottom is small.
RESULT_GROUP_ROW_SEP = 20
SEARCH_SETTLE_SEC = 1.05
SKIP_TOP_PIXELS_IN_RESULTS = 22
# If OCR row aligns but click hits the wrong list row, retry with small Y nudges (client px).
TITLE_CLICK_Y_OFFSETS = (0, 10, -8, 16, -14)

OCR_STOPWORDS = {
    "联系人",
    "群聊",
    "聊天记录",
    "进入聊",
    "搜索",
}
OCR_STOPWORDS_NORMALIZED = {re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "", item).lower() for item in OCR_STOPWORDS}
CONTACT_SECTION_MARKERS = {"联系人", "群聊"}
CONTACT_SECTION_MARKERS_NORMALIZED = {
    re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "", item).lower() for item in CONTACT_SECTION_MARKERS
}
NETWORK_RESULT_MARKERS = (
    "搜索网络结果",
    "最近在搜",
    "显示全部",
    "公众号",
    "AI搜索",
    "听一听",
    "表情",
    "视频号",
)
OCR_CONTEXT_PREFIXES = (
    "工作时间",
    "包含",
    "聊天记录",
)
OCR_ENGINE = None
UIA_DESKTOP = None

user32 = ct.WinDLL("user32", use_last_error=True)
kernel32 = ct.WinDLL("kernel32", use_last_error=True)
gdi32 = ct.WinDLL("gdi32", use_last_error=True)


class BITMAPINFOHEADER(ct.Structure):
    _fields_ = [
        ("biSize", wt.DWORD),
        ("biWidth", ct.c_long),
        ("biHeight", ct.c_long),
        ("biPlanes", wt.WORD),
        ("biBitCount", wt.WORD),
        ("biCompression", wt.DWORD),
        ("biSizeImage", wt.DWORD),
        ("biXPelsPerMeter", ct.c_long),
        ("biYPelsPerMeter", ct.c_long),
        ("biClrUsed", wt.DWORD),
        ("biClrImportant", wt.DWORD),
    ]


class BITMAPINFO(ct.Structure):
    _fields_ = [
        ("bmiHeader", BITMAPINFOHEADER),
        ("bmiColors", wt.DWORD * 3),
    ]

kernel32.GlobalAlloc.argtypes = [wt.UINT, ct.c_size_t]
kernel32.GlobalAlloc.restype = wt.HGLOBAL
kernel32.GlobalLock.argtypes = [wt.HGLOBAL]
kernel32.GlobalLock.restype = ct.c_void_p
kernel32.GlobalUnlock.argtypes = [wt.HGLOBAL]
kernel32.GlobalUnlock.restype = wt.BOOL
kernel32.GlobalFree.argtypes = [wt.HGLOBAL]
kernel32.GlobalFree.restype = wt.HGLOBAL

user32.OpenClipboard.argtypes = [wt.HWND]
user32.OpenClipboard.restype = wt.BOOL
user32.EmptyClipboard.argtypes = []
user32.EmptyClipboard.restype = wt.BOOL
user32.GetClipboardData.argtypes = [wt.UINT]
user32.GetClipboardData.restype = wt.HANDLE
user32.SetClipboardData.argtypes = [wt.UINT, wt.HANDLE]
user32.SetClipboardData.restype = wt.HANDLE
user32.CloseClipboard.argtypes = []
user32.CloseClipboard.restype = wt.BOOL
user32.GetDC.argtypes = [wt.HWND]
user32.GetDC.restype = wt.HDC
user32.GetWindowDC.argtypes = [wt.HWND]
user32.GetWindowDC.restype = wt.HDC
user32.ReleaseDC.argtypes = [wt.HWND, wt.HDC]
user32.ReleaseDC.restype = ct.c_int
user32.IsWindow.argtypes = [wt.HWND]
user32.IsWindow.restype = wt.BOOL
user32.PrintWindow.argtypes = [wt.HWND, wt.HDC, wt.UINT]
user32.PrintWindow.restype = wt.BOOL

gdi32.CreateCompatibleDC.argtypes = [wt.HDC]
gdi32.CreateCompatibleDC.restype = wt.HDC
gdi32.CreateCompatibleBitmap.argtypes = [wt.HDC, ct.c_int, ct.c_int]
gdi32.CreateCompatibleBitmap.restype = wt.HBITMAP
gdi32.SelectObject.argtypes = [wt.HDC, wt.HGDIOBJ]
gdi32.SelectObject.restype = wt.HGDIOBJ
gdi32.BitBlt.argtypes = [wt.HDC, ct.c_int, ct.c_int, ct.c_int, ct.c_int, wt.HDC, ct.c_int, ct.c_int, wt.DWORD]
gdi32.BitBlt.restype = wt.BOOL
gdi32.GetDIBits.argtypes = [wt.HDC, wt.HBITMAP, wt.UINT, wt.UINT, ct.c_void_p, ct.POINTER(BITMAPINFO), wt.UINT]
gdi32.GetDIBits.restype = ct.c_int
gdi32.DeleteObject.argtypes = [wt.HGDIOBJ]
gdi32.DeleteObject.restype = wt.BOOL
gdi32.DeleteDC.argtypes = [wt.HDC]
gdi32.DeleteDC.restype = wt.BOOL


def is_windows() -> bool:
    return os.name == "nt"


class SendError(RuntimeError):
    pass


def ensure_dpi_aware() -> None:
    if not is_windows():
        return
    try:
        u32 = ct.WinDLL("user32", use_last_error=True)
        u32.SetProcessDPIAware()
    except Exception:
        pass


def ensure_utf8_stdio() -> None:
    if not is_windows():
        return
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass
    try:
        kernel32.SetConsoleOutputCP(65001)
        kernel32.SetConsoleCP(65001)
    except Exception:
        pass


def load_mapping(path: Path) -> Dict[str, str]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise SendError(f"Mapping file must be a JSON object: {path}")
    normalized: Dict[str, str] = {}
    for key, value in payload.items():
        if not isinstance(key, str) or not isinstance(value, str):
            raise SendError(f"Mapping file only supports string key/value pairs: {path}")
        normalized[key.strip()] = value.strip()
    return normalized


def resolve_targets(
    *,
    target_ids: Sequence[str],
    current: bool,
    mapping: Optional[Dict[str, str]],
) -> List[Dict[str, Optional[str]]]:
    cleaned = [target.strip() for target in target_ids if target and target.strip()]
    if not cleaned:
        raise SendError("At least one target id is required.")

    resolved: List[Dict[str, Optional[str]]] = []
    if current:
        for target in cleaned:
            resolved.append({"id": target, "mode": "current", "search_keyword": None})
        return resolved

    if not mapping:
        raise SendError("Mapping mode requires --mapping-file.")

    for target in cleaned:
        keyword = mapping.get(target)
        if not keyword:
            raise SendError(f"No search keyword mapping found for target id: {target}")
        resolved.append({"id": target, "mode": "search", "search_keyword": keyword})
    return resolved


def resolve_message(args: argparse.Namespace) -> str:
    if args.message_file:
        text = args.message_file.read_text(encoding="utf-8")
        return text.rstrip("\r\n")
    if args.message is not None:
        return args.message
    raise SendError("Either --message or --message-file is required.")


def get_wechat_processes(path_filter: Optional[Path]) -> List[Dict[str, object]]:
    command = [
        "powershell",
        "-NoProfile",
        "-Command",
        "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
        "Get-Process WeChat,Weixin -ErrorAction SilentlyContinue | "
        "Select-Object Id,ProcessName,Path,MainWindowTitle,MainWindowHandle | ConvertTo-Json -Compress",
    ]
    proc = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="ignore",
        check=False,
    )
    if not proc.stdout.strip():
        return []

    payload = json.loads(proc.stdout)
    items = payload if isinstance(payload, list) else [payload]
    matched: List[Dict[str, object]] = []
    for item in items:
        if not isinstance(item, dict):
            continue
        path = item.get("Path")
        if path_filter is not None and path:
            try:
                if Path(path).resolve() != path_filter.resolve():
                    continue
            except OSError:
                continue
        matched.append(
            {
                "pid": int(item["Id"]),
                "path": path,
                "title": item.get("MainWindowTitle") or "",
                "main_hwnd": int(item.get("MainWindowHandle") or 0),
            }
        )
    return matched


def get_window_text(hwnd: int) -> str:
    length = user32.GetWindowTextLengthW(hwnd)
    title_buffer = ct.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, title_buffer, length + 1)
    return title_buffer.value.strip()


def get_window_class_name(hwnd: int) -> str:
    class_buffer = ct.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, class_buffer, 256)
    return class_buffer.value.strip()


def enum_windows_for_pid(pid: int, *, require_visible: bool = True) -> List[int]:
    hwnds: List[int] = []

    enum_windows_proc = ct.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

    def callback(hwnd: int, lparam: int) -> bool:
        if require_visible and not user32.IsWindowVisible(hwnd):
            return True
        process_id = wt.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ct.byref(process_id))
        if process_id.value != pid:
            return True
        title = get_window_text(hwnd)
        if title or get_window_class_name(hwnd) == "WeChatMainWndForPC":
            hwnds.append(hwnd)
        return True

    user32.EnumWindows(enum_windows_proc(callback), 0)
    return hwnds


def find_wechat_window(path_filter: Optional[Path]) -> int:
    processes = get_wechat_processes(path_filter)
    if not processes:
        raise SendError("No running WeChat process was found.")

    fallback_main_hwnd: Optional[int] = None
    for process in processes:
        main_hwnd = int(process.get("main_hwnd") or 0)
        if main_hwnd and user32.IsWindow(main_hwnd):
            class_name = get_window_class_name(main_hwnd)
            if class_name == "WeChatMainWndForPC":
                if user32.IsWindowVisible(main_hwnd):
                    return main_hwnd
                if fallback_main_hwnd is None:
                    fallback_main_hwnd = main_hwnd

    fallback_any_hwnd: Optional[int] = None
    for process in processes:
        hwnds = enum_windows_for_pid(int(process["pid"]), require_visible=False)
        if hwnds:
            if fallback_any_hwnd is None:
                fallback_any_hwnd = hwnds[0]
            for hwnd in hwnds:
                if get_window_class_name(hwnd) == "WeChatMainWndForPC":
                    if user32.IsWindowVisible(hwnd):
                        return hwnd
                    if fallback_main_hwnd is None:
                        fallback_main_hwnd = hwnd
    if fallback_main_hwnd:
        return fallback_main_hwnd
    if fallback_any_hwnd:
        return fallback_any_hwnd
    raise SendError("No visible WeChat window was found.")


def get_window_rect(hwnd: int) -> Dict[str, int]:
    rect = wt.RECT()
    if not user32.GetWindowRect(hwnd, ct.byref(rect)):
        raise SendError("Failed to get WeChat window rect.")
    return {
        "left": rect.left,
        "top": rect.top,
        "right": rect.right,
        "bottom": rect.bottom,
        "width": rect.right - rect.left,
        "height": rect.bottom - rect.top,
    }


def bring_window_to_front(hwnd: int) -> None:
    user32.ShowWindow(hwnd, SW_RESTORE)
    user32.SetForegroundWindow(hwnd)
    time.sleep(0.3)
    
    for _ in range(10):
        rect = wt.RECT()
        if user32.GetWindowRect(hwnd, ct.byref(rect)):
            width = rect.right - rect.left
            height = rect.bottom - rect.top
            if width > 200 and height > 200 and rect.left > -1000 and rect.top > -1000:
                break
        time.sleep(0.1)
    
    if user32.GetForegroundWindow() == hwnd:
        return

    user32.keybd_event(VK_MENU, 0, 0, 0)
    user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
    time.sleep(0.1)
    user32.SetForegroundWindow(hwnd)
    time.sleep(0.2)


def set_clipboard_text(text: str) -> None:
    data = text.replace("\r\n", "\n").replace("\r", "\n")
    raw = data.encode("utf-16le") + b"\x00\x00"

    if not user32.OpenClipboard(None):
        raise SendError("Failed to open clipboard.")
    try:
        user32.EmptyClipboard()
        hglobal = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(raw))
        if not hglobal:
            raise SendError("GlobalAlloc failed while setting clipboard.")
        locked = kernel32.GlobalLock(hglobal)
        if not locked:
            kernel32.GlobalFree(hglobal)
            raise SendError("GlobalLock failed while setting clipboard.")
        try:
            ct.memmove(locked, raw, len(raw))
        finally:
            kernel32.GlobalUnlock(hglobal)
        if not user32.SetClipboardData(CF_UNICODETEXT, hglobal):
            kernel32.GlobalFree(hglobal)
            raise SendError("SetClipboardData failed.")
    finally:
        user32.CloseClipboard()


def get_clipboard_text() -> str:
    if not user32.OpenClipboard(None):
        raise SendError("Failed to open clipboard for reading.")
    try:
        handle = user32.GetClipboardData(CF_UNICODETEXT)
        if not handle:
            return ""
        locked = kernel32.GlobalLock(handle)
        if not locked:
            return ""
        try:
            return ct.wstring_at(locked)
        finally:
            kernel32.GlobalUnlock(handle)
    finally:
        user32.CloseClipboard()


def key_down(vk: int) -> None:
    user32.keybd_event(vk, 0, 0, 0)


def key_up(vk: int) -> None:
    user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)


def tap_key(vk: int, delay: float = 0.05) -> None:
    key_down(vk)
    time.sleep(delay)
    key_up(vk)


def hotkey(*keys: int, delay: float = 0.05) -> None:
    for key in keys:
        key_down(key)
        time.sleep(delay)
    for key in reversed(keys):
        key_up(key)
        time.sleep(delay)


def click_at(x: int, y: int, pause: float = 0.12) -> None:
    if not user32.SetCursorPos(x, y):
        raise SendError(f"Failed to move cursor to ({x}, {y}).")
    time.sleep(pause)
    user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.03)
    user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(pause)


def click_relative(hwnd: int, x_ratio: float, y_ratio: float) -> None:
    rect = get_window_rect(hwnd)
    x = rect["left"] + int(rect["width"] * x_ratio)
    y = rect["top"] + int(rect["height"] * y_ratio)
    click_at(x, y)


def paste_text(text: str) -> None:
    set_clipboard_text(text)
    time.sleep(0.08)
    hotkey(VK_CONTROL, VK_V)


def clear_with_ctrl_a() -> None:
    hotkey(VK_CONTROL, VK_A)
    time.sleep(0.05)
    tap_key(VK_BACK)


def capture_screen_region_bgra(*, left: int, top: int, width: int, height: int) -> bytes:
    if width <= 0 or height <= 0:
        raise SendError("Invalid capture region.")

    screen_dc = user32.GetDC(None)
    if not screen_dc:
        raise SendError("Failed to get the screen DC.")
    memory_dc = gdi32.CreateCompatibleDC(screen_dc)
    if not memory_dc:
        user32.ReleaseDC(None, screen_dc)
        raise SendError("Failed to create a compatible DC.")

    bitmap = gdi32.CreateCompatibleBitmap(screen_dc, width, height)
    if not bitmap:
        gdi32.DeleteDC(memory_dc)
        user32.ReleaseDC(None, screen_dc)
        raise SendError("Failed to create a compatible bitmap.")

    previous = gdi32.SelectObject(memory_dc, bitmap)
    try:
        if not gdi32.BitBlt(memory_dc, 0, 0, width, height, screen_dc, left, top, SRCCOPY):
            raise SendError("Failed to capture the WeChat window region.")

        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = ct.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = width
        bmi.bmiHeader.biHeight = -height
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = BI_RGB
        buffer = (ct.c_ubyte * (width * height * 4))()
        scanlines = gdi32.GetDIBits(
            memory_dc,
            bitmap,
            0,
            height,
            ct.cast(buffer, ct.c_void_p),
            ct.byref(bmi),
            DIB_RGB_COLORS,
        )
        if scanlines != height:
            raise SendError("Failed to read the captured bitmap.")
        return bytes(buffer)
    finally:
        gdi32.SelectObject(memory_dc, previous)
        gdi32.DeleteObject(bitmap)
        gdi32.DeleteDC(memory_dc)
        user32.ReleaseDC(None, screen_dc)


def capture_window_bgra(hwnd: int) -> tuple[int, int, bytes]:
    rect = get_window_rect(hwnd)
    width = rect["width"]
    height = rect["height"]
    
    if width <= 0 or height <= 0 or rect["left"] < -1000 or rect["top"] < -1000:
        user32.ShowWindow(hwnd, SW_RESTORE)
        for _ in range(10):
            time.sleep(0.1)
            rect = get_window_rect(hwnd)
            width = rect["width"]
            height = rect["height"]
            if width > 0 and height > 0 and rect["left"] > -1000 and rect["top"] > -1000:
                break
        
    if width <= 0 or height <= 0:
        raise SendError("Invalid WeChat window size for capture.")

    window_dc = user32.GetWindowDC(hwnd)
    if not window_dc:
        raise SendError("Failed to get the WeChat window DC.")
    memory_dc = gdi32.CreateCompatibleDC(window_dc)
    if not memory_dc:
        user32.ReleaseDC(hwnd, window_dc)
        raise SendError("Failed to create a compatible DC for WeChat capture.")

    bitmap = gdi32.CreateCompatibleBitmap(window_dc, width, height)
    if not bitmap:
        gdi32.DeleteDC(memory_dc)
        user32.ReleaseDC(hwnd, window_dc)
        raise SendError("Failed to create a bitmap for WeChat capture.")

    previous = gdi32.SelectObject(memory_dc, bitmap)
    try:
        if not user32.PrintWindow(hwnd, memory_dc, PW_RENDERFULLCONTENT):
            raise SendError("PrintWindow failed while capturing the WeChat window.")

        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = ct.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = width
        bmi.bmiHeader.biHeight = -height
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = BI_RGB
        buffer = (ct.c_ubyte * (width * height * 4))()
        scanlines = gdi32.GetDIBits(
            memory_dc,
            bitmap,
            0,
            height,
            ct.cast(buffer, ct.c_void_p),
            ct.byref(bmi),
            DIB_RGB_COLORS,
        )
        if scanlines != height:
            raise SendError("Failed to read the captured WeChat window bitmap.")
        raw = bytes(buffer)
        if not any(raw):
            raise SendError("Captured WeChat window bitmap was empty.")
        return width, height, raw
    finally:
        gdi32.SelectObject(memory_dc, previous)
        gdi32.DeleteObject(bitmap)
        gdi32.DeleteDC(memory_dc)
        user32.ReleaseDC(hwnd, window_dc)


def crop_bgra_region(
    raw: bytes,
    *,
    full_width: int,
    left: int,
    top: int,
    width: int,
    height: int,
) -> bytes:
    stride = full_width * 4
    rows: List[bytes] = []
    for y in range(top, top + height):
        start = y * stride + left * 4
        rows.append(raw[start : start + width * 4])
    return b"".join(rows)


def capture_window_region_bgra(
    hwnd: int,
    *,
    left_ratio: float,
    top_ratio: float,
    width_ratio: float,
    height_ratio: float,
) -> tuple[int, int, bytes]:
    rect = get_window_rect(hwnd)
    
    if rect["left"] < -1000 or rect["top"] < -1000:
        user32.ShowWindow(hwnd, SW_RESTORE)
        for _ in range(10):
            time.sleep(0.1)
            rect = get_window_rect(hwnd)
            if rect["left"] > -1000 and rect["top"] > -1000:
                break
    
    if rect["left"] < -1000 or rect["top"] < -1000:
        raise SendError("WeChat window appears to be minimized. Please restore it first.")
    
    local_left = max(0, int(rect["width"] * left_ratio))
    local_top = max(0, int(rect["height"] * top_ratio))
    region_width = max(40, int(rect["width"] * width_ratio))
    region_height = max(30, int(rect["height"] * height_ratio))
    region_width = min(region_width, rect["width"] - local_left)
    region_height = min(region_height, rect["height"] - local_top)

    try:
        full_width, _, full_raw = capture_window_bgra(hwnd)
        return (
            region_width,
            region_height,
            crop_bgra_region(
                full_raw,
                full_width=full_width,
                left=local_left,
                top=local_top,
                width=region_width,
                height=region_height,
            ),
        )
    except SendError:
        rect = get_window_rect(hwnd)
        if rect["left"] < -1000 or rect["top"] < -1000:
            raise SendError("WeChat window appears to be minimized during capture.")
        return (
            region_width,
            region_height,
            capture_screen_region_bgra(
                left=rect["left"] + local_left,
                top=rect["top"] + local_top,
                width=region_width,
                height=region_height,
            ),
        )


def install_ocr_dependency() -> None:
    proc = subprocess.run(
        [sys.executable, "-m", "pip", "install", "rapidocr-onnxruntime"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="ignore",
        check=False,
    )
    if proc.returncode != 0:
        raise SendError(
            "OCR dependency 'rapidocr-onnxruntime' is required but could not be installed automatically."
        )


def install_uia_dependency() -> None:
    proc = subprocess.run(
        [sys.executable, "-m", "pip", "install", "pywinauto"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="ignore",
        check=False,
    )
    if proc.returncode != 0:
        raise SendError(
            "UI automation dependency 'pywinauto' is required but could not be installed automatically."
        )


def get_uia_desktop():
    global UIA_DESKTOP
    if UIA_DESKTOP is not None:
        return UIA_DESKTOP

    try:
        from pywinauto import Desktop
    except ImportError:
        install_uia_dependency()
        from pywinauto import Desktop

    UIA_DESKTOP = Desktop(backend="uia")
    return UIA_DESKTOP


def get_search_edit_wrapper(hwnd: int):
    try:
        window = get_uia_desktop().window(handle=hwnd)
        search = window.child_window(title="搜索", control_type="Edit")
        if not search.exists(timeout=1):
            return None
        return search.wrapper_object()
    except Exception:
        return None


def get_ocr_engine():
    global OCR_ENGINE
    if OCR_ENGINE is not None:
        return OCR_ENGINE

    try:
        from rapidocr_onnxruntime import RapidOCR
    except ImportError:
        install_ocr_dependency()
        from rapidocr_onnxruntime import RapidOCR

    OCR_ENGINE = RapidOCR()
    return OCR_ENGINE


def normalize_match_text(text: str) -> str:
    return re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "", text).lower()


_TIME_ONLY_RE = re.compile(r"^\s*\d{1,2}:\d{2}\s*$")


def is_time_only_label(text: str) -> bool:
    return bool(_TIME_ONLY_RE.match(text.strip()))


def looks_like_message_snippet_line(text: str) -> bool:
    """Last-message preview rows; should not be used as the clickable contact title."""
    t = text.strip()
    if re.match(r"^\[\d+", t):
        return True
    if "：" in t and len(t) > 20 and any(ch.isdigit() for ch in t[:8]):
        return True
    return False


def region_box_from_ratios(
    hwnd: int,
    *,
    left_ratio: float,
    top_ratio: float,
    width_ratio: float,
    height_ratio: float,
) -> Dict[str, int]:
    rect = get_window_rect(hwnd)
    left = max(0, int(rect["width"] * left_ratio))
    top = max(0, int(rect["height"] * top_ratio))
    width = max(40, int(rect["width"] * width_ratio))
    height = max(30, int(rect["height"] * height_ratio))
    width = min(width, rect["width"] - left)
    height = min(height, rect["height"] - top)
    return {
        "left": left,
        "top": top,
        "width": width,
        "height": height,
    }


def capture_window_region_screen_bgra(
    hwnd: int,
    *,
    left_ratio: float,
    top_ratio: float,
    width_ratio: float,
    height_ratio: float,
) -> tuple[int, int, bytes]:
    """BitBlt from screen so OCR pixels match what the user sees (DWM / DPI)."""
    rect = get_window_rect(hwnd)
    w = rect["width"]
    h = rect["height"]
    local_left = max(0, int(w * left_ratio))
    local_top = max(0, int(h * top_ratio))
    region_w = max(40, int(w * width_ratio))
    region_h = max(30, int(h * height_ratio))
    region_w = min(region_w, w - local_left)
    region_h = min(region_h, h - local_top)
    abs_left = rect["left"] + local_left
    abs_top = rect["top"] + local_top
    raw = capture_screen_region_bgra(left=abs_left, top=abs_top, width=region_w, height=region_h)
    return region_w, region_h, raw


def ocr_window_region(
    hwnd: int,
    *,
    left_ratio: float,
    top_ratio: float,
    width_ratio: float,
    height_ratio: float,
) -> List[Dict[str, object]]:
    import numpy as np

    geometry = region_box_from_ratios(
        hwnd,
        left_ratio=left_ratio,
        top_ratio=top_ratio,
        width_ratio=width_ratio,
        height_ratio=height_ratio,
    )
    width, height, raw = capture_window_region_bgra(
        hwnd,
        left_ratio=left_ratio,
        top_ratio=top_ratio,
        width_ratio=width_ratio,
        height_ratio=height_ratio,
    )
    image = np.frombuffer(raw, dtype=np.uint8).reshape((height, width, 4))[:, :, :3]
    result, _ = get_ocr_engine()(image)
    entries: List[Dict[str, object]] = []
    for item in result or []:
        box, text, score = item
        text = str(text).strip()
        if not text:
            continue
        xs = [float(point[0]) for point in box]
        ys = [float(point[1]) for point in box]
        entries.append(
            {
                "text": text,
                "score": float(score),
                "left": min(xs),
                "top": min(ys),
                "right": max(xs),
                "bottom": max(ys),
                "center_x": (min(xs) + max(xs)) / 2.0 + geometry["left"],
                "center_y": (min(ys) + max(ys)) / 2.0 + geometry["top"],
            }
        )
    return entries


def dedupe_ocr_entries(entries: Sequence[Dict[str, object]]) -> List[Dict[str, object]]:
    """Drop near-duplicate boxes from overlapping crops (same text, ~same position)."""
    ordered = sorted(entries, key=lambda e: -float(e.get("score", 0.0)))
    kept: List[Dict[str, object]] = []
    for e in ordered:
        tx = normalize_match_text(str(e.get("text", "")))
        cx = float(e["center_x"])
        cy = float(e["center_y"])
        dup = False
        for k in kept:
            if normalize_match_text(str(k.get("text", ""))) != tx:
                continue
            if abs(float(k["center_x"]) - cx) < 28.0 and abs(float(k["center_y"]) - cy) < 22.0:
                dup = True
                break
        if not dup:
            kept.append(dict(e))
    return sorted(kept, key=lambda item: (float(item["top"]), float(item["left"])))


def ocr_search_results_entries(hwnd: int) -> List[Dict[str, object]]:
    merged: List[Dict[str, object]] = []
    for reg in SEARCH_RESULTS_REGIONS:
        merged.extend(
            ocr_window_region(
                hwnd,
                left_ratio=reg[0],
                top_ratio=reg[1],
                width_ratio=reg[2],
                height_ratio=reg[3],
            )
        )
    return dedupe_ocr_entries(merged)


def _save_bgra_png(path: Path, width: int, height: int, raw: bytes) -> None:
    try:
        import numpy as np
        from PIL import Image

        arr = np.frombuffer(raw, dtype=np.uint8).reshape((height, width, 4))
        rgb = arr[:, :, [2, 1, 0]]
        Image.fromarray(rgb).save(path)
    except Exception:
        pass


def dump_search_debug(
    hwnd: int,
    debug_dir: Path,
    *,
    keyword: str,
    ocr_entries: Sequence[Dict[str, object]],
    extra: Optional[Dict[str, object]] = None,
) -> None:
    debug_dir.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d-%H%M%S")
    slug = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "_", keyword).strip("_")[:40] or "kw"
    base = debug_dir / f"wechat-search-{stamp}-{slug}"
    serializable = [
        {k: v for k, v in e.items() if k in ("text", "score", "left", "top", "right", "bottom", "center_x", "center_y")}
        for e in ocr_entries
    ]
    payload = {
        "keyword": keyword,
        "regions": [list(r) for r in SEARCH_RESULTS_REGIONS],
        "entries": serializable,
        **(extra or {}),
    }
    Path(str(base) + ".ocr.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    for idx, reg in enumerate(SEARCH_RESULTS_REGIONS):
        try:
            w, h, raw = capture_window_region_screen_bgra(
                hwnd,
                left_ratio=reg[0],
                top_ratio=reg[1],
                width_ratio=reg[2],
                height_ratio=reg[3],
            )
            _save_bgra_png(Path(str(base) + f"-region{idx}.png"), w, h, raw)
        except SendError:
            pass


def text_is_context(text: str) -> bool:
    compact = normalize_match_text(text)
    if compact in OCR_STOPWORDS_NORMALIZED:
        return True
    return any(text.startswith(prefix) for prefix in OCR_CONTEXT_PREFIXES)


def has_contact_search_sections(ocr_entries: Sequence[Dict[str, object]]) -> bool:
    for entry in ocr_entries:
        compact = normalize_match_text(str(entry.get("text", "")).strip())
        if compact in CONTACT_SECTION_MARKERS_NORMALIZED:
            return True
    return False


def looks_like_network_search_results(ocr_entries: Sequence[Dict[str, object]]) -> bool:
    for entry in ocr_entries:
        text = str(entry.get("text", "")).strip()
        if any(marker in text for marker in NETWORK_RESULT_MARKERS):
            return True
    return False


def extract_search_candidates(keyword: str, ocr_entries: Sequence[Dict[str, object]]) -> List[Dict[str, object]]:
    groups: List[List[Dict[str, object]]] = []
    for raw in sorted(ocr_entries, key=lambda item: (float(item["top"]), float(item["left"]))):
        entry = dict(raw)
        text = str(entry["text"]).strip()
        if not text:
            continue
        top_v = float(entry["top"])
        if top_v < SKIP_TOP_PIXELS_IN_RESULTS:
            continue
        if "bottom" not in entry or entry["bottom"] is None:
            entry["bottom"] = top_v + 20.0
        if not groups:
            groups.append([entry])
            continue
        prev_bottom = float(groups[-1][-1]["bottom"])
        gap = top_v - prev_bottom
        if gap > RESULT_GROUP_ROW_SEP:
            groups.append([entry])
        else:
            groups[-1].append(entry)

    keyword_norm = normalize_match_text(keyword)
    candidates: List[Dict[str, object]] = []
    seen = set()
    for group in groups:
        usable = [
            item
            for item in group
            if not text_is_context(str(item["text"])) and not is_time_only_label(str(item["text"]))
        ]
        if not usable:
            continue
        combined_text = " ".join(str(item["text"]).strip() for item in group if str(item["text"]).strip())
        if keyword_norm and keyword_norm not in normalize_match_text(combined_text):
            continue
        title_pool = [item for item in usable if not looks_like_message_snippet_line(str(item["text"]))]
        if not title_pool:
            title_pool = usable
        name_entry = title_pool[0]
        if keyword_norm:
            for item in title_pool:
                if keyword_norm in normalize_match_text(str(item["text"])):
                    name_entry = item
                    break
        name = str(name_entry["text"]).strip()
        if is_time_only_label(name) or not name:
            continue
        contexts = []
        for item in group:
            tx = str(item["text"]).strip()
            if not tx:
                continue
            if (
                abs(float(item["center_x"]) - float(name_entry["center_x"])) < 2.0
                and abs(float(item["center_y"]) - float(name_entry["center_y"])) < 2.0
            ):
                continue
            contexts.append(tx)
        contexts = contexts[:3]
        combined = " ".join([name, *contexts])
        if keyword_norm and keyword_norm not in normalize_match_text(combined):
            continue
        key = normalize_match_text(name)
        if not key or key in seen:
            continue
        seen.add(key)
        candidates.append(
            {
                "name": name,
                "contexts": contexts[:3],
                "center_x": int(round(float(name_entry["center_x"]))),
                "center_y": int(round(float(name_entry["center_y"]))),
            }
        )
    return candidates


def choose_search_candidate(
    keyword: str,
    candidates: Sequence[Dict[str, object]],
    *,
    pick_index: Optional[int] = None,
) -> Dict[str, object]:
    if pick_index is not None:
        if not candidates:
            return {
                "status": "pick_out_of_range",
                "reason": f"--pick-index {pick_index} is invalid; search returned no candidates for '{keyword}'.",
                "candidates": [],
            }
        if pick_index < 1 or pick_index > len(candidates):
            return {
                "status": "pick_out_of_range",
                "reason": (
                    f"--pick-index {pick_index} is out of range; search returned {len(candidates)} candidate(s). "
                    "Re-run without --pick-index to list numbered options, or pick 1..N."
                ),
                "candidates": list(candidates)[:8],
            }
        return {"status": "resolved", "candidate": dict(candidates[pick_index - 1])}

    if not candidates:
        return {"status": "not_found", "reason": f"No search result matched '{keyword}'."}

    keyword_norm = normalize_match_text(keyword)
    exact = [item for item in candidates if normalize_match_text(str(item["name"])) == keyword_norm]
    if len(exact) == 1:
        return {"status": "resolved", "candidate": exact[0]}
    if len(candidates) == 1:
        return {"status": "resolved", "candidate": candidates[0]}
    return {
        "status": "ambiguous",
        "reason": f"Keyword '{keyword}' matched multiple conversations.",
        "candidates": list(candidates)[:8],
    }


def click_window_offset(hwnd: int, x: int, y: int) -> None:
    rect = get_window_rect(hwnd)
    click_at(rect["left"] + x, rect["top"] + y)


def locate_search_box(hwnd: int) -> Optional[Dict[str, int]]:
    entries = ocr_window_region(
        hwnd,
        left_ratio=SEARCH_BOX_REGION[0],
        top_ratio=SEARCH_BOX_REGION[1],
        width_ratio=SEARCH_BOX_REGION[2],
        height_ratio=SEARCH_BOX_REGION[3],
    )
    for entry in entries:
        text = str(entry["text"])
        if "搜索" in text or "搜素" in text or "搜" in text:
            return {
                "x": int(round(float(entry["center_x"]))),
                "y": int(round(float(entry["center_y"]))),
            }
    return None


def read_chat_title(hwnd: int) -> str:
    merged: List[Dict[str, object]] = []
    for reg in TITLE_REGIONS:
        merged.extend(
            ocr_window_region(
                hwnd,
                left_ratio=reg[0],
                top_ratio=reg[1],
                width_ratio=reg[2],
                height_ratio=reg[3],
            )
        )
    ordered = sorted(merged, key=lambda item: (float(item["top"]), float(item["left"])))
    seen = set()
    parts: List[str] = []
    for item in ordered:
        t = str(item["text"]).strip()
        if not t:
            continue
        key = normalize_match_text(t)
        if key and key not in seen:
            seen.add(key)
            parts.append(t)
    return " ".join(parts).strip()


def title_matches_expected(*, title: str, expected_name: str, keyword: str) -> bool:
    title_norm = normalize_match_text(title)
    expected_norm = normalize_match_text(expected_name)
    keyword_norm = normalize_match_text(keyword)
    if not title_norm:
        return False
    if title_norm == expected_norm or title_norm in expected_norm or expected_norm in title_norm:
        return True
    if keyword_norm and title_norm == keyword_norm and expected_norm == keyword_norm:
        return True
    return SequenceMatcher(None, title_norm, expected_norm).ratio() >= 0.72


def capture_region_signature(
    hwnd: int,
    *,
    left_ratio: float,
    top_ratio: float,
    width_ratio: float,
    height_ratio: float,
    cols: int,
    rows: int,
) -> tuple[float, ...]:
    width, height, raw = capture_window_region_screen_bgra(
        hwnd,
        left_ratio=left_ratio,
        top_ratio=top_ratio,
        width_ratio=width_ratio,
        height_ratio=height_ratio,
    )
    stride = width * 4
    signature: List[float] = []
    for row in range(rows):
        y0 = int(row * height / rows)
        y1 = max(y0 + 1, int((row + 1) * height / rows))
        sample_y_step = max(1, (y1 - y0) // 4)
        for col in range(cols):
            x0 = int(col * width / cols)
            x1 = max(x0 + 1, int((col + 1) * width / cols))
            sample_x_step = max(1, (x1 - x0) // 4)
            total = 0
            count = 0
            for y in range(y0, y1, sample_y_step):
                for x in range(x0, x1, sample_x_step):
                    index = y * stride + x * 4
                    b = raw[index]
                    g = raw[index + 1]
                    r = raw[index + 2]
                    total += (r * 299 + g * 587 + b * 114) // 1000
                    count += 1
            signature.append(total / max(1, count))
    return tuple(signature)


def capture_chat_context_signature(hwnd: int) -> Dict[str, tuple[float, ...]]:
    return {
        "header": capture_region_signature(
            hwnd,
            left_ratio=0.43,
            top_ratio=0.02,
            width_ratio=0.53,
            height_ratio=0.14,
            cols=12,
            rows=4,
        ),
        "body": capture_region_signature(
            hwnd,
            left_ratio=0.47,
            top_ratio=0.16,
            width_ratio=0.49,
            height_ratio=0.26,
            cols=12,
            rows=6,
        ),
    }


def hamming_distance(left: Sequence[int], right: Sequence[int]) -> int:
    if len(left) != len(right):
        raise SendError("Cannot compare context signatures of different sizes.")
    return sum(1 for lhs, rhs in zip(left, right) if lhs != rhs)


def mean_abs_distance(left: Sequence[float], right: Sequence[float]) -> float:
    if len(left) != len(right):
        raise SendError("Cannot compare context signatures of different sizes.")
    return sum(abs(lhs - rhs) for lhs, rhs in zip(left, right)) / len(left)


def verify_chat_switched(
    *,
    before_signature: Dict[str, Sequence[float]],
    after_signature: Dict[str, Sequence[float]],
    keyword: str,
    min_header_change: float = MIN_HEADER_CHANGE,
    min_body_change: float = MIN_BODY_CHANGE,
) -> Dict[str, float]:
    header_distance = mean_abs_distance(before_signature["header"], after_signature["header"])
    body_distance = mean_abs_distance(before_signature["body"], after_signature["body"])
    if header_distance < min_header_change or body_distance < min_body_change:
        raise SendError(
            f"Search for '{keyword}' did not appear to switch the active chat; aborted to avoid mis-send."
        )
    return {
        "header_distance": round(header_distance, 2),
        "body_distance": round(body_distance, 2),
        "context_distance": round(header_distance + body_distance, 2),
    }


def dismiss_transient_overlays() -> None:
    tap_key(VK_ESCAPE)
    time.sleep(0.08)
    tap_key(VK_ESCAPE)
    time.sleep(0.12)


def normalize_text(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").strip()


def verify_search_box_keyword(hwnd: int, keyword: str) -> None:
    wrapper = get_search_edit_wrapper(hwnd)
    if wrapper is not None:
        try:
            current_value = normalize_text(str(wrapper.iface_value.CurrentValue or ""))
        except Exception:
            current_value = ""
        if current_value:
            if normalize_match_text(current_value) != normalize_match_text(keyword):
                raise SendError(
                    f"Failed to place '{keyword}' into the WeChat search box; aborted to avoid mis-send."
                )
            return

    previous_clipboard = get_clipboard_text()
    try:
        hotkey(VK_CONTROL, VK_A)
        time.sleep(0.05)
        hotkey(VK_CONTROL, VK_C)
        time.sleep(0.08)
        captured = normalize_text(get_clipboard_text())
    finally:
        set_clipboard_text(previous_clipboard)

    if captured != normalize_text(keyword):
        raise SendError(
            f"Failed to place '{keyword}' into the WeChat search box; aborted to avoid mis-send."
        )


def focus_search_box(hwnd: int) -> None:
    # Directly click the left contact-search box. Blindly sending Escape
    # here is unsafe on personal WeChat because it can hide the main
    # window and give focus back to the host app before typing begins.
    bring_window_to_front(hwnd)
    wrapper = get_search_edit_wrapper(hwnd)
    if wrapper is not None:
        try:
            wrapper.click_input()
            time.sleep(0.2)
            return
        except Exception:
            try:
                wrapper.set_focus()
                time.sleep(0.2)
                return
            except Exception:
                pass
    box = locate_search_box(hwnd)
    if box:
        click_window_offset(hwnd, box["x"], box["y"])
    else:
        click_relative(hwnd, SEARCH_BOX_X_RATIO, SEARCH_BOX_Y_RATIO)
    time.sleep(0.12)


def focus_message_input(hwnd: int) -> None:
    # Message editor sits in the large right pane near the bottom.
    click_relative(hwnd, 0.68, 0.90)
    time.sleep(0.15)


def resolve_chat_by_keyword(
    hwnd: int,
    keyword: str,
    *,
    pick_index: Optional[int] = None,
    debug_dir: Optional[Path] = None,
) -> Dict[str, object]:
    last_title: Optional[str] = None
    last_candidate: Optional[Dict[str, object]] = None
    last_candidates: List[Dict[str, object]] = []

    for attempt_idx, dy in enumerate(TITLE_CLICK_Y_OFFSETS):
        focus_search_box(hwnd)
        clear_with_ctrl_a()
        paste_text(keyword)
        time.sleep(SEARCH_SETTLE_SEC)
        verify_search_box_keyword(hwnd, keyword)
        ocr_entries = ocr_search_results_entries(hwnd)
        candidates = extract_search_candidates(keyword, ocr_entries)
        if not has_contact_search_sections(ocr_entries) and looks_like_network_search_results(ocr_entries):
            return {
                "status": "not_found",
                "reason": f"No contact or chatroom search result matched '{keyword}'.",
                "candidate_count": 0,
                "candidates": [],
            }
        if debug_dir is not None and attempt_idx == 0:
            dump_search_debug(
                hwnd,
                debug_dir,
                keyword=keyword,
                ocr_entries=ocr_entries,
                extra={"phase": "after_search", "candidates_preview": len(candidates)},
            )
        selection = choose_search_candidate(keyword, candidates, pick_index=pick_index)
        selection["candidate_count"] = len(candidates)
        if selection["status"] != "resolved":
            if "candidates" not in selection:
                selection["candidates"] = candidates[:8]
            for idx, cand in enumerate(selection.get("candidates") or [], start=1):
                if isinstance(cand, dict):
                    cand["pick_index"] = idx
            return selection

        candidate = dict(selection["candidate"])
        last_candidates = list(candidates)
        click_window_offset(
            hwnd,
            int(candidate["center_x"]),
            int(candidate["center_y"]) + int(dy),
        )
        time.sleep(0.8)
        title = read_chat_title(hwnd)
        last_title = title
        last_candidate = candidate
        if title_matches_expected(title=title, expected_name=str(candidate["name"]), keyword=keyword):
            return {
                "status": "resolved",
                "candidate_count": len(candidates),
                "candidates": candidates[:8],
                "candidate": candidate,
                "title": title,
            }

    assert last_candidate is not None
    return {
        "status": "title_mismatch",
        "reason": (
            f"Opened chat title '{last_title or '(empty)'}' did not match '{last_candidate['name']}' "
            f"after {len(TITLE_CLICK_Y_OFFSETS)} click offset attempt(s)."
        ),
        "candidate_count": len(last_candidates),
        "candidates": last_candidates[:8],
        "candidate": last_candidate,
        "title": last_title,
    }


def send_message_to_current_chat(hwnd: int, message: str, press_enter: bool) -> None:
    focus_message_input(hwnd)
    paste_text(message)
    time.sleep(0.15)
    if press_enter:
        tap_key(VK_RETURN)


def execute_send_plan(
    *,
    hwnd: int,
    plan: Sequence[Dict[str, Optional[str]]],
    message: str,
    press_enter: bool,
    per_target_delay: float,
    dry_run: bool,
    debug_dir: Optional[Path] = None,
) -> List[Dict[str, object]]:
    results: List[Dict[str, object]] = []
    for item in plan:
        target_id = str(item["id"])
        mode = str(item["mode"])
        keyword = item.get("search_keyword")
        action: Dict[str, object] = {
            "id": target_id,
            "mode": mode,
            "search_keyword": keyword,
            "status": "planned" if dry_run else "pending",
        }

        if not dry_run:
            bring_window_to_front(hwnd)
            time.sleep(0.25)
            if mode == "search":
                if not keyword:
                    raise SendError(f"Target {target_id} is missing a search keyword.")
                raw_pick = item.get("pick_index")
                pick_index = int(raw_pick) if raw_pick is not None else None
                resolution = resolve_chat_by_keyword(
                    hwnd,
                    str(keyword),
                    pick_index=pick_index,
                    debug_dir=debug_dir,
                )
                action["status"] = str(resolution["status"])
                action["candidate_count"] = int(resolution.get("candidate_count", 0))
                if "reason" in resolution:
                    action["reason"] = str(resolution["reason"])
                if "candidates" in resolution:
                    action["candidates"] = resolution["candidates"]
                if "candidate" in resolution:
                    action["selected_candidate"] = resolution["candidate"]
                if "title" in resolution:
                    action["matched_title"] = resolution["title"]

                if resolution["status"] != "resolved":
                    results.append(action)
                    continue

                action["status"] = "sent"
            else:
                action["status"] = "sent"

            send_message_to_current_chat(hwnd, message, press_enter)
            time.sleep(per_target_delay)

        results.append(action)
    return results


def cmd_send(args: argparse.Namespace) -> int:
    ensure_dpi_aware()
    message = resolve_message(args)
    mapping = load_mapping(args.mapping_file) if args.mapping_file else None
    plan = resolve_targets(
        target_ids=args.ids,
        current=args.current,
        mapping=mapping,
    )
    if args.pick_index is not None:
        for entry in plan:
            if entry["mode"] == "search":
                entry["pick_index"] = int(args.pick_index)
    hwnd = find_wechat_window(args.wechat_path)
    bring_window_to_front(hwnd)
    time.sleep(args.focus_delay)

    results = execute_send_plan(
        hwnd=hwnd,
        plan=plan,
        message=message,
        press_enter=not args.no_enter,
        per_target_delay=args.per_target_delay,
        dry_run=args.dry_run,
        debug_dir=getattr(args, "debug_dir", None),
    )

    print(f"window_handle={hex(hwnd)}")
    print(f"message={message}")
    print(f"target_count={len(results)}")
    print(f"dry_run={args.dry_run}")
    for item in results:
        print(json.dumps(item, ensure_ascii=False))

    blocking = {"ambiguous", "not_found", "title_mismatch", "pick_out_of_range"}
    if not args.dry_run and any(str(row.get("status")) in blocking for row in results):
        print("skill_exit=needs_user_action")
        return 3
    print("skill_exit=ok")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Stable WeChat sender using UI automation")
    subparsers = parser.add_subparsers(dest="command", required=True)

    send_parser = subparsers.add_parser("send", help="Send a message via current chat or mapped searches")
    send_parser.add_argument(
        "--ids",
        nargs="+",
        required=True,
        help="One or more target ids. In --current mode these are for logging only.",
    )
    send_parser.add_argument("--message", help="Message text to send")
    send_parser.add_argument(
        "--message-file",
        type=Path,
        help="Read the message text from a UTF-8 file. Prefer this on terminals with encoding issues.",
    )
    send_parser.add_argument(
        "--current",
        action="store_true",
        help="Send to the currently opened chat instead of searching by mapping keyword.",
    )
    send_parser.add_argument(
        "--mapping-file",
        type=Path,
        help="JSON mapping file: {\"chatroom_id\": \"search keyword\"}",
    )
    send_parser.add_argument(
        "--wechat-path",
        type=Path,
        help="Optional exact WeChat.exe path filter.",
    )
    send_parser.add_argument(
        "--focus-delay",
        type=float,
        default=0.8,
        help="Delay after focusing the WeChat window. Default: 0.8 seconds",
    )
    send_parser.add_argument(
        "--per-target-delay",
        type=float,
        default=0.8,
        help="Delay between targets. Default: 0.8 seconds",
    )
    send_parser.add_argument(
        "--no-enter",
        action="store_true",
        help="Paste the message but do not press Enter.",
    )
    send_parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Resolve and print the send plan without sending anything.",
    )
    send_parser.add_argument(
        "--pick-index",
        type=int,
        default=None,
        help="1-based index into search OCR candidates when multiple rows match (after user disambiguation).",
    )
    send_parser.add_argument(
        "--debug-dir",
        type=Path,
        default=None,
        help="Save search-region screenshots and OCR JSON here when resolving by keyword.",
    )
    send_parser.set_defaults(func=cmd_send)
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    if not is_windows():
        raise SendError("This sender currently supports Windows only.")
    ensure_utf8_stdio()
    args = build_parser().parse_args(argv)
    return int(args.func(args))


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SendError as exc:
        print(f"error={exc}", file=sys.stderr)
        raise SystemExit(2)
