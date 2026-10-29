#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Global-hotkey companion for @local/dsh-formula-mathtype (the "B" mode).

The DSH plugin starts this as a small resident process. Pressing the configured
hotkey (default Ctrl+Alt+F) anywhere on the desktop:

1. reads the formula screenshot from the clipboard,
2. runs the same SimpleTex OCR + normalization + MathType OLE injection as the
   sidecar (``helper.run_request``),
3. shows a small **native Win32 toast** in the top-right corner with progress and
   the result. The toast is a layered WS_EX_NOACTIVATE window, so Word keeps the
   focus; no Tk/Tcl is involved (the local Python's Tcl package is incomplete).

The combo is captured with a low-level keyboard hook (`WH_KEYBOARD_LL`) rather
than `RegisterHotKey`, so the key never reaches the focused application. That
matters even for Ctrl+Alt+F: Word binds it to "insert footnote", so without the
hook every insertion would also add a footnote. Because this combo has no Alt
modifier, no Word keytip mode is triggered; the Esc-cleanup path only runs for
Alt/Win-based hotkeys.

Diagnostics are written to ``%LOCALAPPDATA%\\dsh-formula-mathtype\\hotkey.log``.

Standalone use (for testing)::

    python helper.py hotkey --hotkey Ctrl+Alt+F
"""

from __future__ import annotations

import ctypes
import queue
import re
import sys
import threading
import time
from ctypes import wintypes

# --- Win32 constants -------------------------------------------------------

WM_HOTKEY = 0x0312
WM_PAINT = 0x000F
WM_TIMER = 0x0113
WM_ERASEBKGND = 0x0014
WM_APP_UPDATE = 0x8001
WM_APP_EXIT = 0x8002
WM_APP_TRIGGER = 0x8003
WH_KEYBOARD_LL = 13
HC_ACTION = 0
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_SYSKEYDOWN = 0x0104
WM_SYSKEYUP = 0x0105

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
MOD_NOREPEAT = 0x4000
HOTKEY_ID = 0x5346
TIMER_ID = 0x5347
ERROR_ALREADY_EXISTS = 183

WS_POPUP = 0x80000000
WS_EX_TOPMOST = 0x00000008
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_NOACTIVATE = 0x08000000
WS_EX_LAYERED = 0x00080000
SW_HIDE = 0
SW_SHOWNOACTIVATE = 4
HWND_TOPMOST = -1
SWP_NOACTIVATE = 0x0010
SWP_SHOWWINDOW = 0x0040
LWA_ALPHA = 0x00000002
DT_LEFT = 0x0000
DT_WORDBREAK = 0x0010
DT_NOPREFIX = 0x0800
DT_CALCRECT = 0x0400
TRANSPARENT = 1
CLEARTYPE_QUALITY = 5
DEFAULT_CHARSET = 1
FW_NORMAL = 400
SM_CXSCREEN = 0

CLASS_NAME = "DshFormulaToast"
TOAST_TITLE = "dsh-formula-toast"


class WNDCLASSW(ctypes.Structure):
    _fields_ = [
        ("style", wintypes.UINT),
        ("lpfnWndProc", ctypes.c_void_p),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", wintypes.HINSTANCE),
        ("hIcon", wintypes.HICON),
        ("hCursor", wintypes.HANDLE),
        ("hbrBackground", wintypes.HBRUSH),
        ("lpszMenuName", wintypes.LPCWSTR),
        ("lpszClassName", wintypes.LPCWSTR),
    ]


class PAINTSTRUCT(ctypes.Structure):
    _fields_ = [
        ("hdc", wintypes.HDC),
        ("fErase", wintypes.BOOL),
        ("rcPaint", wintypes.RECT),
        ("fRestore", wintypes.BOOL),
        ("fIncUpdate", wintypes.BOOL),
        ("rgbReserved", ctypes.c_byte * 32),
    ]


class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("vkCode", wintypes.DWORD),
        ("scanCode", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_size_t),
    ]


WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
HOOKPROC = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)

_user32 = ctypes.WinDLL("user32", use_last_error=True)
_gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

_user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
_user32.DefWindowProcW.restype = ctypes.c_ssize_t
_user32.RegisterClassW.argtypes = [ctypes.POINTER(WNDCLASSW)]
_user32.RegisterClassW.restype = ctypes.c_ushort
_user32.CreateWindowExW.argtypes = [
    wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, wintypes.LPVOID,
]
_user32.CreateWindowExW.restype = wintypes.HWND
_user32.DestroyWindow.argtypes = [wintypes.HWND]
_user32.DestroyWindow.restype = wintypes.BOOL
_user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
_user32.ShowWindow.restype = wintypes.BOOL
_user32.UpdateWindow.argtypes = [wintypes.HWND]
_user32.UpdateWindow.restype = wintypes.BOOL
_user32.InvalidateRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT), wintypes.BOOL]
_user32.InvalidateRect.restype = wintypes.BOOL
_user32.BeginPaint.argtypes = [wintypes.HWND, ctypes.POINTER(PAINTSTRUCT)]
_user32.BeginPaint.restype = wintypes.HDC
_user32.EndPaint.argtypes = [wintypes.HWND, ctypes.POINTER(PAINTSTRUCT)]
_user32.EndPaint.restype = wintypes.BOOL
_user32.GetClientRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
_user32.GetClientRect.restype = wintypes.BOOL
_user32.GetDC.argtypes = [wintypes.HWND]
_user32.GetDC.restype = wintypes.HDC
_user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
_user32.ReleaseDC.restype = ctypes.c_int
_user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, wintypes.UINT]
_user32.SetWindowPos.restype = wintypes.BOOL
_user32.SetLayeredWindowAttributes.argtypes = [wintypes.HWND, wintypes.DWORD, ctypes.c_byte, wintypes.DWORD]
_user32.SetLayeredWindowAttributes.restype = wintypes.BOOL
_user32.SetTimer.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.UINT, wintypes.LPVOID]
_user32.SetTimer.restype = wintypes.UINT
_user32.KillTimer.argtypes = [wintypes.HWND, wintypes.UINT]
_user32.KillTimer.restype = wintypes.BOOL
_user32.SetWindowRgn.argtypes = [wintypes.HWND, wintypes.HANDLE, wintypes.BOOL]
_user32.SetWindowRgn.restype = ctypes.c_int
_user32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]
_user32.GetMessageW.restype = ctypes.c_int
_user32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
_user32.TranslateMessage.restype = wintypes.BOOL
_user32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]
_user32.DispatchMessageW.restype = ctypes.c_ssize_t
_user32.RegisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT]
_user32.RegisterHotKey.restype = wintypes.BOOL
_user32.UnregisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int]
_user32.UnregisterHotKey.restype = wintypes.BOOL
_user32.PostThreadMessageW.argtypes = [wintypes.DWORD, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
_user32.PostThreadMessageW.restype = wintypes.BOOL
_user32.SetWindowsHookExW.argtypes = [ctypes.c_int, HOOKPROC, wintypes.HINSTANCE, wintypes.DWORD]
_user32.SetWindowsHookExW.restype = wintypes.HANDLE
_user32.UnhookWindowsHookEx.argtypes = [wintypes.HANDLE]
_user32.UnhookWindowsHookEx.restype = wintypes.BOOL
_user32.CallNextHookEx.argtypes = [wintypes.HANDLE, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM]
_user32.CallNextHookEx.restype = ctypes.c_ssize_t
_user32.GetAsyncKeyState.argtypes = [ctypes.c_int]
_user32.GetAsyncKeyState.restype = ctypes.c_short
_user32.GetForegroundWindow.argtypes = []
_user32.GetForegroundWindow.restype = wintypes.HWND
_user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
_user32.GetWindowThreadProcessId.restype = wintypes.DWORD
_user32.keybd_event.argtypes = [wintypes.BYTE, wintypes.BYTE, wintypes.DWORD, ctypes.c_size_t]
_user32.keybd_event.restype = None
_user32.GetSystemMetrics.argtypes = [ctypes.c_int]
_user32.GetSystemMetrics.restype = ctypes.c_int
_user32.SetProcessDPIAware.argtypes = []
_user32.SetProcessDPIAware.restype = wintypes.BOOL
try:
    _user32.GetDpiForSystem.argtypes = []
    _user32.GetDpiForSystem.restype = wintypes.UINT
except AttributeError:
    pass

_user32.FillRect.argtypes = [wintypes.HDC, ctypes.POINTER(wintypes.RECT), wintypes.HBRUSH]
_user32.FillRect.restype = ctypes.c_int
_gdi32.CreateSolidBrush.argtypes = [wintypes.DWORD]
_gdi32.CreateSolidBrush.restype = wintypes.HBRUSH
_gdi32.DeleteObject.argtypes = [wintypes.HANDLE]
_gdi32.DeleteObject.restype = wintypes.BOOL
_gdi32.CreateFontW.argtypes = [
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, wintypes.LPCWSTR,
]
_gdi32.CreateFontW.restype = wintypes.HANDLE
_gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HANDLE]
_gdi32.SelectObject.restype = wintypes.HANDLE
_gdi32.SetTextColor.argtypes = [wintypes.HDC, wintypes.DWORD]
_gdi32.SetTextColor.restype = wintypes.DWORD
_gdi32.SetBkMode.argtypes = [wintypes.HDC, ctypes.c_int]
_gdi32.SetBkMode.restype = ctypes.c_int
_user32.DrawTextW.argtypes = [wintypes.HDC, wintypes.LPCWSTR, ctypes.c_int, ctypes.POINTER(wintypes.RECT), wintypes.UINT]
_user32.DrawTextW.restype = ctypes.c_int
_gdi32.CreateRoundRectRgn.argtypes = [ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int]
_gdi32.CreateRoundRectRgn.restype = wintypes.HANDLE

_kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
_kernel32.GetModuleHandleW.restype = wintypes.HMODULE
_kernel32.GetCurrentThreadId.argtypes = []
_kernel32.GetCurrentThreadId.restype = wintypes.DWORD
_kernel32.CreateMutexW.argtypes = [ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR]
_kernel32.CreateMutexW.restype = wintypes.HANDLE
_kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
_kernel32.OpenProcess.restype = wintypes.HANDLE
_kernel32.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
_kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
_kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
_kernel32.CloseHandle.restype = wintypes.BOOL


def _rgb(red: int, green: int, blue: int) -> int:
    return red | (green << 8) | (blue << 16)


def _append_log(line: str) -> None:
    """Append one timestamped line to the hotkey log; never raises."""
    try:
        import os
        from pathlib import Path

        base = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "dsh-formula-mathtype"
        base.mkdir(parents=True, exist_ok=True)
        log_path = base / "hotkey.log"
        if log_path.exists() and log_path.stat().st_size > 1_000_000:
            log_path.write_text("", encoding="utf-8")
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(time.strftime("[%Y-%m-%d %H:%M:%S] ") + line.replace("\n", " | ") + "\n")
    except Exception:
        pass


def _foreground_process_name() -> str:
    """Basename (lowercase) of the process owning the foreground window."""
    try:
        hwnd = _user32.GetForegroundWindow()
        if not hwnd:
            return ""
        pid = wintypes.DWORD()
        _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        handle = _kernel32.OpenProcess(0x1000, False, int(pid.value))
        if not handle:
            return ""
        try:
            buffer = ctypes.create_unicode_buffer(1024)
            size = wintypes.DWORD(1024)
            if not _kernel32.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
                return ""
            return buffer.value.rsplit("\\", 1)[-1].lower()
        finally:
            _kernel32.CloseHandle(handle)
    except Exception:
        return ""


def parse_hotkey(spec: str) -> tuple[int, int]:
    """Turn "Ctrl+Alt+F" into (MOD_*, virtual-key code)."""
    text = (spec or "").strip()
    if not text:
        raise ValueError("热键为空")
    modifiers = 0
    key: int | None = None
    for part in [piece.strip() for piece in text.split("+") if piece.strip()]:
        lowered = part.lower()
        if lowered in ("ctrl", "control"):
            modifiers |= MOD_CONTROL
        elif lowered == "alt":
            modifiers |= MOD_ALT
        elif lowered == "shift":
            modifiers |= MOD_SHIFT
        elif lowered in ("win", "windows", "meta"):
            modifiers |= MOD_WIN
        elif re.fullmatch(r"[fF]([1-9]|1[0-9]|2[0-4])", part):
            key = 0x70 + int(part[1:]) - 1
        elif len(part) == 1 and part.isascii() and (part.isalpha() or part.isdigit()):
            key = ord(part.upper())
        else:
            raise ValueError("无法解析热键片段：%s" % part)
    if key is None:
        raise ValueError("热键需要包含一个主键，例如 Ctrl+Alt+F")
    if modifiers == 0:
        raise ValueError("热键至少需要一个修饰键（Ctrl/Alt/Shift/Win）")
    return modifiers | MOD_NOREPEAT, key


def _acquire_single_instance() -> bool:
    """Named mutex so a stale helper cannot double-register the hotkey."""
    try:
        handle = _kernel32.CreateMutexW(None, False, "dsh-formula-mathtype-hotkey")
        if not handle:
            return False
        if ctypes.get_last_error() == ERROR_ALREADY_EXISTS:
            return False
        globals()["_MUTEX_HANDLE"] = handle
        return True
    except Exception:
        return True


class ToastWindow:
    """Borderless, click-through, no-activate toast drawn with GDI."""

    def __init__(self, scale: float) -> None:
        self.scale = scale
        self.hwnd: int | None = None
        self.text = ""
        self.error = False
        self.font = _gdi32.CreateFontW(
            -int(15 * scale), 0, 0, 0, FW_NORMAL, 0, 0, 0,
            DEFAULT_CHARSET, 0, 0, CLEARTYPE_QUALITY, 0, "Microsoft YaHei UI",
        )
        self._proc = WNDPROC(self._wndproc)
        class_struct = WNDCLASSW()
        class_struct.style = 0
        class_struct.lpfnWndProc = ctypes.cast(self._proc, ctypes.c_void_p)
        class_struct.hInstance = _kernel32.GetModuleHandleW(None)
        class_struct.lpszClassName = CLASS_NAME
        _user32.RegisterClassW(ctypes.byref(class_struct))
        self.hwnd = _user32.CreateWindowExW(
            WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_LAYERED,
            CLASS_NAME, TOAST_TITLE, WS_POPUP,
            0, 0, 10, 10, None, None, class_struct.hInstance, None,
        )
        if not self.hwnd:
            raise OSError("CreateWindowExW failed")
        _user32.SetLayeredWindowAttributes(self.hwnd, 0, 255, LWA_ALPHA)

    # -- sizing / painting -------------------------------------------------

    def _measure(self, text: str) -> tuple[int, int]:
        hdc = _user32.GetDC(self.hwnd)
        try:
            old = _gdi32.SelectObject(hdc, self.font)
            try:
                rect = wintypes.RECT(0, 0, int(430 * self.scale), 0)
                _user32.DrawTextW(
                    hdc, text, -1, ctypes.byref(rect), DT_LEFT | DT_WORDBREAK | DT_NOPREFIX | DT_CALCRECT
                )
                return max(rect.right - rect.left, 60), max(rect.bottom - rect.top, 18)
            finally:
                _gdi32.SelectObject(hdc, old)
        finally:
            _user32.ReleaseDC(self.hwnd, hdc)

    def _paint(self, hwnd: int) -> None:
        paint = PAINTSTRUCT()
        hdc = _user32.BeginPaint(hwnd, ctypes.byref(paint))
        try:
            rect = wintypes.RECT()
            _user32.GetClientRect(hwnd, ctypes.byref(rect))
            border = _rgb(227, 176, 176) if self.error else _rgb(208, 215, 222)
            background = _rgb(253, 236, 236) if self.error else _rgb(255, 255, 255)
            foreground = _rgb(164, 0, 0) if self.error else _rgb(31, 35, 40)

            brush = _gdi32.CreateSolidBrush(border)
            _user32.FillRect(hdc, ctypes.byref(rect), brush)
            _gdi32.DeleteObject(brush)

            inner = wintypes.RECT(1, 1, rect.right - 1, rect.bottom - 1)
            brush = _gdi32.CreateSolidBrush(background)
            _user32.FillRect(hdc, ctypes.byref(inner), brush)
            _gdi32.DeleteObject(brush)

            pad = int(14 * self.scale)
            text_rect = wintypes.RECT(pad, pad, rect.right - pad, rect.bottom - pad)
            old_font = _gdi32.SelectObject(hdc, self.font)
            _gdi32.SetBkMode(hdc, TRANSPARENT)
            _gdi32.SetTextColor(hdc, foreground)
            _user32.DrawTextW(hdc, self.text, -1, ctypes.byref(text_rect), DT_LEFT | DT_WORDBREAK | DT_NOPREFIX)
            _gdi32.SelectObject(hdc, old_font)
        finally:
            _user32.EndPaint(hwnd, ctypes.byref(paint))

    def _wndproc(self, hwnd, message, wparam, lparam):
        if message == WM_PAINT:
            self._paint(hwnd)
            return 0
        if message == WM_TIMER:
            self.hide()
            return 0
        if message == WM_ERASEBKGND:
            return 1
        return _user32.DefWindowProcW(hwnd, message, wparam, lparam)

    # -- show / hide -------------------------------------------------------

    def show(self, text: str, *, error: bool = False, ttl_ms: int = 0) -> None:
        if not self.hwnd:
            return
        self.text = text
        self.error = error
        pad = int(14 * self.scale)
        text_width, text_height = self._measure(text)
        width = text_width + 2 * pad
        height = text_height + 2 * pad
        screen_width = _user32.GetSystemMetrics(SM_CXSCREEN)
        x = max(screen_width - width - int(24 * self.scale), 0)
        y = int(56 * self.scale)
        _user32.SetWindowPos(self.hwnd, HWND_TOPMOST, x, y, width, height, SWP_NOACTIVATE)
        region = _gdi32.CreateRoundRectRgn(0, 0, width + 1, height + 1, int(10 * self.scale), int(10 * self.scale))
        if region:
            _user32.SetWindowRgn(self.hwnd, region, True)
        _user32.InvalidateRect(self.hwnd, None, True)
        _user32.ShowWindow(self.hwnd, SW_SHOWNOACTIVATE)
        _user32.UpdateWindow(self.hwnd)
        _user32.KillTimer(self.hwnd, TIMER_ID)
        if ttl_ms > 0:
            _user32.SetTimer(self.hwnd, TIMER_ID, ttl_ms, None)

    def hide(self) -> None:
        if not self.hwnd:
            return
        _user32.KillTimer(self.hwnd, TIMER_ID)
        _user32.ShowWindow(self.hwnd, SW_HIDE)

    def destroy(self) -> None:
        if self.hwnd:
            _user32.DestroyWindow(self.hwnd)
            self.hwnd = None
        if self.font:
            _gdi32.DeleteObject(self.font)
            self.font = None


class HotkeyApp:
    """Owns the hotkey message loop and the toast; OCR work runs on a worker."""

    def __init__(self, *, hotkey: str, token: str, token_file: str, model: str) -> None:
        self.hotkey = hotkey
        self.token = token
        self.token_file = token_file
        self.model = model
        self.events: queue.Queue = queue.Queue()
        self.busy = False
        self.thread_id = 0
        self.toast: ToastWindow | None = None
        self.exit_code = 0
        self.modifiers = 0
        self.vk = 0
        self.hook = None
        self._hook_proc_ref = None
        self._latched = False
        self._needs_esc = False

    def _post_update(self) -> None:
        try:
            _user32.PostThreadMessageW(self.thread_id, WM_APP_UPDATE, 0, 0)
        except Exception:
            pass

    def _schedule_exit(self, delay: float) -> None:
        def _fire() -> None:
            time.sleep(delay)
            try:
                _user32.PostThreadMessageW(self.thread_id, WM_APP_EXIT, 0, 0)
            except Exception:
                pass

        threading.Thread(target=_fire, daemon=True).start()

    def run(self) -> int:
        try:
            _user32.SetProcessDPIAware()
        except Exception:
            pass
        self.thread_id = int(_kernel32.GetCurrentThreadId())
        scale = 1.0
        try:
            scale = max(1.0, float(_user32.GetDpiForSystem()) / 96.0)
        except Exception:
            scale = 1.0
        try:
            self.toast = ToastWindow(scale)
        except Exception as exc:
            _append_log("toast unavailable: %r" % (exc,))
            self.toast = None

        try:
            modifiers, vk = parse_hotkey(self.hotkey)
        except ValueError as exc:
            return self._fatal(str(exc))
        self.modifiers = modifiers & (MOD_ALT | MOD_CONTROL | MOD_SHIFT | MOD_WIN)
        self.vk = vk
        has_alt = bool(self.modifiers & MOD_ALT)
        has_ctrl = bool(self.modifiers & MOD_CONTROL)
        has_win = bool(self.modifiers & MOD_WIN)
        # Alt/Win alone or with Shift can put Word into keytip/Start mode; Ctrl+Alt
        # (the default Ctrl+Alt+F) does not, so no Esc is sent for it.
        self._needs_esc = (has_alt or has_win) and not (has_alt and has_ctrl)
        self._hook_proc_ref = HOOKPROC(self._hook_proc)
        self.hook = _user32.SetWindowsHookExW(
            WH_KEYBOARD_LL, self._hook_proc_ref, _kernel32.GetModuleHandleW(None), 0
        )
        if not self.hook:
            return self._fatal(
                "安装全局键盘钩子失败（热键 %s 无法注册）。"
                "可在 cordis.patch.yml 的插件配置里换一个 hotkey，例如 Ctrl+Alt+M。" % self.hotkey
            )
        _append_log("hotkey registered: %s (low-level hook)" % self.hotkey)
        if self.toast is not None:
            self.toast.show("公式热键已就绪：%s" % self.hotkey, ttl_ms=1800)
        try:
            return self._message_loop()
        finally:
            if self.hook:
                try:
                    _user32.UnhookWindowsHookEx(self.hook)
                except Exception:
                    pass
                self.hook = None
            if self.toast is not None:
                self.toast.destroy()

    def _fatal(self, message: str) -> int:
        _append_log("fatal: " + message)
        if self.toast is not None:
            self.toast.show(message, error=True, ttl_ms=8000)
        self._schedule_exit(9.0)
        self.exit_code = 3
        try:
            return self._message_loop()
        finally:
            if self.toast is not None:
                self.toast.destroy()

    def _message_loop(self) -> int:
        message = wintypes.MSG()
        while True:
            result = int(_user32.GetMessageW(ctypes.byref(message), None, 0, 0))
            if result <= 0:
                break
            if int(message.hWnd or 0) == 0:
                if message.message == WM_APP_UPDATE:
                    self._drain_events()
                elif message.message == WM_APP_EXIT:
                    break
                elif message.message == WM_APP_TRIGGER:
                    self._on_press()
                continue
            _user32.TranslateMessage(ctypes.byref(message))
            _user32.DispatchMessageW(ctypes.byref(message))
        return self.exit_code

    # -- low-level keyboard hook -------------------------------------------

    def _post_trigger(self) -> None:
        try:
            _user32.PostThreadMessageW(self.thread_id, WM_APP_TRIGGER, 0, 0)
        except Exception:
            pass

    def _modifiers_match(self) -> bool:
        def down(vk: int) -> bool:
            return bool(_user32.GetAsyncKeyState(vk) & 0x8000)

        ctrl = down(0x11)  # VK_CONTROL
        alt = down(0x12)  # VK_MENU
        shift = down(0x10)  # VK_SHIFT
        win = down(0x5B) or down(0x5C)  # VK_LWIN / VK_RWIN
        return (
            ctrl == bool(self.modifiers & MOD_CONTROL)
            and alt == bool(self.modifiers & MOD_ALT)
            and shift == bool(self.modifiers & MOD_SHIFT)
            and win == bool(self.modifiers & MOD_WIN)
        )

    def _hook_proc(self, n_code, w_param, l_param):
        """Swallow the configured combo so Word/other apps never see it."""
        try:
            if n_code == HC_ACTION:
                keyboard = ctypes.cast(l_param, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
                vk = int(keyboard.vkCode)
                message = int(w_param)
                if message in (WM_KEYDOWN, WM_SYSKEYDOWN):
                    if vk == self.vk and self._modifiers_match():
                        if not self._latched:
                            self._latched = True
                            self._post_trigger()
                        return 1  # swallow the key-down (and auto-repeats)
                elif message in (WM_KEYUP, WM_SYSKEYUP):
                    if vk == self.vk and self._latched:
                        self._latched = False
                        return 1  # swallow the matching key-up
        except Exception:
            pass
        return _user32.CallNextHookEx(self.hook or 0, n_code, w_param, l_param)

    def _dismiss_word_keytips(self) -> None:
        """Alt-based combos put Word into keytip mode; Esc clears it."""
        if not self._needs_esc:
            return
        if _foreground_process_name() != "winword.exe":
            return
        try:
            _user32.keybd_event(0x1B, 0, 0, 0)  # VK_ESCAPE down
            _user32.keybd_event(0x1B, 0, 2, 0)  # up
            _append_log("cleared Word keytip mode")
        except Exception:
            pass

    def _schedule_menu_dismiss(self) -> None:
        if not self._needs_esc:
            return

        def _fire() -> None:
            time.sleep(0.3)
            self._dismiss_word_keytips()

        threading.Thread(target=_fire, daemon=True, name="hotkey-keytips").start()

    # -- events ------------------------------------------------------------

    def _on_press(self) -> None:
        if self.busy:
            _append_log("hotkey pressed while busy; ignored")
            if self.toast is not None:
                self.toast.show("正在处理上一条公式，请稍候…", ttl_ms=1400)
            return
        self.busy = True
        _append_log("hotkey pressed")
        self._schedule_menu_dismiss()
        if self.toast is not None:
            self.toast.show("正在识别公式…", ttl_ms=0)
        threading.Thread(target=self._work, daemon=True, name="hotkey-worker").start()

    def _work(self) -> None:
        import helper

        try:
            if not self.token:
                self.token = helper.read_token(self.token_file)
            result = helper.run_request({"action": "run", "mode": "clipboard"}, token=self.token, model=self.model)
            self.events.put(("result", result))
        except Exception as exc:
            self.events.put(("error", str(exc)))
        finally:
            self._post_update()

    def _drain_events(self) -> None:
        while True:
            try:
                kind, payload = self.events.get_nowait()
            except queue.Empty:
                return
            self.busy = False
            if kind == "result":
                latex = str(payload.get("latex") or "")
                conf = payload.get("conf")
                title = "✓ 已插入 MathType 公式"
                if isinstance(conf, (int, float)):
                    title += "（conf %.2f）" % float(conf)
                preview = latex if len(latex) <= 150 else latex[:147] + "…"
                _append_log("inserted: " + latex[:300])
                if self.toast is not None:
                    self.toast.show(title + ("\n" + preview if preview else ""), ttl_ms=2600)
            else:
                _append_log("insert failed: " + str(payload))
                if self.toast is not None:
                    self.toast.show("✗ 插入失败\n" + str(payload), error=True, ttl_ms=8000)


def run_hotkey_app(*, hotkey: str, token: str, token_file: str, model: str) -> int:
    """Entry point used by helper.py's ``hotkey`` subcommand."""
    if not _acquire_single_instance():
        _append_log("another hotkey helper is already running; exiting")
        print("另一个热键助手已在运行。", file=sys.stderr)
        return 2
    _append_log("starting hotkey helper (hotkey=%s, model=%s)" % (hotkey, model))
    app = HotkeyApp(hotkey=hotkey, token=token, token_file=token_file, model=model)
    try:
        return app.run()
    except Exception as exc:
        _append_log("fatal: %r" % (exc,))
        return 1