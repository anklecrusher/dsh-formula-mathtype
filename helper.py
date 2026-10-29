#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Windows sidecar for @local/dsh-formula-mathtype.

What it does
------------
One request inserts one formula into the *active* Word document:

1. read the formula image from the Windows clipboard (or a file path),
2. POST it to the SimpleTex formula-recognition open API and take ``res.latex``,
3. hand that LaTeX to MathType's own SDK interface (``MT6.dll``:
   ``MTAPIConnect`` / ``MTXFormSetTranslator`` / ``MTXFormEqn``), which puts a
   real MathType OLE object (``Equation.DSMT4``) on the clipboard,
4. paste it into Word through Word COM (``GetActiveObject`` + ``Selection.Paste``).

No MathType window is opened, no Word macro runs, and no image or OMML is
involved: the result is an editable MathType equation object.

Protocol (``serve`` mode, one JSON object per line on stdin/stdout)::

    -> {"id": 1, "action": "run", "mode": "clipboard"}
    -> {"id": 2, "action": "run", "mode": "latex", "latex": "\\frac{a}{b}"}
    -> {"id": 3, "action": "run", "mode": "image", "imagePath": "C:/x.png"}
    <- {"id": 1, "ok": true, "latex": "...", "conf": 0.95, "inserted": 1, ...}
    <- {"id": 2, "ok": false, "error": "..."}

The token is read from the file configured with ``--token-file`` (default:
``%LOCALAPPDATA%\\dsh-formula-mathtype\\simpletex.token``). The token is never
printed or written anywhere else.
"""

from __future__ import annotations

import argparse
import ctypes
import json
import os
import re
import sys
import tempfile
import time
from pathlib import Path

PROG = "formula-mathtype"

# MathType MT6.dll SDK constants (from the MathType SDK / mtsdk.h).
TRANSLATOR_SLOT = 4
LOCAL_BUFFER_SOURCE = -3
CLIPBOARD_SOURCE = -2
FORMAT_TEXT = 7
FORMAT_MATHTYPE_OLE = 6
LATEX_TRANSLATOR = b"LaTeX.tdl"

SIMPLETEX_TURBO_URL = "https://server.simpletex.cn/api/latex_ocr_turbo"
SIMPLETEX_STANDARD_URL = "https://server.simpletex.cn/api/latex_ocr"

DEFAULT_TOKEN_FILE = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "dsh-formula-mathtype" / "simpletex.token"

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff"}
MIME_BY_SUFFIX = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".bmp": "image/bmp",
    ".gif": "image/gif",
    ".webp": "image/webp",
    ".tif": "image/tiff",
    ".tiff": "image/tiff",
}


class RunError(RuntimeError):
    """A request failure that still carries the recognized LaTeX, if any."""

    def __init__(self, message: str, *, latex: str | None = None, conf: float | None = None) -> None:
        super().__init__(message)
        self.latex = latex
        self.conf = conf


def log(*parts) -> None:
    """Diagnostics go to stderr; stdout is reserved for the JSON protocol."""
    print(PROG + ":", *parts, file=sys.stderr, flush=True)


def respond(**payload) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def read_token(path: str) -> str:
    token_file = Path(path) if path else DEFAULT_TOKEN_FILE
    try:
        raw = token_file.read_text(encoding="utf-8-sig", errors="replace")
    except FileNotFoundError:
        raise RunError(
            "找不到 SimpleTex 令牌文件：%s\n"
            "请把用户授权令牌（UAT）写进该文件（一行，前后不要有空格）后重试。" % token_file
        ) from None
    for line in reversed(raw.splitlines()):
        line = line.strip()
        if line and not line.startswith("#"):
            return line
    raise RunError("令牌文件里没有有效的令牌行：%s" % token_file)


# --------------------------------------------------------------------------
# MathType: LaTeX -> clipboard MathType OLE object, via the MT6.dll SDK API
# --------------------------------------------------------------------------

def _strip_exe_path(command: str) -> str:
    command = (command or "").strip()
    if not command:
        return ""
    if command.startswith('"'):
        end = command.find('"', 1)
        return command[1:end] if end > 1 else command.strip('"')
    match = re.search(r"\.exe", command, flags=re.IGNORECASE)
    return command[: match.end()] if match else command.split(" ")[0]


def find_mathtype_dll() -> Path:
    """Locate the MT6.dll that matches this Python's bitness."""
    candidates: list[Path] = []
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CLASSES_ROOT,
            r"CLSID\{0002CE03-0000-0000-C000-000000000046}\LocalServer32",
        ) as key:
            raw, _ = winreg.QueryValueEx(key, "")
        exe = _strip_exe_path(str(raw))
        if exe:
            root = Path(exe).resolve().parent
            candidates.extend(
                [
                    root / "System" / "64" / "MT6.dll",
                    root / "System" / "32" / "MT6.dll",
                    root / "System" / "MT6.dll",
                ]
            )
    except OSError:
        pass
    candidates.extend(
        [
            Path(r"C:\Program Files\MathType\System\64\MT6.dll"),
            Path(r"C:\Program Files\MathType\System\32\MT6.dll"),
            Path(r"C:\Program Files (x86)\MathType\System\64\MT6.dll"),
            Path(r"C:\Program Files (x86)\MathType\System\32\MT6.dll"),
        ]
    )
    is_64bit = ctypes.sizeof(ctypes.c_void_p) == 8
    preferred = [c for c in candidates if ("System\\64" in str(c) or "System/64" in str(c)) == is_64bit]
    ordered = preferred + [c for c in candidates if c not in preferred]
    for candidate in ordered:
        if candidate.exists():
            return candidate
    raise RunError("找不到 MathType 的 MT6.dll：请确认已安装 MathType 7。")


class MathTypeEngine:
    """Keeps the MathType API connection warm for the lifetime of the sidecar."""

    def __init__(self) -> None:
        self._dll = None
        self._connected = False

    def _load(self):
        if self._dll is not None:
            return self._dll
        dll_path = find_mathtype_dll()
        dll = ctypes.WinDLL(str(dll_path))
        dll.MTAPIConnect.argtypes = [ctypes.c_int, ctypes.c_int]
        dll.MTAPIConnect.restype = ctypes.c_int
        dll.MTAPIDisconnect.argtypes = []
        dll.MTAPIDisconnect.restype = ctypes.c_int
        dll.MTXFormSetTranslator.argtypes = [ctypes.c_int, ctypes.c_char_p]
        dll.MTXFormSetTranslator.restype = ctypes.c_int
        dll.MTXFormEqn.argtypes = [
            ctypes.c_int,  # source type
            ctypes.c_int,  # source format
            ctypes.c_void_p,  # source data
            ctypes.c_int,  # source length
            ctypes.c_int,  # destination type
            ctypes.c_int,  # destination format
            ctypes.c_void_p,  # destination data
            ctypes.c_int,  # destination length
            ctypes.c_char_p,  # translator name override
            ctypes.c_void_p,  # reserved
        ]
        dll.MTXFormEqn.restype = ctypes.c_int
        self._dll = dll
        log("loaded MathType dll:", dll_path)
        return dll

    def connect(self):
        dll = self._load()
        if self._connected:
            return dll
        rc = int(dll.MTAPIConnect(0, 30))
        if rc != 0:
            raise RunError("连接 MathType API 失败 (MTAPIConnect rc=%d)" % rc)
        rc = int(dll.MTXFormSetTranslator(TRANSLATOR_SLOT, LATEX_TRANSLATOR))
        if rc != 0:
            raise RunError("设置 MathType LaTeX 转换器失败 (rc=%d)" % rc)
        self._connected = True
        return dll

    def latex_to_clipboard_object(self, latex: str) -> None:
        """Convert LaTeX text into a MathType OLE object on the clipboard."""
        payload = latex.encode("utf-8") + b"\x00"
        for attempt in range(2):
            dll = self.connect()
            rc = int(
                dll.MTXFormEqn(
                    LOCAL_BUFFER_SOURCE,
                    FORMAT_TEXT,
                    ctypes.c_char_p(payload),
                    len(payload),
                    CLIPBOARD_SOURCE,
                    FORMAT_MATHTYPE_OLE,
                    None,
                    0,
                    b"",
                    None,
                )
            )
            if rc == 0:
                return
            # A stale connection (MathType restarted) is the usual cause; retry once.
            self.close()
            if attempt == 1:
                raise RunError("MathType 转换 LaTeX 失败 (MTXFormEqn rc=%d)" % rc)

    def close(self) -> None:
        if self._dll is not None and self._connected:
            try:
                self._dll.MTAPIDisconnect()
            except Exception:
                pass
        self._connected = False


# --------------------------------------------------------------------------
# Word: paste the clipboard object at the caret and verify it landed
# --------------------------------------------------------------------------

def insert_into_active_word() -> int:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            raise RunError(
                "没有找到正在运行的 Word。请先打开要写入的报告文档，"
                "把光标放到要插入公式的位置，再重新触发。"
            ) from None
        if int(word.Documents.Count) < 1:
            raise RunError("Word 里没有打开的文档，请先打开报告文档并定位光标。")
        doc = word.ActiveDocument
        before = int(doc.InlineShapes.Count)
        try:
            word.Selection.Paste()
        except Exception as exc:
            raise RunError("粘贴到 Word 失败：%s" % exc) from exc
        time.sleep(0.25)
        after = int(doc.InlineShapes.Count)
        if after <= before:
            raise RunError(
                "已在 Word 中执行粘贴，但文档里没有新增 MathType 对象；"
                "请确认光标在正文中且 MathType 已安装。"
            )
        return after - before
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _set_clipboard_text(text: str) -> None:
    try:
        import win32clipboard
        import win32con

        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        pass


# --------------------------------------------------------------------------
# Clipboard image -> SimpleTex OCR -> LaTeX
# --------------------------------------------------------------------------

def grab_clipboard_image() -> tuple[str, bool]:
    """Return (path, remove_after_use)."""
    try:
        from PIL import ImageGrab
    except ImportError as exc:
        raise RunError("缺少 Pillow：请先安装 Pillow（pip install pillow）。") from exc
    try:
        data = ImageGrab.grabclipboard()
    except Exception as exc:
        raise RunError("读取剪贴板失败：%s" % exc) from exc
    if data is None:
        raise RunError("剪贴板里没有图片。请先用 Win+Shift+S（或截图工具）截取公式，再触发插入。")
    if isinstance(data, list):
        for item in data:
            path = Path(str(item))
            if path.is_file() and path.suffix.lower() in IMAGE_SUFFIXES:
                return str(path), False
        raise RunError("剪贴板里是文件列表，但没有可用的图片文件。")
    fd, tmp = tempfile.mkstemp(prefix="dsh-formula-", suffix=".png")
    os.close(fd)
    try:
        data.convert("RGB").save(tmp, "PNG")
    except Exception:
        Path(tmp).unlink(missing_ok=True)
        raise RunError("保存剪贴板图片失败。") from None
    return tmp, True


def _strip_latex_wrappers(latex: str) -> str:
    text = latex.strip()
    if len(text) >= 4 and (
        (text.startswith("$$") and text.endswith("$$"))
        or (text.startswith("\\(") and text.endswith("\\)"))
        or (text.startswith("\\[") and text.endswith("\\]"))
    ):
        return text[2:-2].strip()
    if len(text) >= 2 and text.startswith("$") and text.endswith("$"):
        return text[1:-1].strip()
    return text


def _mime_for(path: str) -> str:
    return MIME_BY_SUFFIX.get(Path(path).suffix.lower(), "image/png")


def _extract_latex(data: dict) -> tuple[str, float | None]:
    if not data.get("status"):
        err = data.get("err_info") or {}
        detail = " ".join(str(err.get(k) or "") for k in ("err_type", "err_msg")).strip()
        raise RunError("SimpleTex 识别失败：%s" % (detail or "未知错误"))
    res = data.get("res") or {}
    latex = str(res.get("latex") or "").strip()
    if not latex:
        raise RunError("SimpleTex 没有返回 latex 字段。")
    if latex == "[DOCIMG]":
        raise RunError("这张图看起来是整页文档而不是单条公式；请只截取公式区域。")
    if latex == "[EMPTY]":
        raise RunError("图片里没有识别到公式（识别结果为空）。")
    conf = res.get("conf")
    try:
        conf = float(conf) if conf is not None else None
    except (TypeError, ValueError):
        conf = None
    return _strip_latex_wrappers(latex), conf


def recognize(image_path: str, token: str, model: str) -> tuple[str, float | None]:
    import requests

    url = SIMPLETEX_STANDARD_URL if model == "standard" else SIMPLETEX_TURBO_URL
    path = Path(image_path)
    try:
        with path.open("rb") as handle:
            files = {"file": (path.name or "formula.png", handle, _mime_for(image_path))}
            response = requests.post(url, headers={"token": token}, files=files, timeout=90)
    except requests.RequestException as exc:
        raise RunError("请求 SimpleTex 失败（网络问题或服务不可用）：%s" % exc) from exc
    try:
        data = response.json()
    except ValueError:
        raise RunError(
            "SimpleTex 返回的不是 JSON（HTTP %s）：%s" % (response.status_code, response.text[:200])
        ) from None
    return _extract_latex(data)


# --------------------------------------------------------------------------
# Request handling
# --------------------------------------------------------------------------

def run_request(request: dict, engine: MathTypeEngine, *, token: str, model: str) -> dict:
    started = time.time()
    mode = str(request.get("mode") or "").strip().lower()
    if not mode:
        mode = "latex" if request.get("latex") else "clipboard"

    cleanup_path: str | None = None
    try:
        if mode == "latex":
            latex = _strip_latex_wrappers(str(request.get("latex") or ""))
            if not latex:
                raise RunError("mode=latex 时必须提供 latex 文本。")
            conf: float | None = None
        elif mode == "image":
            image_path = str(request.get("imagePath") or "").strip()
            if not image_path:
                raise RunError("mode=image 时必须提供 imagePath。")
            if not Path(image_path).is_file():
                raise RunError("找不到图片文件：%s" % image_path)
            latex, conf = recognize(image_path, token, model)
        elif mode == "clipboard":
            image_path, should_remove = grab_clipboard_image()
            cleanup_path = image_path if should_remove else None
            latex, conf = recognize(image_path, token, model)
        else:
            raise RunError("未知的 mode：%s" % mode)

        try:
            engine.latex_to_clipboard_object(latex)
        except Exception as exc:
            raise RunError(
                "%s\n(公式 LaTeX 已复制到系统剪贴板，可手动粘贴到 Word 后用 MathType 转换。)" % exc,
                latex=latex,
                conf=conf,
            ) from exc

        try:
            inserted = insert_into_active_word()
        except Exception as exc:
            _set_clipboard_text(latex)
            raise RunError(
                "%s\n(公式 LaTeX 已复制到系统剪贴板，可手动粘贴到 Word 后用 MathType 转换。)" % exc,
                latex=latex,
                conf=conf,
            ) from exc

        return {
            "ok": True,
            "mode": mode,
            "latex": latex,
            "conf": conf,
            "inserted": inserted,
            "elapsedMs": int((time.time() - started) * 1000),
            "message": "已插入 MathType 公式",
        }
    finally:
        if cleanup_path:
            try:
                Path(cleanup_path).unlink()
            except OSError:
                pass


def serve(*, token_file: str, model: str) -> int:
    if hasattr(sys.stdin, "reconfigure"):
        sys.stdin.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    engine = MathTypeEngine()
    token = ""
    try:
        token = read_token(token_file)
    except Exception as exc:
        log("token unavailable:", exc)
    log("sidecar ready (model=%s, token=%s)" % (model, "loaded" if token else "MISSING"))

    try:
        while True:
            line = sys.stdin.readline()
            if not line:
                break
            line = line.strip()
            if not line:
                continue
            try:
                request = json.loads(line)
            except ValueError:
                log("dropped malformed request line")
                continue
            request_id = request.get("id")
            action = str(request.get("action") or "run").lower()
            if action in ("quit", "shutdown", "exit"):
                respond(id=request_id, ok=True, message="bye")
                break
            if action == "warmup":
                # Pay MathType engine start-up now instead of on the first formula.
                try:
                    engine.connect()
                    respond(id=request_id, ok=True, message="warm")
                except Exception as exc:
                    log("warmup failed:", exc)
                    respond(id=request_id, ok=False, error=str(exc))
                continue
            try:
                if not token:
                    token = read_token(token_file)
                response = run_request(request, engine, token=token, model=model)
                response["id"] = request_id
                respond(**response)
            except RunError as exc:
                log("request failed:", exc)
                payload = {"id": request_id, "ok": False, "error": str(exc)}
                if exc.latex:
                    payload["latex"] = exc.latex
                if exc.conf is not None:
                    payload["conf"] = exc.conf
                respond(**payload)
            except Exception as exc:  # keep the sidecar alive for the next request
                log("request crashed:", exc)
                respond(id=request_id, ok=False, error="内部错误：%s" % exc)
    finally:
        engine.close()
        log("sidecar exiting")
    return 0


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(prog=PROG)
    sub = parser.add_subparsers(dest="command", required=True)

    serve_parser = sub.add_parser("serve", help="run the JSON-lines sidecar")
    serve_parser.add_argument("--token-file", default="")
    serve_parser.add_argument("--model", default="turbo", choices=["turbo", "standard"])

    run_parser = sub.add_parser("run", help="one-shot run, for manual testing")
    run_parser.add_argument("--mode", default="clipboard", choices=["clipboard", "latex", "image"])
    run_parser.add_argument("--latex", default="")
    run_parser.add_argument("--image-path", default="")
    run_parser.add_argument("--token-file", default="")
    run_parser.add_argument("--model", default="turbo", choices=["turbo", "standard"])

    args = parser.parse_args(argv)
    if args.command == "serve":
        return serve(token_file=args.token_file, model=args.model)

    request: dict = {"action": "run", "mode": args.mode}
    if args.latex:
        request["latex"] = args.latex
    if args.image_path:
        request["imagePath"] = args.image_path
    engine = MathTypeEngine()
    try:
        token = read_token(args.token_file)
        result = run_request(request, engine, token=token, model=args.model)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False, indent=2))
        return 1
    finally:
        engine.close()


if __name__ == "__main__":
    raise SystemExit(main())