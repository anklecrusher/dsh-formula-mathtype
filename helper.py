#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Windows sidecar for @local/dsh-formula-mathtype.

One request inserts one formula into the *active* Word document:

1. read the formula image from the Windows clipboard (or a file path),
2. POST it to the SimpleTex formula-recognition open API and take ``res.latex``,
3. normalize the LaTeX into the subset MathType's TeX parser accepts
   (``aligned`` -> ``array``; a tall left brace in the source image is restored
   with ``\\left\\{...\\right.``),
4. hand it to ``injector/MathTypeInjector.exe``, which creates a real MathType
   OLE object (``Equation.DSMT4``) at the caret through Word COM and feeds the
   equation server its ``TeX Input Language`` clipboard format.

Why the injector is a separate C# executable: creating the OLE object and
talking to its ``IDataObject`` needs COM interfaces that MathType's SDK
``MTXFormEqn`` cannot provide in the text -> equation direction (it is an
export API; used backwards, Word embeds the literal LaTeX text, which appears
as plain source or as a black block). See injector/MathTypeInjector.cs. The exe
is compiled from source on first use with the .NET Framework compiler that
ships with Windows; no Office PIA or prebuilt binary is required.

Protocol (``serve`` mode, one JSON object per line on stdin/stdout)::

    -> {"id": 1, "action": "run", "mode": "clipboard"}
    -> {"id": 2, "action": "run", "mode": "latex", "latex": "\\frac{a}{b}"}
    -> {"id": 3, "action": "run", "mode": "image", "imagePath": "C:/x.png"}
    <- {"id": 1, "ok": true, "latex": "...", "raw_latex": "...", "conf": 0.95, "elapsedMs": 1234}
    <- {"id": 2, "ok": false, "error": "..."}

The token is read from the file configured with ``--token-file`` (default:
``%LOCALAPPDATA%\\dsh-formula-mathtype\\simpletex.token``). It is never printed
or written anywhere else.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
import time
from pathlib import Path

PROG = "formula-mathtype"
BASE_DIR = Path(__file__).resolve().parent
INJECTOR_DIR = BASE_DIR / "injector"
INJECTOR_SRC = INJECTOR_DIR / "MathTypeInjector.cs"
INJECTOR_EXE = INJECTOR_DIR / "MathTypeInjector.exe"

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

CREATE_NO_WINDOW = 0x08000000
INJECTOR_TIMEOUT_SECONDS = 120


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
# LaTeX normalization for MathType's TeX input parser
# --------------------------------------------------------------------------

_MULTILINE_ENVS = {"aligned", "align", "align*", "gathered", "split", "alignat", "alignat*"}
_ENV_RE = re.compile(
    r"\\begin\{(?P<env>[A-Za-z*]+)\}(?:\[[^\]]*\])?(?:\{[^{}]*\})?(?P<body>.*?)\\end\{(?P=env)\}",
    re.S,
)


def _strip_wrappers(latex: str) -> str:
    text = latex.strip()
    while True:
        if len(text) >= 4 and (
            (text.startswith("$$") and text.endswith("$$"))
            or (text.startswith("\\[") and text.endswith("\\]"))
            or (text.startswith("\\(") and text.endswith("\\)"))
        ):
            text = text[2:-2].strip()
        elif len(text) >= 2 and text.startswith("$") and text.endswith("$"):
            text = text[1:-1].strip()
        else:
            return text


def _split_rows(body: str) -> list[str]:
    return [row.strip() for row in re.split(r"\\\\", body) if row.strip()]


def normalize_latex(latex: str, *, brace_hint: bool | None = None) -> str:
    """Rewrite LaTeX into constructs MathType's "TeX Input Language" accepts.

    Verified against MathType 7: ``array``/``matrix``/``cases``/``\\left\\{...\\right.``
    are understood, while ``aligned``/``align``/``split``/``gathered`` are not and
    would render as red error glyphs.
    """
    text = _strip_wrappers(latex)
    text = text.replace("\r", " ").replace("\n", " ")
    text = re.sub(
        r"\\(?:nonumber|notag|displaystyle|textstyle|scriptstyle|scriptscriptstyle)\b", "", text
    )
    text = re.sub(r"\\(?:label|tag)\{[^{}]*\}", "", text)
    text = re.sub(r"\\(?:begin|end)\{(?:equation\*?|displaymath|math)\}", "", text).strip()

    match = _ENV_RE.search(text)
    if match is not None:
        env = match.group("env")
        if env in _MULTILINE_ENVS:
            rows = _split_rows(match.group("body"))
            joined = " \\\\ ".join(row.replace("&", "") for row in rows)
            if brace_hint:
                return r"\left\{\begin{array}{l}" + joined + r"\end{array}\right."
            return r"\begin{array}{l}" + joined + r"\end{array}"
        if env == "cases":
            rows = _split_rows(match.group("body"))
            # MathType's cases template is single-column; a second column is laid
            # out with spacing instead of an alignment tab.
            joined = " \\\\ ".join(row.replace("&", r"\quad ") if "&" in row else row for row in rows)
            return r"\begin{cases}" + joined + r"\end{cases}"

    if "\\\\" in text and "\\begin{" not in text:
        rows = _split_rows(text)
        joined = " \\\\ ".join(row.replace("&", "") for row in rows)
        if brace_hint:
            return r"\left\{\begin{array}{l}" + joined + r"\end{array}\right."
        return r"\begin{array}{l}" + joined + r"\end{array}"
    return text


def detect_left_brace(image_path: str) -> bool:
    """True when the formula image starts with one tall curly brace.

    SimpleTex often describes a braced equation system as ``aligned`` and drops
    the brace; this heuristic recovers it. A brace is a tall, narrow dark shape
    whose column coverage is most of the formula height.
    """
    try:
        import numpy as np
        from PIL import Image
    except ImportError:
        return False
    try:
        image = Image.open(image_path).convert("L")
        array = np.asarray(image)
        dark = array < 128
        rows = np.where(dark.any(axis=1))[0]
        cols = np.where(dark.any(axis=0))[0]
        if rows.size == 0 or cols.size == 0:
            return False
        y0, y1, x0, x1 = int(rows[0]), int(rows[-1]), int(cols[0]), int(cols[-1])
        height, width = y1 - y0 + 1, x1 - x0 + 1
        if height < 20 or width < 20:
            return False
        left_width = max(2, int(width * 0.12))
        region = dark[y0 : y1 + 1, x0 : x0 + left_width]
        column_coverage = region.sum(axis=0)
        return bool((column_coverage >= 0.55 * height).any())
    except Exception:
        return False


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
    return latex, conf


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
# Insertion: build/run the C# OLE injector
# --------------------------------------------------------------------------

def find_csc() -> str:
    windir = os.environ.get("WINDIR", r"C:\Windows")
    candidates = [
        Path(windir) / "Microsoft.NET" / "Framework64" / "v4.0.30319" / "csc.exe",
        Path(windir) / "Microsoft.NET" / "Framework" / "v4.0.30319" / "csc.exe",
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    raise RunError("找不到 .NET Framework 编译器 csc.exe，无法构建 MathType 注入器。")


def ensure_injector() -> Path:
    if INJECTOR_EXE.exists() and INJECTOR_EXE.stat().st_size > 0:
        return INJECTOR_EXE
    if not INJECTOR_SRC.exists():
        raise RunError("找不到注入器源码：%s" % INJECTOR_SRC)
    INJECTOR_DIR.mkdir(parents=True, exist_ok=True)
    compiler = find_csc()
    try:
        proc = subprocess.run(
            [compiler, "-nologo", "-target:exe", "-out:" + str(INJECTOR_EXE), str(INJECTOR_SRC)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=180,
            creationflags=CREATE_NO_WINDOW,
        )
    except subprocess.TimeoutExpired:
        raise RunError("构建 MathType 注入器超时。") from None
    if not INJECTOR_EXE.exists():
        detail = ((proc.stdout or "") + (proc.stderr or "")).strip()[-800:]
        raise RunError("构建 MathType 注入器失败：%s" % (detail or "csc 没有产生 exe"))
    log("built injector:", INJECTOR_EXE)
    return INJECTOR_EXE


def inject_latex(latex: str) -> dict:
    exe = ensure_injector()
    fd, temp_name = tempfile.mkstemp(prefix="dsh-formula-", suffix=".tex")
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        temp_path.write_text(latex, encoding="utf-8")
        try:
            proc = subprocess.run(
                [str(exe), "--mode", "insert", "--file", str(temp_path), "--format", "tex"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=INJECTOR_TIMEOUT_SECONDS,
                creationflags=CREATE_NO_WINDOW,
            )
        except subprocess.TimeoutExpired:
            raise RunError("插入 Word 超时（%d 秒）。" % INJECTOR_TIMEOUT_SECONDS) from None
    finally:
        temp_path.unlink(missing_ok=True)

    payload = None
    for line in reversed([line.strip() for line in (proc.stdout or "").splitlines() if line.strip()]):
        try:
            payload = json.loads(line)
        except ValueError:
            continue
        break
    if proc.returncode != 0 or not isinstance(payload, dict) or payload.get("ok") is not True:
        detail = ""
        if isinstance(payload, dict):
            detail = str(payload.get("error") or "")
        if not detail:
            stderr_lines = [line for line in (proc.stderr or "").splitlines() if line.strip()]
            detail = stderr_lines[-1] if stderr_lines else "未知错误"
        raise RunError("插入 Word 失败：%s" % detail, latex=latex)
    try:
        width = float(payload.get("width") or 0)
        height = float(payload.get("height") or 0)
    except (TypeError, ValueError):
        width = height = 0.0
    return {"width": width, "height": height}


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
# Request handling
# --------------------------------------------------------------------------

def run_request(request: dict, *, token: str, model: str) -> dict:
    started = time.time()
    mode = str(request.get("mode") or "").strip().lower()
    if not mode:
        mode = "latex" if request.get("latex") else "clipboard"

    cleanup_path: str | None = None
    try:
        raw_latex: str
        conf: float | None = None
        brace: bool | None = None

        if mode == "latex":
            raw_latex = _strip_wrappers(str(request.get("latex") or ""))
            if not raw_latex:
                raise RunError("mode=latex 时必须提供 latex 文本。")
            latex = normalize_latex(raw_latex)
        elif mode in ("image", "clipboard"):
            if mode == "image":
                image_path = str(request.get("imagePath") or "").strip()
                if not image_path:
                    raise RunError("mode=image 时必须提供 imagePath。")
                if not Path(image_path).is_file():
                    raise RunError("找不到图片文件：%s" % image_path)
            else:
                image_path, should_remove = grab_clipboard_image()
                cleanup_path = image_path if should_remove else None
            raw_latex, conf = recognize(image_path, token, model)
            brace = detect_left_brace(image_path)
            latex = normalize_latex(raw_latex, brace_hint=brace)
        else:
            raise RunError("未知的 mode：%s" % mode)

        try:
            inserted = inject_latex(latex)
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
            "raw_latex": raw_latex,
            "conf": conf,
            "brace": brace,
            "width": inserted["width"],
            "height": inserted["height"],
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

    token = ""
    try:
        token = read_token(token_file)
    except Exception as exc:
        log("token unavailable:", exc)
    try:
        ensure_injector()
    except Exception as exc:
        log("injector not ready:", exc)
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
                # Build the injector now so the first formula is not delayed.
                try:
                    path = ensure_injector()
                    respond(id=request_id, ok=True, message="warm", injector=str(path))
                except Exception as exc:
                    log("warmup failed:", exc)
                    respond(id=request_id, ok=False, error=str(exc))
                continue
            try:
                if not token:
                    token = read_token(token_file)
                response = run_request(request, token=token, model=model)
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
    serve_parser.add_argument("--model", default="standard", choices=["turbo", "standard"])

    run_parser = sub.add_parser("run", help="one-shot run, for manual testing")
    run_parser.add_argument("--mode", default="clipboard", choices=["clipboard", "latex", "image"])
    run_parser.add_argument("--latex", default="")
    run_parser.add_argument("--image-path", default="")
    run_parser.add_argument("--token-file", default="")
    run_parser.add_argument("--model", default="standard", choices=["turbo", "standard"])

    hotkey_parser = sub.add_parser("hotkey", help="run the global-hotkey companion (B mode)")
    hotkey_parser.add_argument("--hotkey", default="Ctrl+Alt+F")
    hotkey_parser.add_argument("--token-file", default="")
    hotkey_parser.add_argument("--model", default="standard", choices=["turbo", "standard"])

    args = parser.parse_args(argv)
    if args.command == "serve":
        return serve(token_file=args.token_file, model=args.model)

    if args.command == "hotkey":
        token = ""
        try:
            token = read_token(args.token_file)
        except Exception as exc:
            log("token unavailable:", exc)
        try:
            ensure_injector()
        except Exception as exc:
            log("injector build failed:", exc)
        from hotkey_runner import run_hotkey_app

        return run_hotkey_app(
            hotkey=args.hotkey,
            token=token,
            token_file=args.token_file,
            model=args.model,
        )

    request: dict = {"action": "run", "mode": args.mode}
    if args.latex:
        request["latex"] = args.latex
    if args.image_path:
        request["imagePath"] = args.image_path
    try:
        token = read_token(args.token_file)
        result = run_request(request, token=token, model=args.model)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False, indent=2))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())