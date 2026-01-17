import base64
import io
import json
import sys
import time
from dataclasses import dataclass
from typing import Any, Optional


MCP_PROTOCOL_VERSION = "2024-11-05"


@dataclass(frozen=True)
class WindowTarget:
    hwnd_class: Optional[str] = None
    title_contains: Optional[str] = None
    include_window_frame: bool = False


def _json_dumps(obj: Any) -> str:
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"))


def _write(obj: Any) -> None:
    sys.stdout.write(_json_dumps(obj) + "\n")
    sys.stdout.flush()


def _error(id_: Any, code: int, message: str, data: Any = None) -> None:
    payload: dict[str, Any] = {"jsonrpc": "2.0", "id": id_, "error": {"code": code, "message": message}}
    if data is not None:
        payload["error"]["data"] = data
    _write(payload)


def _result(id_: Any, result: Any) -> None:
    _write({"jsonrpc": "2.0", "id": id_, "result": result})


def _notify(method: str, params: Any = None) -> None:
    payload: dict[str, Any] = {"jsonrpc": "2.0", "method": method}
    if params is not None:
        payload["params"] = params
    _write(payload)


def _tool_list() -> dict[str, Any]:
    return {
        "tools": [
            {
                "name": "capture",
                "description": "Capture a screenshot (window client area by default). Returns an image for the LLM to see.",
                "inputSchema": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "target": {"type": "string", "enum": ["window", "screen"], "default": "window"},
                        "hwnd_class": {"type": "string"},
                        "title_contains": {"type": "string"},
                        "include_window_frame": {"type": "boolean", "default": False},
                        "format": {"type": "string", "enum": ["jpeg", "png"], "default": "jpeg"},
                        "quality": {"type": "integer", "minimum": 1, "maximum": 95, "default": 65},
                        "max_width": {"type": "integer", "minimum": 64, "maximum": 4096, "default": 1280},
                        "max_height": {"type": "integer", "minimum": 64, "maximum": 4096, "default": 720},
                    },
                },
            },
            {
                "name": "click",
                "description": "Click at normalized coordinates (x,y in [0..1]) relative to the target (window client area by default).",
                "inputSchema": {
                    "type": "object",
                    "additionalProperties": False,
                    "required": ["x", "y"],
                    "properties": {
                        "target": {"type": "string", "enum": ["window", "screen"], "default": "window"},
                        "hwnd_class": {"type": "string"},
                        "title_contains": {"type": "string"},
                        "include_window_frame": {"type": "boolean", "default": False},
                        "mode": {"type": "string", "enum": ["postmessage", "pydirectinput"], "default": "postmessage"},
                        "x": {"type": "number", "minimum": 0, "maximum": 1},
                        "y": {"type": "number", "minimum": 0, "maximum": 1},
                        "button": {"type": "string", "enum": ["left", "right", "middle"], "default": "left"},
                        "clicks": {"type": "integer", "minimum": 1, "maximum": 10, "default": 1},
                        "interval_ms": {"type": "integer", "minimum": 0, "maximum": 5000, "default": 60},
                    },
                },
            },
            {
                "name": "key",
                "description": "Send a key to the target window (default postmessage).",
                "inputSchema": {
                    "type": "object",
                    "additionalProperties": False,
                    "required": ["key"],
                    "properties": {
                        "hwnd_class": {"type": "string"},
                        "title_contains": {"type": "string"},
                        "mode": {"type": "string", "enum": ["postmessage", "pydirectinput"], "default": "postmessage"},
                        "key": {"type": "string"},
                        "action": {"type": "string", "enum": ["press", "down", "up"], "default": "press"},
                        "repeat": {"type": "integer", "minimum": 1, "maximum": 50, "default": 1},
                        "interval_ms": {"type": "integer", "minimum": 0, "maximum": 5000, "default": 35},
                    },
                },
            },
            {
                "name": "wait",
                "description": "Sleep for the specified duration.",
                "inputSchema": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "ms": {"type": "integer", "minimum": 0, "maximum": 600000, "default": 250}
                    },
                },
            },
        ]
    }


def _try_load_default_hwnd_class() -> Optional[str]:
    try:
        from config import config as ok_config  # type: ignore

        windows_cfg = ok_config.get("windows") if isinstance(ok_config, dict) else None
        if isinstance(windows_cfg, dict):
            hwnd_class = windows_cfg.get("hwnd_class")
            if isinstance(hwnd_class, str) and hwnd_class.strip():
                return hwnd_class.strip()
        return None
    except Exception:
        return None


def _find_window(target: WindowTarget) -> Optional[int]:
    try:
        import win32gui  # type: ignore
    except Exception:
        return None

    matches: list[int] = []

    def enum_cb(hwnd: int, _extra: Any) -> None:
        try:
            if not win32gui.IsWindow(hwnd):
                return
            if not win32gui.IsWindowVisible(hwnd):
                return
            if target.hwnd_class:
                if win32gui.GetClassName(hwnd) != target.hwnd_class:
                    return
            if target.title_contains:
                title = win32gui.GetWindowText(hwnd) or ""
                if target.title_contains.lower() not in title.lower():
                    return
            matches.append(hwnd)
        except Exception:
            return

    try:
        win32gui.EnumWindows(enum_cb, None)
    except Exception:
        return None

    if not matches:
        return None

    def area(hwnd: int) -> int:
        try:
            l, t, r, b = win32gui.GetWindowRect(hwnd)
            return max(0, r - l) * max(0, b - t)
        except Exception:
            return 0

    return max(matches, key=area)


def _get_capture_rect(target: str, window_target: WindowTarget) -> tuple[int, int, int, int]:
    if target == "screen":
        try:
            import win32api  # type: ignore

            width = int(win32api.GetSystemMetrics(0))
            height = int(win32api.GetSystemMetrics(1))
            return 0, 0, width, height
        except Exception:
            return 0, 0, 1920, 1080

    try:
        import win32gui  # type: ignore
    except Exception as e:
        raise RuntimeError("pywin32 is required for window capture") from e

    hwnd = _find_window(window_target)
    if hwnd is None:
        raise RuntimeError("target window not found")

    if window_target.include_window_frame:
        l, t, r, b = win32gui.GetWindowRect(hwnd)
        return int(l), int(t), int(r), int(b)

    # client area rect (no borders/title bar)
    cl, ct, cr, cb = win32gui.GetClientRect(hwnd)
    left, top = win32gui.ClientToScreen(hwnd, (cl, ct))
    right, bottom = win32gui.ClientToScreen(hwnd, (cr, cb))
    return int(left), int(top), int(right), int(bottom)


def _pil_grab(bbox: tuple[int, int, int, int]):
    try:
        from PIL import ImageGrab  # type: ignore
    except Exception as e:
        raise RuntimeError("pillow is required for capture") from e

    # all_screens is important when the window is not on primary monitor
    return ImageGrab.grab(bbox=bbox, all_screens=True)


def _encode_image(img, fmt: str, quality: int, max_width: int, max_height: int) -> tuple[str, str, int, int]:
    try:
        from PIL import Image  # type: ignore
    except Exception as e:
        raise RuntimeError("pillow is required for capture") from e

    if max_width > 0 and max_height > 0:
        copy = img.copy()
        copy.thumbnail((max_width, max_height))
        img = copy

    buf = io.BytesIO()
    if fmt == "png":
        img.save(buf, format="PNG")
        mime = "image/png"
    else:
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        img.save(buf, format="JPEG", quality=int(quality), optimize=True)
        mime = "image/jpeg"

    data = base64.b64encode(buf.getvalue()).decode("ascii")
    return data, mime, int(img.width), int(img.height)


def _tool_capture(args: dict[str, Any]) -> dict[str, Any]:
    target = (args.get("target") or "window").lower()
    fmt = (args.get("format") or "jpeg").lower()
    quality = int(args.get("quality") or 65)
    max_width = int(args.get("max_width") or 1280)
    max_height = int(args.get("max_height") or 720)

    hwnd_class = args.get("hwnd_class") or _try_load_default_hwnd_class()
    title_contains = args.get("title_contains")
    include_window_frame = bool(args.get("include_window_frame") or False)

    rect = _get_capture_rect(
        target,
        WindowTarget(hwnd_class=str(hwnd_class) if hwnd_class else None, title_contains=title_contains, include_window_frame=include_window_frame),
    )
    img = _pil_grab(rect)
    data, mime, w, h = _encode_image(img, fmt=fmt, quality=quality, max_width=max_width, max_height=max_height)

    l, t, r, b = rect
    return {
        "content": [
            {
                "type": "text",
                "text": f"capture ok: target={target} rect=({l},{t},{r},{b}) size={w}x{h} (x,y are normalized [0..1] within this rect)",
            },
            {"type": "image", "data": data, "mimeType": mime},
        ],
        "isError": False,
        "metadata": {
            "target": target,
            "rect": {"left": l, "top": t, "right": r, "bottom": b},
            "width": w,
            "height": h,
            "mimeType": mime,
        },
    }


def _tool_wait(args: dict[str, Any]) -> dict[str, Any]:
    ms = int(args.get("ms") or 250)
    time.sleep(ms / 1000.0)
    return {"content": [{"type": "text", "text": f"wait ok: {ms}ms"}], "isError": False, "metadata": {"ms": ms}}


def _tool_click(args: dict[str, Any]) -> dict[str, Any]:
    target = (args.get("target") or "window").lower()
    mode = (args.get("mode") or "postmessage").lower()
    button = (args.get("button") or "left").lower()
    clicks = int(args.get("clicks") or 1)
    interval_ms = int(args.get("interval_ms") or 60)

    x = float(args["x"])
    y = float(args["y"])

    hwnd_class = args.get("hwnd_class") or _try_load_default_hwnd_class()
    title_contains = args.get("title_contains")
    include_window_frame = bool(args.get("include_window_frame") or False)

    rect = _get_capture_rect(
        target,
        WindowTarget(hwnd_class=str(hwnd_class) if hwnd_class else None, title_contains=title_contains, include_window_frame=include_window_frame),
    )
    l, t, r, b = rect
    width = max(1, r - l)
    height = max(1, b - t)
    px = int(round(l + x * width))
    py = int(round(t + y * height))

    if mode == "pydirectinput":
        try:
            import pydirectinput  # type: ignore
        except Exception as e:
            raise RuntimeError("pydirectinput is required for click mode=pydirectinput") from e

        pydirectinput.moveTo(px, py)
        for i in range(clicks):
            if button == "right":
                pydirectinput.click(button="right")
            elif button == "middle":
                pydirectinput.click(button="middle")
            else:
                pydirectinput.click(button="left")
            if i + 1 < clicks and interval_ms > 0:
                time.sleep(interval_ms / 1000.0)
    else:
        try:
            import win32api  # type: ignore
            import win32con  # type: ignore
            import win32gui  # type: ignore
        except Exception as e:
            raise RuntimeError("pywin32 is required for click mode=postmessage") from e

        hwnd = _find_window(WindowTarget(hwnd_class=str(hwnd_class) if hwnd_class else None, title_contains=title_contains, include_window_frame=include_window_frame))
        if hwnd is None:
            raise RuntimeError("target window not found")

        # WM_*BUTTON* expects client coordinates in lParam
        cx, cy = win32gui.ScreenToClient(hwnd, (px, py))
        lparam = (int(cy) << 16) | (int(cx) & 0xFFFF)
        if button == "right":
            down, up, wparam = win32con.WM_RBUTTONDOWN, win32con.WM_RBUTTONUP, win32con.MK_RBUTTON
        elif button == "middle":
            down, up, wparam = win32con.WM_MBUTTONDOWN, win32con.WM_MBUTTONUP, win32con.MK_MBUTTON
        else:
            down, up, wparam = win32con.WM_LBUTTONDOWN, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON

        for i in range(clicks):
            win32api.PostMessage(hwnd, down, wparam, lparam)
            win32api.PostMessage(hwnd, up, 0, lparam)
            if i + 1 < clicks and interval_ms > 0:
                time.sleep(interval_ms / 1000.0)

    return {
        "content": [{"type": "text", "text": f"click ok: target={target} mode={mode} button={button} at x={x:.4f},y={y:.4f} (screen {px},{py}) clicks={clicks}"}],
        "isError": False,
        "metadata": {
            "target": target,
            "mode": mode,
            "button": button,
            "x": x,
            "y": y,
            "screen": {"x": px, "y": py},
            "rect": {"left": l, "top": t, "right": r, "bottom": b},
            "clicks": clicks,
        },
    }


def _vk_from_key(name: str) -> int:
    try:
        import win32con  # type: ignore
        import win32api  # type: ignore
    except Exception as e:
        raise RuntimeError("pywin32 is required for key mode=postmessage") from e

    key = name.strip().lower()
    special: dict[str, int] = {
        "esc": win32con.VK_ESCAPE,
        "escape": win32con.VK_ESCAPE,
        "space": win32con.VK_SPACE,
        "tab": win32con.VK_TAB,
        "enter": win32con.VK_RETURN,
        "return": win32con.VK_RETURN,
        "backspace": win32con.VK_BACK,
        "delete": win32con.VK_DELETE,
        "up": win32con.VK_UP,
        "down": win32con.VK_DOWN,
        "left": win32con.VK_LEFT,
        "right": win32con.VK_RIGHT,
        "shift": win32con.VK_SHIFT,
        "lshift": win32con.VK_LSHIFT,
        "rshift": win32con.VK_RSHIFT,
        "ctrl": win32con.VK_CONTROL,
        "control": win32con.VK_CONTROL,
        "lctrl": win32con.VK_LCONTROL,
        "rctrl": win32con.VK_RCONTROL,
        "alt": win32con.VK_MENU,
        "menu": win32con.VK_MENU,
        "lalt": win32con.VK_LMENU,
        "ralt": win32con.VK_RMENU,
    }
    if key in special:
        return int(special[key])
    if key.startswith("f") and key[1:].isdigit():
        n = int(key[1:])
        if 1 <= n <= 24:
            return int(getattr(win32con, f"VK_F{n}"))
    if len(key) == 1 and key.isalnum():
        return ord(key.upper())

    vk_scan = win32api.VkKeyScan(key)
    if vk_scan == -1:
        raise ValueError(f"Unsupported key: {name}")
    return vk_scan & 0xFF


def _tool_key(args: dict[str, Any]) -> dict[str, Any]:
    mode = (args.get("mode") or "postmessage").lower()
    key = str(args["key"])
    action = (args.get("action") or "press").lower()
    repeat = int(args.get("repeat") or 1)
    interval_ms = int(args.get("interval_ms") or 35)

    hwnd_class = args.get("hwnd_class") or _try_load_default_hwnd_class()
    title_contains = args.get("title_contains")

    if mode == "pydirectinput":
        try:
            import pydirectinput  # type: ignore
        except Exception as e:
            raise RuntimeError("pydirectinput is required for key mode=pydirectinput") from e

        for i in range(repeat):
            if action == "down":
                pydirectinput.keyDown(key)
            elif action == "up":
                pydirectinput.keyUp(key)
            else:
                pydirectinput.press(key)
            if i + 1 < repeat and interval_ms > 0:
                time.sleep(interval_ms / 1000.0)
        return {"content": [{"type": "text", "text": f"key ok: mode={mode} key={key} action={action} repeat={repeat}"}], "isError": False}

    try:
        import win32api  # type: ignore
        import win32con  # type: ignore
    except Exception as e:
        raise RuntimeError("pywin32 is required for key mode=postmessage") from e

    hwnd = _find_window(WindowTarget(hwnd_class=str(hwnd_class) if hwnd_class else None, title_contains=title_contains))
    if hwnd is None:
        raise RuntimeError("target window not found")

    vk = _vk_from_key(key)
    down_msg, up_msg = win32con.WM_KEYDOWN, win32con.WM_KEYUP

    for i in range(repeat):
        if action in ("press", "down"):
            win32api.PostMessage(hwnd, down_msg, vk, 0)
        if action in ("press", "up"):
            win32api.PostMessage(hwnd, up_msg, vk, 0)
        if i + 1 < repeat and interval_ms > 0:
            time.sleep(interval_ms / 1000.0)

    return {"content": [{"type": "text", "text": f"key ok: mode={mode} key={key} action={action} repeat={repeat}"}], "isError": False}


def _handle_call(name: str, args: Any) -> dict[str, Any]:
    if args is None:
        args = {}
    if not isinstance(args, dict):
        raise ValueError("arguments must be an object")

    if name == "capture":
        return _tool_capture(args)
    if name == "click":
        return _tool_click(args)
    if name == "key":
        return _tool_key(args)
    if name == "wait":
        return _tool_wait(args)
    raise ValueError(f"unknown tool: {name}")


def _handle_request(req: dict[str, Any]) -> None:
    if "method" not in req:
        return
    method = req.get("method")
    id_ = req.get("id")
    params = req.get("params") or {}

    if method == "initialize":
        client_version = params.get("protocolVersion") or MCP_PROTOCOL_VERSION
        _result(
            id_,
            {
                "protocolVersion": client_version,
                "capabilities": {
                    "tools": {"listChanged": False},
                    "resources": {"subscribe": False, "listChanged": False},
                    "prompts": {"listChanged": False},
                },
                "serverInfo": {"name": "ok-ww-mcp", "version": "0.1.0"},
            },
        )
        return

    # Some clients send this as a notification (no id)
    if method in ("initialized", "notifications/initialized"):
        return

    if method == "tools/list":
        _result(id_, _tool_list())
        return

    if method == "resources/list":
        _result(id_, {"resources": []})
        return

    if method == "prompts/list":
        _result(id_, {"prompts": []})
        return

    if method == "tools/call":
        try:
            name = params.get("name")
            arguments = params.get("arguments")
            if not isinstance(name, str) or not name:
                raise ValueError("params.name is required")
            result = _handle_call(name, arguments)
            _result(id_, result)
        except Exception as e:
            _error(id_, -32000, str(e))
        return

    if id_ is not None:
        _error(id_, -32601, f"Method not found: {method}")


def main() -> None:
    stdin_buffer = getattr(sys.stdin, "buffer", None)
    if stdin_buffer is None:
        stream = (s.encode("utf-8") for s in sys.stdin)
    else:
        stream = stdin_buffer

    for raw_line in stream:
        try:
            line = raw_line.decode("utf-8-sig", errors="replace").strip()
        except Exception:
            continue
        if not line:
            continue

        try:
            req = json.loads(line)
            if isinstance(req, dict):
                _handle_request(req)
        except Exception:
            continue


if __name__ == "__main__":
    main()
