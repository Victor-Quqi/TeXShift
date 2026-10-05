"""Renders the README demo scenes to GIF.

Each scene is an HTML page whose animation is a pure function of time. This script
seeks it frame by frame in headless Chromium, then encodes the frames with ffmpeg.

Requirements: Windows, Python Playwright with Chromium (`python -m playwright install chromium`),
ffmpeg on PATH, and the vendored MathJax restored by `setup-mathjax.ps1`. When OneNote is
installed, its title bar icon is extracted into the ignored `.cache/` folder; otherwise the
scenes fall back to a traced icon.

Usage:
    python tools/readme-demo/render.py                 # all scenes
    python tools/readme-demo/render.py roundtrip       # selected scenes
    python tools/readme-demo/render.py --frames-only   # keep PNG frames for inspection
"""

import argparse
import ctypes
import shutil
import subprocess
import sys
import tempfile
import winreg
from ctypes import wintypes
from pathlib import Path

from PIL import Image
from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[1]
OUT_DIR = REPO / "docs" / "images"

# GIF delays are stored in centiseconds; 30 ms frames keep playback evenly paced.
FRAME_INTERVAL = 0.03

SCENES = ["roundtrip", "math-mermaid", "selection"]

ICON_CACHE = HERE / ".cache" / "onenote.png"
ONENOTE_APP_PATH = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\onenote.exe"
DEFAULT_ONENOTE_EXE = r"C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE"


class ICONINFO(ctypes.Structure):
    _fields_ = [("fIcon", wintypes.BOOL), ("xHotspot", wintypes.DWORD), ("yHotspot", wintypes.DWORD),
                ("hbmMask", wintypes.HBITMAP), ("hbmColor", wintypes.HBITMAP)]


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
                ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
                ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
                ("biClrImportant", wintypes.DWORD)]


def onenote_exe():
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, ONENOTE_APP_PATH) as key:
            path = Path(winreg.QueryValue(key, None).strip('"'))
    except OSError:
        path = Path(DEFAULT_ONENOTE_EXE)
    return path if path.is_file() else None


def cache_app_icon(size=256):
    """Extracts the flat title bar icon from ONENOTE.EXE; the Start tile PNGs carry a drop shadow."""
    exe = onenote_exe()
    if exe is None:
        print("OneNote not found; scenes use the traced app icon")
        return
    user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
    user32.PrivateExtractIconsW.argtypes = [wintypes.LPCWSTR, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                            ctypes.POINTER(wintypes.HICON), ctypes.POINTER(wintypes.UINT),
                                            wintypes.UINT, wintypes.UINT]
    user32.GetIconInfo.argtypes = [wintypes.HICON, ctypes.POINTER(ICONINFO)]
    user32.GetDC.argtypes = [wintypes.HWND]
    user32.GetDC.restype = wintypes.HDC
    user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
    user32.DestroyIcon.argtypes = [wintypes.HICON]
    gdi32.GetDIBits.argtypes = [wintypes.HDC, wintypes.HBITMAP, wintypes.UINT, wintypes.UINT, ctypes.c_void_p,
                                ctypes.POINTER(BITMAPINFOHEADER), wintypes.UINT]
    gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]

    icon, icon_id = wintypes.HICON(), wintypes.UINT()
    if user32.PrivateExtractIconsW(str(exe), 0, size, size, ctypes.byref(icon), ctypes.byref(icon_id), 1, 0) < 1:
        print(f"No icon in {exe}; scenes use the traced app icon")
        return
    info = ICONINFO()
    user32.GetIconInfo(icon, ctypes.byref(info))
    header = BITMAPINFOHEADER(biSize=ctypes.sizeof(BITMAPINFOHEADER), biWidth=size, biHeight=-size,
                              biPlanes=1, biBitCount=32)
    pixels = ctypes.create_string_buffer(size * size * 4)
    dc = user32.GetDC(None)
    gdi32.GetDIBits(dc, info.hbmColor, 0, size, pixels, ctypes.byref(header), 0)
    user32.ReleaseDC(None, dc)
    gdi32.DeleteObject(info.hbmColor)
    gdi32.DeleteObject(info.hbmMask)
    user32.DestroyIcon(icon)

    ICON_CACHE.parent.mkdir(exist_ok=True)
    Image.frombuffer("RGBA", (size, size), pixels.raw, "raw", "BGRA", 0, 1).save(ICON_CACHE)


def capture(page, scene, frames_dir, scale):
    page.goto((HERE / f"{scene}.html").as_uri() + "?capture")
    page.wait_for_function("window.__ready === true", timeout=60_000)
    duration = page.evaluate("window.__demo.duration")
    window = page.locator(".window")
    count = int(round(duration / FRAME_INTERVAL))
    for i in range(count):
        page.evaluate(f"window.__demo.seek({i * FRAME_INTERVAL:.4f})")
        window.screenshot(path=str(frames_dir / f"{i:05d}.png"), animations="disabled", scale="device")
    print(f"{scene}: {count} frames, {duration:.2f}s at {scale}x")
    return count


def trailing_hold(frames_dir, count):
    """Number of identical frames at the end; mpdecimate keeps only the first of them."""
    last = (frames_dir / f"{count - 1:05d}.png").read_bytes()
    first = count - 1
    while first > 0 and (frames_dir / f"{first - 1:05d}.png").read_bytes() == last:
        first -= 1
    return count - first


def encode(frames_dir, count, target):
    # mpdecimate drops exact repeats so static holds cost one frame with a long delay.
    # The muxer cannot infer the delay of the final frame, so it is passed explicitly.
    graph = (
        "mpdecimate=hi=1:lo=1:frac=0,split[a][b];"
        "[a]palettegen=max_colors=256:stats_mode=full[p];"
        "[b][p]paletteuse=dither=bayer:bayer_scale=5:diff_mode=rectangle"
    )
    subprocess.run(
        [
            "ffmpeg", "-v", "error", "-y",
            "-framerate", f"{1 / FRAME_INTERVAL:.6f}",
            "-i", str(frames_dir / "%05d.png"),
            "-filter_complex", graph,
            "-fps_mode", "vfr",
            "-loop", "0",
            "-final_delay", str(round(trailing_hold(frames_dir, count) * FRAME_INTERVAL * 100)),
            str(target),
        ],
        check=True,
    )
    print(f"  -> {target.relative_to(REPO)} ({target.stat().st_size / 1024:.0f} KiB)")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("scenes", nargs="*", default=SCENES)
    parser.add_argument("--scale", type=float, default=1.5, help="device pixel ratio of the captured frames")
    parser.add_argument("--frames-only", action="store_true", help="keep PNG frames instead of encoding a GIF")
    args = parser.parse_args()

    unknown = [s for s in args.scenes if s not in SCENES]
    if unknown:
        sys.exit(f"Unknown scene(s): {', '.join(unknown)}")

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    cache_app_icon()
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(viewport={"width": 1200, "height": 900}, device_scale_factor=args.scale)
        for scene in args.scenes:
            frames_dir = Path(tempfile.mkdtemp(prefix=f"texshift-{scene}-"))
            try:
                count = capture(page, scene, frames_dir, args.scale)
                if args.frames_only:
                    print(f"  frames kept in {frames_dir}")
                    continue
                encode(frames_dir, count, OUT_DIR / f"{scene}.gif")
            finally:
                if not args.frames_only:
                    shutil.rmtree(frames_dir, ignore_errors=True)
        browser.close()


if __name__ == "__main__":
    main()
