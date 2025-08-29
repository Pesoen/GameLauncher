import os
import sys
import shlex
import json
import ctypes
import subprocess
from pathlib import Path
from configparser import ConfigParser
import tkinter as tk
from tkinter import ttk, messagebox

# ---------------- High-DPI awareness ----------------
try:
    ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
except Exception:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

# ---------------- Paths & config ----------------
BASE = Path(sys.argv[0]).resolve().parent
SELF_EXE = Path(sys.argv[0]).resolve()
SELF_NAME = SELF_EXE.stem
CONF = BASE / f"{SELF_NAME}.conf"

# icon cache folder + metadata file (stored next to the launcher)
CACHE_DIR = BASE / f".{SELF_NAME}"
CACHE_DIR.mkdir(exist_ok=True)
CACHE_META = CACHE_DIR / "cache.json"

# ---------------- Config reading ----------------
def read_config():
    if not CONF.exists():
        messagebox.showerror(
            "Configuration missing",
            f"Required file '{SELF_NAME}.conf' was not found.\nPlease contact an administrator."
        )
        sys.exit(4)
    cp = ConfigParser()
    cp.optionxform = str
    cp.read(CONF, encoding="utf-8")

    title = cp.get("meta", "title", fallback="Launcher").strip()
    window_icon = cp.get("meta", "window_icon", fallback="").strip()

    raw_items = cp.get("apps", "items", fallback="")
    lines = [ln.strip() for ln in raw_items.splitlines() if ln.strip() and not ln.strip().startswith(("#",";"))]
    apps = []
    for ln in lines:
        parts = [p.strip() for p in ln.split("|")]
        exe = parts[0]
        meta = {"args": "", "title": "", "icon": "", "elevated": ""}
        for p in parts[1:]:
            if "=" in p:
                k, v = p.split("=", 1)
                meta[k.strip().lower()] = v.strip()
        apps.append({
            "exe": exe,
            "args": meta["args"],
            "title": meta["title"],
            "icon": meta["icon"],
            "elevated": meta["elevated"].lower() in ("1","true","yes","y","on")
        })
    return title, window_icon, apps

# ---------------- Hidden PowerShell helpers ----------------
CREATE_NO_WINDOW = 0x08000000

def run_powershell_hidden(ps_script: str) -> int:
    # Hidden window, bypass exec policy; no profile; no logo
    try:
        completed = subprocess.run(
            ["powershell", "-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-WindowStyle", "Hidden", "-Command", ps_script],
            creationflags=CREATE_NO_WINDOW
        )
        return completed.returncode
    except Exception:
        return 1

def extract_icon_png_from_exe(exe_path: Path, out_png: Path) -> bool:
    exe_path = str(exe_path)
    out_png = str(out_png)
    ps = rf"""
    try {{
        Add-Type -AssemblyName System.Drawing | Out-Null
        $icon=[System.Drawing.Icon]::ExtractAssociatedIcon('{exe_path}')
        if ($null -eq $icon) {{ exit 2 }}
        $bmp = $icon.ToBitmap()
        $bmp.Save('{out_png}', [System.Drawing.Imaging.ImageFormat]::Png)
        exit 0
    }} catch {{ exit 1 }}
    """
    code = run_powershell_hidden(ps)
    return code == 0 and os.path.exists(out_png)

def convert_ico_to_png(ico_path: Path, out_png: Path) -> bool:
    ico_path = str(ico_path)
    out_png = str(out_png)
    ps = rf"""
    try {{
        Add-Type -AssemblyName System.Drawing | Out-Null
        $ico = New-Object System.Drawing.Icon('{ico_path}')
        $bmp = $ico.ToBitmap()
        $bmp.Save('{out_png}', [System.Drawing.Imaging.ImageFormat]::Png)
        exit 0
    }} catch {{ exit 1 }}
    """
    code = run_powershell_hidden(ps)
    return code == 0 and os.path.exists(out_png)

def extract_icon_ico_from_exe(exe_path: Path, out_ico: Path) -> bool:
    exe_path = str(exe_path)
    out_ico = str(out_ico)
    ps = rf"""
    try {{
        Add-Type -AssemblyName System.Drawing | Out-Null
        $icon=[System.Drawing.Icon]::ExtractAssociatedIcon('{exe_path}')
        if ($null -eq $icon) {{ exit 2 }}
        $fs = New-Object System.IO.FileStream('{out_ico}','Create')
        $icon.Save($fs)
        $fs.Close()
        exit 0
    }} catch {{ exit 1 }}
    """
    code = run_powershell_hidden(ps)
    return code == 0 and os.path.exists(out_ico)

def get_exe_file_description(exe_path: Path) -> str:
    exe_path = str(exe_path)
    ps = rf"""
    try {{
        (Get-Item '{exe_path}').VersionInfo.FileDescription
    }} catch {{ '' }}
    """
    try:
        out = subprocess.check_output(
            ["powershell", "-NoLogo", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps],
            creationflags=CREATE_NO_WINDOW,
            universal_newlines=True, errors="ignore"
        )
        return out.strip()
    except Exception:
        return ""

# ---------------- Caching logic ----------------
def load_cache_meta():
    if CACHE_META.exists():
        try:
            return json.loads(CACHE_META.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}

def save_cache_meta(meta: dict):
    try:
        CACHE_META.write_text(json.dumps(meta, indent=2), encoding="utf-8")
    except Exception:
        pass

def ensure_icons(title, window_icon, apps, root_for_splash=None):
    """Ensure all needed button/window icons are generated in CACHE_DIR.
       Shows a tiny splash if generation is required.
    """
    prev_meta = load_cache_meta()
    conf_mtime = CONF.stat().st_mtime if CONF.exists() else 0
    prev_mtime = prev_meta.get("conf_mtime", 0)

    # Decide targets
    targets = []  # list of (source_spec, out_png_path, mode)
    # Window icon: if meta points to .exe, we create an .ico in cache for the window icon
    window_ico_target = None
    if window_icon:
        p = Path(window_icon)
        if not p.is_absolute():
            p = (BASE / p).resolve()
        if p.suffix.lower() == ".exe":
            window_ico_target = (p, CACHE_DIR / "window_icon.ico")
        # if it's .ico, we can use it directly—no need to cache

    # App icons
    for app in apps:
        exe = Path(app["exe"])
        if not exe.is_absolute():
            exe = (BASE / exe).resolve()
        icon_spec = app.get("icon", "").strip()

        # Resolve what PNG we want to place alongside (inline icon)
        if icon_spec:
            ip = Path(icon_spec)
            if not ip.is_absolute():
                ip = (BASE / ip).resolve()
            if ip.suffix.lower() == ".png":
                # already a png; no caching required
                continue
            elif ip.suffix.lower() == ".ico":
                out_png = CACHE_DIR / f"btn_{ip.stem}.png"
                targets.append(("ico", ip, out_png))
            elif ip.suffix.lower() == ".exe":
                out_png = CACHE_DIR / f"btn_{ip.stem}.png"
                targets.append(("exe", ip, out_png))
        else:
            # No icon specified: extract from the app exe
            out_png = CACHE_DIR / f"btn_{exe.stem}.png"
            targets.append(("exe", exe, out_png))

    # Determine if we need regeneration:
    need_regen = (conf_mtime != prev_mtime)
    if not need_regen:
        # If any target PNG/ICO missing, regen is needed.
        for kind, src, out_png in targets:
            if not out_png.exists():
                need_regen = True
                break
        if window_ico_target and not window_ico_target[1].exists():
            need_regen = True

    if not need_regen:
        return  # All good

    # Show a tiny splash while generating
    splash = None
    if root_for_splash is not None:
        splash = tk.Toplevel(root_for_splash)
        splash.title("Please wait")
        splash.resizable(False, False)
        ttk.Label(splash, text="Generating icons...", padding=12).pack()
        pb = ttk.Progressbar(splash, mode="indeterminate", length=220)
        pb.pack(padx=12, pady=(0, 12))
        pb.start(10)
        # Center over parent
        splash.update_idletasks()
        x = root_for_splash.winfo_x() + (root_for_splash.winfo_width() - splash.winfo_width()) // 2
        y = root_for_splash.winfo_y() + (root_for_splash.winfo_height() - splash.winfo_height()) // 2
        splash.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        splash.update()

    try:
        # Generate window icon ICO if needed
        if window_ico_target:
            src_exe, out_ico = window_ico_target
            try:
                out_ico.unlink(missing_ok=True)
            except Exception:
                pass
            extract_icon_ico_from_exe(src_exe, out_ico)

        # Generate app PNGs
        for kind, src, out_png in targets:
            try:
                out_png.unlink(missing_ok=True)
            except Exception:
                pass
            if kind == "exe":
                extract_icon_png_from_exe(src, out_png)
            elif kind == "ico":
                convert_ico_to_png(src, out_png)

        # Save new conf mtime
        save_cache_meta({"conf_mtime": conf_mtime})
    finally:
        if splash is not None:
            splash.destroy()

# ---------------- Launch helpers ----------------
def launch_normal(exe: Path, args: str, workdir: Path):
    cmd = [str(exe)] + (shlex.split(args, posix=False) if args else [])
    try:
        subprocess.Popen(cmd, cwd=str(workdir if workdir else exe.parent))
    except Exception as e:
        messagebox.showerror("Launch failed", f"Could not launch:\n{exe}\n\n{e}")

def launch_elevated(exe: Path, args: str, workdir: Path):
    ShellExecuteW = ctypes.windll.shell32.ShellExecuteW
    params = args if args else ""
    try:
        rc = ShellExecuteW(None, "runas", str(exe), params, str(workdir if workdir else exe.parent), 1)
        if rc <= 32:
            messagebox.showerror("Launch failed (elevated)",
                                 f"ShellExecuteW error code: {rc}\n\nFile: {exe}")
    except Exception as e:
        messagebox.showerror("Launch failed (elevated)",
                             f"Could not launch elevated:\n{exe}\n\n{e}")

# ---------------- UI ----------------
class LauncherApp:
    def __init__(self, title, window_icon, apps):
        self.title = title
        self.window_icon_spec = window_icon
        self.apps = apps

        self.root = tk.Tk()
        self.root.title(f"{self.title} Launcher")
        self.root.geometry("560x460")
        self.root.minsize(480, 360)

        # DPI scaling
        try:
            self.root.update_idletasks()
            hwnd = self.root.winfo_id()
            dpi = ctypes.windll.user32.GetDpiForWindow(hwnd)
            self.root.tk.call('tk', 'scaling', dpi/72.0)
        except Exception:
            pass

        # Ensure icons (may show splash briefly)
        ensure_icons(self.title, self.window_icon_spec, self.apps, root_for_splash=self.root)

        # Apply window icon (from conf: .ico or .exe; for exe we used cache)
        self.apply_window_icon(self.window_icon_spec)

        # Main frame + scroll
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(outer, highlightthickness=0)
        scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self.list_frame = ttk.Frame(canvas)
        self.list_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.list_frame, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # Build rows
        self.tk_images = []
        self.build_rows()

    def apply_window_icon(self, spec: str):
        if not spec:
            return
        p = Path(spec)
        if not p.is_absolute():
            p = (BASE / p).resolve()
        try:
            if p.suffix.lower() == ".ico" and p.exists():
                self.root.iconbitmap(default=str(p))
            elif p.suffix.lower() == ".exe" and p.exists():
                cached = CACHE_DIR / "window_icon.ico"
                if cached.exists():
                    self.root.iconbitmap(default=str(cached))
        except Exception:
            pass

    def build_rows(self):
        ttk.Label(self.list_frame, text="Application", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=1, sticky="w", padx=(6,4), pady=(0,6)
        )
        ttk.Label(self.list_frame, text="").grid(row=0, column=0)  # icon column
        r = 1

        for app in self.apps:
            exe = Path(app["exe"])
            if not exe.is_absolute():
                exe = (BASE / exe).resolve()

            # Determine title
            title = app["title"].strip()
            if not title:
                desc = get_exe_file_description(exe) if exe.exists() else ""
                title = desc if desc else exe.name

            # Determine icon image to use
            icon_img = self.load_icon_image(app.get("icon",""), exe, size=24)

            # Build row
            lbl_icon = tk.Label(self.list_frame, image=icon_img) if icon_img else tk.Label(self.list_frame, width=2)
            if icon_img: self.tk_images.append(icon_img)
            lbl_icon.grid(row=r, column=0, padx=(0,6), pady=6, sticky="w")

            missing = not exe.exists()
            ttk.Label(self.list_frame, text=(title + (" (missing)" if missing else ""))).grid(row=r, column=1, sticky="w")

            args = app["args"]
            elevated = app["elevated"]

            def make_launch(exe=exe, args=args, elevated=elevated):
                def _go():
                    if missing:
                        return
                    if elevated:
                        launch_elevated(exe, args, exe.parent)
                    else:
                        launch_normal(exe, args, exe.parent)
                return _go

            btn_text = "Not Found" if missing else "Launch"
            state = "disabled" if missing else "normal"
            ttk.Button(self.list_frame, text=btn_text, state=state, command=make_launch()).grid(
                row=r, column=2, padx=6, pady=6, sticky="e"
            )
            r += 1

        self.list_frame.columnconfigure(1, weight=1)

    def load_icon_image(self, icon_spec: str, exe_path: Path, size: int = 24):
        """
        Returns a tk.PhotoImage or None. Prefers PNG.
        Uses CACHE_DIR for generated images to avoid temp churn.
        """
        try:
            if icon_spec:
                p = Path(icon_spec)
                if not p.is_absolute():
                    p = (BASE / p).resolve()
                if p.suffix.lower() == ".png" and p.exists():
                    return tk.PhotoImage(file=str(p))
                elif p.suffix.lower() == ".ico" and p.exists():
                    out_png = CACHE_DIR / f"btn_{p.stem}.png"
                    if out_png.exists():
                        return tk.PhotoImage(file=str(out_png))
                    # should already be generated by ensure_icons; fallback try:
                    if convert_ico_to_png(p, out_png) and out_png.exists():
                        return tk.PhotoImage(file=str(out_png))
                elif p.suffix.lower() == ".exe" and p.exists():
                    out_png = CACHE_DIR / f"btn_{p.stem}.png"
                    if out_png.exists():
                        return tk.PhotoImage(file=str(out_png))
                    if extract_icon_png_from_exe(p, out_png) and out_png.exists():
                        return tk.PhotoImage(file=str(out_png))
            else:
                # no spec: extract from target exe
                out_png = CACHE_DIR / f"btn_{exe_path.stem}.png"
                if out_png.exists():
                    return tk.PhotoImage(file=str(out_png))
                if exe_path.exists() and extract_icon_png_from_exe(exe_path, out_png) and out_png.exists():
                    return tk.PhotoImage(file=str(out_png))
        except Exception:
            pass
        return None

    def run(self):
        self.root.mainloop()

# ---------------- main ----------------
def main():
    if os.name != "nt":
        print("Windows only.")
        sys.exit(2)

    title, window_icon, apps = read_config()
    app = LauncherApp(title, window_icon, apps)
    app.run()

if __name__ == "__main__":
    main()