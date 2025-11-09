import os
import requests
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from py7zr import SevenZipFile
import re
import json
import tempfile
import shutil
import threading
import subprocess
import sys

# Constants
GITHUB_API_URL = "https://api.github.com/repos/FreeCAD/FreeCAD/releases"
VERSION_FILE = "last_version.json"
CONFIG_FILE = "config.json"
DOWNLOADS_DIR = "downloads"

def get_latest_weekly_asset():
    headers = {"Accept": "application/vnd.github.v3+json", "User-Agent": "FreeCAD-Updater"}
    response = requests.get(GITHUB_API_URL, headers=headers, timeout=15)
    response.raise_for_status()
    releases = response.json()
    pattern = re.compile(r"^FreeCAD_weekly-\d{4}\.\d{2}\.\d{2}-Windows-x86_64-py311\.7z$")
    for release in releases:
        for asset in release.get("assets", []):
            name = asset.get("name", "")
            if pattern.match(name):
                return {
                    "name": name,
                    "url": asset.get("browser_download_url")
                }
    return None

def load_last_version():
    if os.path.exists(VERSION_FILE):
        with open(VERSION_FILE, "r", encoding="utf-8") as f:
            try:
                return json.load(f).get("version")
            except Exception:
                return None
    return None

def save_last_version(version):
    with open(VERSION_FILE, "w", encoding="utf-8") as f:
        json.dump({"version": version}, f)

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f)
    except Exception:
        pass

def copy_contents(src_dir, dst_dir):
    for root, dirs, files in os.walk(src_dir):
        rel = os.path.relpath(root, src_dir)
        dest_root = os.path.join(dst_dir, rel) if rel != "." else dst_dir
        os.makedirs(dest_root, exist_ok=True)
        for f in files:
            src_file = os.path.join(root, f)
            dst_file = os.path.join(dest_root, f)
            try:
                shutil.copy2(src_file, dst_file)
            except Exception as e:
                raise Exception(f"Error copying {src_file} -> {dst_file}: {e}")

def download_and_extract(asset, install_dir, progress_callback=None):
    if not install_dir:
        raise Exception("Installation folder not specified.")
    os.makedirs(install_dir, exist_ok=True)
    os.makedirs(DOWNLOADS_DIR, exist_ok=True)

    headers = {"User-Agent": "FreeCAD-Updater"}

    # Try to obtain remote size (HEAD)
    remote_size = 0
    try:
        head = requests.head(asset["url"], headers=headers, timeout=15, allow_redirects=True)
        remote_size = int(head.headers.get("content-length", 0) or 0)
    except Exception:
        remote_size = 0

    cached_file = os.path.join(DOWNLOADS_DIR, asset["name"])
    use_cached = False
    if os.path.isfile(cached_file) and remote_size:
        try:
            if os.path.getsize(cached_file) == remote_size:
                use_cached = True
        except Exception:
            use_cached = False

    temp_dir = tempfile.mkdtemp(prefix="freecad_updater_")
    try:
        # If cached file is valid, reuse it
        if use_cached:
            temp_file = cached_file
            total = remote_size
            downloaded = total
            if progress_callback:
                try:
                    progress_callback(downloaded, total)
                except Exception:
                    pass
        else:
            # Download and save into cache (overwrites existing incomplete file)
            out_path = cached_file
            resp = requests.get(asset["url"], headers=headers, stream=True, timeout=60)
            resp.raise_for_status()
            total = int(resp.headers.get("content-length", 0) or remote_size or 0)
            downloaded = 0
            with open(out_path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if progress_callback:
                            try:
                                progress_callback(downloaded, total)
                            except Exception:
                                pass
            temp_file = out_path

        # Extract into extraction dir
        extract_dir = os.path.join(temp_dir, "extracted")
        os.makedirs(extract_dir, exist_ok=True)

        # 1) Try py7zr first
        py7zr_error = None
        try:
            with SevenZipFile(temp_file, mode='r') as archive:
                archive.extractall(path=extract_dir)
        except Exception as e_py:
            py7zr_error = e_py

            # 2) Fallback: try external 7z.exe
            seven = None
            # 1) If running as PyInstaller onefile/onedir, check the extracted bundle folder (sys._MEIPASS)
            meipass = getattr(sys, "_MEIPASS", None)
            if meipass:
                candidate = os.path.join(meipass, "7z.exe")
                if os.path.isfile(candidate):
                    seven = candidate

            # 2) If running frozen, also check next to the exe (useful with --add-binary in onedir)
            if not seven and getattr(sys, "frozen", False):
                exe_dir = os.path.dirname(sys.executable)
                candidate = os.path.join(exe_dir, "7z.exe")
                if os.path.isfile(candidate):
                    seven = candidate

            # 3) Finally check PATH and common Program Files location
            if not seven:
                seven = shutil.which("7z") or shutil.which("7z.exe")
            if not seven:
                pf = os.environ.get("ProgramFiles", r"C:\Program Files")
                candidate = os.path.join(pf, "7-Zip", "7z.exe")
                if os.path.isfile(candidate):
                    seven = candidate
            if seven:
                try:
                    cmd = [seven, "x", "-y", f"-o{extract_dir}", temp_file]
                    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=300)
                except subprocess.CalledProcessError as ce:
                    raise Exception(f"7z extraction failed: {ce.stderr or ce.stdout}") from ce
            else:
                # 3) Fallback: try Windows Shell (Explorer) if pywin32 available and shell extension for .7z is registered
                try:
                    import time
                    import importlib
                    try:
                        win32com_client = importlib.import_module("win32com.client")
                    except ModuleNotFoundError as mnfe:
                        # Explicitly surface a clear error if pywin32 is not installed
                        raise Exception("pywin32 (win32com) is not installed; cannot use Explorer shell to extract .7z archives.") from mnfe

                    Dispatch = win32com_client.Dispatch
                    shell = Dispatch("Shell.Application")
                    archive_ns = shell.NameSpace(temp_file)
                    if archive_ns is None:
                        raise Exception("Shell cannot open archive (no shell extension for .7z).")
                    dest_ns = shell.NameSpace(extract_dir)
                    dest_ns.CopyHere(archive_ns.Items(), 20)  # 20 = no UI + do not show progress
                    # wait for files to appear (timeout)
                    deadline = time.time() + 120
                    while time.time() < deadline:
                        if any(os.scandir(extract_dir)):
                            break
                        time.sleep(0.5)
                    else:
                        raise Exception("Shell extraction timed out.")
                except Exception as e_shell:
                    msg = str(py7zr_error) if py7zr_error else "unknown py7zr error"
                    raise Exception(
                        "Could not extract archive. py7zr failed and no 7z.exe found. "
                        "If you can extract this file in Explorer, that means a shell extension (e.g. 7‑Zip) is installed — install 7‑Zip or pywin32. "
                        f"py7zr error: {msg}. Shell error: {e_shell}"
                    ) from e_shell

        # Choose root of extracted content
        entries = os.listdir(extract_dir)
        if len(entries) == 1 and os.path.isdir(os.path.join(extract_dir, entries[0])):
            src_root = os.path.join(extract_dir, entries[0])
        else:
            src_root = extract_dir

        copy_contents(src_root, install_dir)

    finally:
        # remove only the temporary extraction folder; keep cached archive in DOWNLOADS_DIR
        try:
            shutil.rmtree(temp_dir)
        except Exception:
            pass

def detect_installed_version(install_dir):
    """
    Try to detect the installed FreeCAD version by running FreeCAD.exe or FreeCADCmd.exe with --version.
    Returns a tuple (version_text, revision) or (None, None).
    """
    if not install_dir:
        return (None, None)
    exe_candidates = [
        os.path.join(install_dir, "bin", "FreeCAD.exe"),
        os.path.join(install_dir, "bin", "FreeCADCmd.exe"),
        os.path.join(install_dir, "FreeCAD.exe"),
        os.path.join(install_dir, "FreeCADCmd.exe"),
    ]
    for exe in exe_candidates:
        if os.path.isfile(exe):
            try:
                p = subprocess.run([exe, "--version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=10)
                out = (p.stdout or "").strip()
                err = (p.stderr or "").strip()
                text = out if out else err
                if not text:
                    continue

                # Extract version (e.g. "FreeCAD 0.21.1")
                m = re.search(r"FreeCAD\s+([^\s,()]+)", text, re.IGNORECASE)
                if m:
                    version = f"FreeCAD {m.group(1)}"
                else:
                    # fallback: first non-empty line
                    version = text.splitlines()[0].strip()

                # Search for revision/commit in various formats
                rev = None
                rev_patterns = [
                    r"Revision[:\s]*([0-9A-Za-z\-]+)",
                    r"\brev(?:ision)?[:\s]*([0-9A-Za-z\-]+)",
                    r"commit[:\s]*([0-9a-f]{7,40})",
                    r"\b([0-9a-f]{7,40})\b",
                ]
                for rp in rev_patterns:
                    rm = re.search(rp, text, re.IGNORECASE)
                    if rm:
                        rev = rm.group(1)
                        break

                return (version, rev)
            except Exception:
                continue
    return (None, None)

class FreeCADUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FreeCAD Weekly Updater")
        self.install_dir = tk.StringVar()

        cfg = load_config()
        if cfg.get("install_dir"):
            self.install_dir.set(cfg.get("install_dir"))

        frm = tk.Frame(root)
        frm.pack(padx=10, pady=10)

        tk.Label(frm, text="FreeCAD installation folder:").grid(row=0, column=0, sticky="w")
        tk.Entry(frm, textvariable=self.install_dir, width=50).grid(row=1, column=0, columnspan=3, pady=(2,8), sticky="w")
        tk.Button(frm, text="Select folder", command=self.select_folder).grid(row=1, column=3, padx=5)
        tk.Button(frm, text="Detect installed version", command=self.update_installed_version_label).grid(row=2, column=3, padx=5)

        self.check_btn = tk.Button(frm, text="Check and update", command=self.check_and_update, width=20)
        self.check_btn.grid(row=2, column=0, pady=8)

        self.progress = ttk.Progressbar(frm, length=400, mode="determinate")
        self.progress.grid(row=3, column=0, columnspan=4, pady=(5,0))

        self.status = tk.Label(frm, text="", anchor="w")
        self.status.grid(row=4, column=0, columnspan=4, sticky="w", pady=(4,0))

        self.installed_label = tk.Label(frm, text="Installed version: (not detected)", anchor="w")
        self.installed_label.grid(row=5, column=0, columnspan=4, sticky="w", pady=(6,0))

        # Do NOT call update_installed_version_label() at startup to avoid any popup or early detection.
        # The user can click "Detect installed version" to populate this field manually.

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.install_dir.set(folder)
            cfg = load_config()
            cfg["install_dir"] = folder
            save_config(cfg)
            # Do not auto-detect at folder select to avoid popups; user should press detect explicitly.
            # self.update_installed_version_label()

    def update_installed_version_label(self):
        install_dir = self.install_dir.get()
        ver, rev = detect_installed_version(install_dir)
        if ver:
            if rev:
                self.installed_label.config(text=f"Installed version: {ver} (rev: {rev})")
            else:
                self.installed_label.config(text=f"Installed version: {ver}")
        else:
            # if detection fails, try reading last_version from config as a fallback
            cfg = load_config()
            last = cfg.get("last_version") or load_last_version()
            if last:
                self.installed_label.config(text=f"Installed version (from record): {last}")
            else:
                self.installed_label.config(text="Installed version: (not detected)")

    def set_ui_busy(self, busy=True):
        state = "disabled" if busy else "normal"
        self.check_btn.config(state=state)

    def update_progress_safe(self, downloaded, total):
        def ui_update():
            if total:
                pct = int(downloaded / total * 100)
                self.progress.config(value=pct)
                self.status.config(text=f"Downloading: {pct}% ({downloaded // 1024} KB / {total // 1024} KB)")
            else:
                # indeterminate if total unknown
                self.progress.config(mode="indeterminate")
                self.progress.start(10)
                self.status.config(text=f"Downloading: {downloaded // 1024} KB")
        self.root.after(0, ui_update)

    def reset_progress_safe(self):
        def ui_reset():
            self.progress.stop()
            self.progress.config(mode="determinate", value=0)
            self.status.config(text="")
        self.root.after(0, ui_reset)

    def run_update_thread(self, asset, install_dir, latest_version):
        def worker():
            try:
                self.update_progress_safe(0, 1)
                download_and_extract(asset, install_dir, progress_callback=self.update_progress_safe)
                save_last_version(latest_version)
                cfg = load_config()
                cfg["install_dir"] = install_dir
                cfg["last_version"] = latest_version
                save_config(cfg)
                # update installed version label after copying files
                self.root.after(0, self.update_installed_version_label)
                self.root.after(0, lambda: messagebox.showinfo("Update complete", f"Updated to version {latest_version}."))
            except requests.HTTPError as he:
                # bind he so it's available when the lambda runs
                self.root.after(0, lambda he=he: messagebox.showerror("HTTP Error", str(he)))
            except Exception as e:
                # bind e so it's available when the lambda runs
                self.root.after(0, lambda e=e: messagebox.showerror("Error", str(e)))
            finally:
                self.root.after(0, self.reset_progress_safe)
                self.root.after(0, lambda: self.set_ui_busy(False))

        t = threading.Thread(target=worker, daemon=True)
        t.start()

    def check_and_update(self):
        try:
            asset = get_latest_weekly_asset()
            if not asset:
                messagebox.showerror("Error", "No weekly build found.")
                return
            latest_version = asset["name"]
            current_saved = load_last_version()

            install_dir = self.install_dir.get()
            if not install_dir:
                messagebox.showerror("Error", "Select the FreeCAD installation folder.")
                return

            # detect installed version (only when user requested)
            installed_ver = detect_installed_version(install_dir)
            if isinstance(installed_ver, tuple):
                v, r = installed_ver
                if v:
                    installed_str = v
                    if r:
                        installed_str += f" (rev: {r})"
                else:
                    installed_str = "(not detected)"
            else:
                installed_str = installed_ver or "(not detected)"

            # Show comparison installed vs available
            msg_lines = []
            msg_lines.append(f"Installed: {installed_str}")
            msg_lines.append(f"Available: {latest_version}")
            if current_saved:
                msg_lines.append(f"(Last downloaded by this app: {current_saved})")
            msg = "\n".join(msg_lines)

            if latest_version == current_saved:
                messagebox.showinfo("Up to date", f"You already downloaded the latest version.\n\n{msg}")
                return

            if not messagebox.askyesno("New version available", f"A new version was detected.\n\n{msg}\n\nDo you want to download and update now?"):
                return

            # start background update
            self.set_ui_busy(True)
            self.run_update_thread(asset, install_dir, latest_version)
        except requests.HTTPError as he:
            messagebox.showerror("HTTP Error", str(he))
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = FreeCADUpdaterApp(root)
    root.mainloop()
