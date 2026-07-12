"""
Local Windows AI Assistant (Offline) — MVP (Name-based file/folder handling)
=========================================================================

This single-file Python app runs locally on Windows and performs two categories of tasks:
1) Navigating Around (system-level operations)
2) Using Software Systems (operate common apps & files)

What's new in this version
--------------------------
- **You no longer need to type full paths.** Just say file/folder names.
- **Smart search & disambiguation**: If there are multiple matches, it will ask which one.
- **Folder aliases**: "Desktop", "Documents", "Downloads", "Pictures", "Music", "Videos" (and short forms) are expanded automatically.
- Works fully **offline** after dependencies are installed once.

Install (recommended, then run):
    py -m pip install --upgrade pip
    py -m pip install pyautogui pywinauto psutil send2trash pyttsx3 python-docx openpyxl python-pptx pillow PyPDF2 comtypes pycaw pywin32

Run:
    py local_ai.py

Examples you can type (no full paths!):
    create folder Work in Documents
    copy A.txt to Desktop
    move budget.xlsx to Documents
    delete temp.log
    open notepad
    open notes.txt
    open file report.docx
    write doc "Hello from AI" to Documents/hello.docx
    spreadsheet sum 1,2,3 to Documents/sum.xlsx
    presentation title "My Deck" to Desktop/deck.pptx
    set volume 30
    volume up
    brightness down
    play media movie.mp4
    read pdf manual.pdf
    search pdf "AI" in research.pdf
    battery
    storage C:
    shutdown (or restart / sleep / lock / logout)
    alt tab
    help
    quit

"""
from __future__ import annotations
import os
import re
import sys
import time
import shutil
import glob
import subprocess
from dataclasses import dataclass
from typing import Optional, List, Tuple

# Optional deps guarded imports
try:
    import psutil
except Exception:
    psutil = None

try:
    import pyautogui
except Exception:
    pyautogui = None

try:
    from send2trash import send2trash
except Exception:
    send2trash = None
try:
    import pyttsx3
except Exception:
    pyttsx3 = None

# Windows-specific optional libs
try:
    import win32gui  # noqa: F401
    import win32con  # noqa: F401
    import win32process  # noqa: F401
except Exception as e:
    print("Win32 modules not available:", e)

# Volume via PyCAW
try:
    from ctypes import POINTER, cast
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
except Exception:
    AudioUtilities = IAudioEndpointVolume = None

# Office & media libs
try:
    from docx import Document
except Exception:
    Document = None

try:
    from openpyxl import Workbook
except Exception:
    Workbook = None

try:
    from pptx import Presentation
except Exception:
    Presentation = None

try:
    from PIL import Image
except Exception:
    Image = None

try:
    import PyPDF2
except Exception:
    PyPDF2 = None

# ----------------------- Helpers -----------------------

def say(text: str):
    """Speak + print feedback (if pyttsx3 is available)."""
    print(text)
    if pyttsx3:
        try:
            engine = pyttsx3.init()
            engine.setProperty('rate', 180)  # Normal speaking speed
            engine.setProperty('volume', 0.8)  # Slightly lower volume for smoothness
            voices = engine.getProperty('voices')
            if voices and len(voices) > 1:
                engine.setProperty('voice', voices[1].id)  # Use second voice if available (often smoother/female)
            engine.say(text)
            engine.runAndWait()
        except Exception:
            pass


def answer_science(query: str):
    facts = {
        "gravity": "Gravity is the force that attracts objects with mass towards each other.",
        "photosynthesis": "Photosynthesis is the process by which plants use sunlight, CO2, and water to produce glucose and oxygen.",
        "dna": "DNA is the molecule that carries genetic information in living organisms.",
        "evolution": "Evolution is the change in the characteristics of a species over several generations.",
        "quantum": "Quantum physics deals with the behavior of matter and energy at atomic and subatomic scales.",
        "relativity": "Relativity is Einstein's theory describing how space and time are linked for objects moving at constant speeds.",
        "atom": "An atom is the smallest unit of ordinary matter that forms a chemical element.",
        "cell": "A cell is the basic structural and functional unit of all forms of life.",
        "planet": "A planet is a celestial body that orbits a star and has cleared its orbital path of other debris.",
        "star": "A star is a luminous sphere of plasma held together by its own gravity.",
    }
    query_lower = query.lower()
    for key, ans in facts.items():
        if key in query_lower:
            say(ans)
            return
    say("I'm sorry, I don't have information on that science topic. Ask me about gravity, photosynthesis, DNA, or evolution.")


def tell_joke():
    jokes = [
        "Why don't scientists trust atoms? Because they make up everything!",
        "What do you call fake spaghetti? An impasta!",
        "Why did the scarecrow win an award? Because he was outstanding in his field!",
        "What do you get when you cross a snowman and a vampire? Frostbite!",
        "Why don't eggs tell jokes? They'd crack each other up!",
    ]
    import random
    say(random.choice(jokes))


def tell_quote():
    quotes = [
        "The only way to do great work is to love what you do. - Steve Jobs",
        "Believe you can and you're halfway there. - Theodore Roosevelt",
        "The future belongs to those who believe in the beauty of their dreams. - Eleanor Roosevelt",
        "You miss 100% of the shots you don't take. - Wayne Gretzky",
        "The best way to predict the future is to create it. - Peter Drucker",
    ]
    import random
    say(random.choice(quotes))


def get_time():
    import datetime
    say(f"The current time is {datetime.datetime.now().strftime('%I:%M %p')}")


def get_date():
    import datetime
    say(f"Today's date is {datetime.datetime.now().strftime('%B %d, %Y')}")


def calculate_math(query):
    # Remove non-math parts and evaluate safely
    query = re.sub(r"[^\d\+\-\*/\(\)\.\s]", "", query)
    try:
        result = eval(query, {"__builtins__": None}, {})
        say(f"The result is {result}")
    except:
        say("Sorry, I can't calculate that math expression.")


def norm(path: str) -> str:
    return os.path.normpath(os.path.expanduser(path.strip().strip('"')))


def ensure_dir(path: str):
    if path and not os.path.exists(path):
        os.makedirs(path, exist_ok=True)


def exists(path: str) -> bool:
    return os.path.exists(path)


# --------- Name-based search, disambiguation & aliases ---------

HOME = os.path.expanduser("~")
DEFAULT_SEARCH_DIRS = [
    os.path.join(HOME, "Desktop"),
    os.path.join(HOME, "Documents"),
    os.path.join(HOME, "Downloads"),
    os.path.join(HOME, "Pictures"),
    os.path.join(HOME, "Music"),
    os.path.join(HOME, "Videos"),
]

ALIAS_MAP = {
    # common folders
    "desktop": os.path.join(HOME, "Desktop"),
    "documents": os.path.join(HOME, "Documents"),
    "docs": os.path.join(HOME, "Documents"),
    "downloads": os.path.join(HOME, "Downloads"),
    "pictures": os.path.join(HOME, "Pictures"),
    "photos": os.path.join(HOME, "Pictures"),
    "music": os.path.join(HOME, "Music"),
    "videos": os.path.join(HOME, "Videos"),
    # shorthand drive roots
    "c:": "C:/",
    "d:": "D:/",
    "e:": "E:/",
}


def is_probable_path(s: str) -> bool:
    # treat strings containing slashes or a colon as path-like
    return ("/" in s) or ("\\" in s) or re.match(r"^[a-zA-Z]:", s) is not None


def expand_alias(token: str) -> str:
    t = token.strip().strip('"').lower()
    if t in ALIAS_MAP:
        return ALIAS_MAP[t]
    return token


def list_all_fixed_drives() -> List[str]:
    letters = []
    if psutil:
        try:
            for p in psutil.disk_partitions(all=False):
                if p.fstype and re.match(r"^[A-Z]:\\\\$", p.device, re.I):
                    letters.append(p.device[:2])
        except Exception:
            pass
    # fallback: guess C..H
    if not letters:
        letters = [f"{ch}:" for ch in "CDEFGHI"]
    return letters


def _walk_collect_matches(base: str, needle: str, want_dir: Optional[bool]) -> List[str]:
    matches: List[str] = []
    needle_low = needle.lower()
    for root, dirs, files in os.walk(base):
        try:
            if want_dir is not False:
                for d in dirs:
                    if d.lower() == needle_low:
                        matches.append(os.path.join(root, d))
            if want_dir is not True:
                for f in files:
                    if f.lower() == needle_low:
                        matches.append(os.path.join(root, f))
        except PermissionError:
            continue
        if len(matches) >= 200:  # safety cap
            break
    # If no exact matches, try substring contains
    if not matches:
        for root, dirs, files in os.walk(base):
            try:
                if want_dir is not False:
                    for d in dirs:
                        if needle_low in d.lower():
                            matches.append(os.path.join(root, d))
                if want_dir is not True:
                    for f in files:
                        if needle_low in f.lower():
                            matches.append(os.path.join(root, f))
            except PermissionError:
                continue
            if len(matches) >= 200:
                break
    return matches


def find_by_name(name: str, want_dir: Optional[bool] = None, search_dirs: Optional[List[str]] = None,
                 allow_drive_wide: bool = True) -> List[str]:
    """Search common folders first. If not found and allow_drive_wide, ask to broaden to drives."""
    name = name.strip().strip('"')
    name = os.path.basename(name)  # use just the last component if user wrote a path-ish thing

    if search_dirs is None:
        search_dirs = [p for p in DEFAULT_SEARCH_DIRS if os.path.exists(p)] or [HOME]

    matches: List[str] = []
    for base in search_dirs:
        matches.extend(_walk_collect_matches(base, name, want_dir))
    matches = list(dict.fromkeys(matches))  # unique preserve order

    if matches:
        return matches

    if not allow_drive_wide:
        return []

    # Ask user if they want to search across drives
    print(f"I couldn't find '{name}' in common folders.")
    choice = input("Search all drives? (y/n): ").strip().lower()
    if choice != 'y':
        return []

    for letter in list_all_fixed_drives():
        base = f"{letter}/"
        if os.path.exists(base):
            print(f"  scanning {base} … (press Ctrl+C to cancel)")
            try:
                matches.extend(_walk_collect_matches(base, name, want_dir))
            except KeyboardInterrupt:
                print("  cancelled drive scan")
                break
            if len(matches) >= 200:
                break

    return list(dict.fromkeys(matches))


def choose_one(name: str, candidates: List[str]) -> Optional[str]:
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0]

    print(f"Found multiple matches for '{name}':")
    for i, p in enumerate(candidates, 1):
        base = p.replace(HOME, "~")
        print(f"[{i}] {base}")
    while True:
        ans = input("Which one? (enter number or 0 to cancel): ").strip()
        if not ans:
            continue
        if ans == '0':
            return None
        if ans.isdigit():
            idx = int(ans)
            if 1 <= idx <= len(candidates):
                return candidates[idx - 1]
        print("Invalid choice.")


def resolve_any(user_token: str, want_dir: Optional[bool] = None) -> Optional[str]:
    """Resolve a user-supplied token which may be a path, alias, or bare name."""
    token = expand_alias(user_token)
    if is_probable_path(token) and os.path.exists(norm(token)):
        return norm(token)
    # Treat as name search
    cands = find_by_name(token, want_dir=want_dir)
    return choose_one(token, cands)


# ----------------------- System Controller -----------------------

class SystemController:
    """System-level operations: power, files, windows, settings, devices."""

    # ----- Power & Session Control -----
    def shutdown(self):
        say("Shutting down…")
        subprocess.Popen(["shutdown", "/s", "/t", "0"])  # will terminate session

    def restart(self):
        say("Restarting…")
        subprocess.Popen(["shutdown", "/r", "/t", "0"])  # will terminate session

    def sleep(self):
        say("Sleeping…")
        subprocess.call(["rundll32.exe", "powrprof.dll,SetSuspendState", "0,1,0"])  # may require hibernate disabled

    def lock(self):
        say("Locking the workstation…")
        subprocess.call(["rundll32.exe", "user32.dll,LockWorkStation"])  # shell verb

    def logout(self):
        say("Signing out…")
        subprocess.Popen(["shutdown", "/l"])  # log off

    # ----- Storage & Battery -----
    def storage_usage(self, drive: str = "C:"):
        drive = norm(drive)
        total, used, free = shutil.disk_usage(drive)
        say(f"Storage on {drive}: total={total//(2**30)}GB, used={used//(2**30)}GB, free={free//(2**30)}GB")

    def battery_status(self):
        if not psutil:
            say("psutil not installed. Battery info unavailable.")
            return
        batt = psutil.sensors_battery()
        if not batt:
            say("No battery detected.")
        else:
            plugged = "plugged in" if batt.power_plugged else "on battery"
            secs = batt.secsleft
            if secs in (-1, getattr(psutil, 'POWER_TIME_UNLIMITED', -2)):
                left = "unknown"
            else:
                left = f"{secs//3600}h {(secs%3600)//60}m"
            say(f"Battery: {int(batt.percent)}% ({plugged}), time left: {left}")

    # ----- Files & Folders -----
    def create_folder(self, path_or_name: str, in_dir: Optional[str] = None):
        if in_dir:
            base = expand_alias(in_dir)
            base = norm(base)
            if not exists(base):
                # try resolve name for the base folder
                base = resolve_any(in_dir, want_dir=True) or base
        else:
            base = None

        if base:
            target = os.path.join(base, os.path.basename(path_or_name.strip('"')))
        else:
            # If user gave a probable path, use it; otherwise create in current working dir / Documents
            if is_probable_path(path_or_name):
                target = norm(path_or_name)
            else:
                target = os.path.join(ALIAS_MAP.get('documents', os.path.join(HOME, 'Documents')), path_or_name)
        ensure_dir(target)
        say(f"Created/ensured folder: {target}")

    def copy(self, src_token: str, dst_token: str):
        src = resolve_any(src_token)
        if not src:
            say(f"Could not resolve source '{src_token}'")
            return
        dst_token = expand_alias(dst_token)
        dst = norm(dst_token)
        if not is_probable_path(dst_token):
            # it's likely a folder name or alias; resolve as directory
            d = resolve_any(dst_token, want_dir=True)
            if d:
                dst = d
        if os.path.isdir(dst):
            dst = os.path.join(dst, os.path.basename(src))
        ensure_dir(os.path.dirname(dst))
        if os.path.isdir(src):
            shutil.copytree(src, dst, dirs_exist_ok=True)
        else:
            shutil.copy2(src, dst)
        say(f"Copied '{src}' to '{dst}'")

    def move(self, src_token: str, dst_token: str):
        src = resolve_any(src_token)
        if not src:
            say(f"Could not resolve source '{src_token}'")
            return
        dst_token = expand_alias(dst_token)
        dst = norm(dst_token)
        if not is_probable_path(dst_token):
            d = resolve_any(dst_token, want_dir=True)
            if d:
                dst = d
        # If destination is an existing directory, move inside it
        if os.path.isdir(dst):
            dst = os.path.join(dst, os.path.basename(src))
        ensure_dir(os.path.dirname(dst))
        shutil.move(src, dst)
        say(f"Moved '{src}' to '{dst}'")

    def rename(self, src_token: str, new_name: str):
        src = resolve_any(src_token)
        if not src:
            say(f"Could not resolve '{src_token}'")
            return
        dst = norm(os.path.join(os.path.dirname(src), new_name))
        os.rename(src, dst)
        say(f"Renamed '{src}' to '{dst}'")

    def delete(self, token: str):
        target = resolve_any(token)
        if not target:
            say(f"Could not resolve '{token}'")
            return
        if send2trash and os.path.exists(target):
            send2trash(target)
            say(f"Sent to Recycle Bin: {target}")
        else:
            if os.path.isdir(target):
                shutil.rmtree(target)
            else:
                os.remove(target)
            say(f"Deleted permanently: {target}")

    def search(self, directory_token: str, pattern: str) -> List[str]:
        # directory token may be alias or name
        directory_token = expand_alias(directory_token)
        directory = norm(directory_token)
        if not exists(directory):
            d = resolve_any(directory_token, want_dir=True)
            if not d:
                say(f"Could not resolve directory '{directory_token}'")
                return []
            directory = d
        matches = glob.glob(os.path.join(directory, pattern), recursive=True)
        say(f"Found {len(matches)} items matching '{pattern}' in '{directory}'")
        for m in matches[:20]:
            print(" - ", m)
        return matches

    def open_file(self, token: str):
        target = resolve_any(token)
        if not target:
            say(f"Could not resolve '{token}'")
            return
        os.startfile(target)
        say(f"Opened: {target}")

    # ----- Windows / Apps -----
    def open_app(self, name: str):
        name = name.strip().lower()
        mapping = {
            "notepad": "notepad.exe",
            "calculator": "calc.exe",
            "paint": "mspaint.exe",
            "word": "winword.exe",
            "excel": "excel.exe",
            "powerpoint": "powerpnt.exe",
            "vlc": "vlc.exe",
            "media player": "wmplayer.exe",
            "adobe reader": "AcroRd32.exe",
        }
        exe = mapping.get(name, name)
        try:
            subprocess.Popen([exe])
            say(f"Launched app: {exe}")
        except FileNotFoundError:
            say(f"App not found: {exe}. Try full path or ensure app is installed.")

    def close_app_by_name(self, name: str):
        if not psutil:
            say("psutil not installed; cannot enumerate processes.")
            return
        name = name.lower()
        killed = 0
        for p in psutil.process_iter(["name"]):
            try:
                if p.info["name"] and name in p.info["name"].lower():
                    p.terminate()
                    killed += 1
            except Exception:
                pass
        say(f"Closed {killed} process(es) matching '{name}'")

    def alt_tab(self):
        if pyautogui:
            pyautogui.keyDown('alt')
            pyautogui.press('tab')
            time.sleep(0.1)
            pyautogui.keyUp('alt')
            say("Switched app (Alt+Tab)")
        else:
            say("pyautogui not installed; cannot Alt+Tab.")

    # ----- Settings: Volume & Brightness & Radios -----
    def set_volume(self, percent: int):
        if AudioUtilities is None:
            say("PyCAW not installed; cannot set volume.")
            return
        percent = max(0, min(100, int(percent)))
        import comtypes
        comtypes.CoInitialize()
        devices = AudioUtilities.GetSpeakers()
        device = getattr(devices, 'device', devices)
        interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevelScalar(percent/100.0, None)
        say(f"Volume set to {percent}%")

    def volume_up(self, step: int = 5):
        if AudioUtilities is None:
            say("PyCAW not installed; cannot control volume.")
            return
        import comtypes
        comtypes.CoInitialize()
        devices = AudioUtilities.GetSpeakers()
        device = getattr(devices, 'device', devices)
        interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        current = volume.GetMasterVolumeLevelScalar()
        newv = max(0.0, min(1.0, current + step/100.0))
        volume.SetMasterVolumeLevelScalar(newv, None)
        say(f"Volume: {int(newv*100)}%")

    def volume_down(self, step: int = 5):
        self.volume_up(-abs(step))

    def set_brightness(self, percent: int):
        try:
            percent = max(0, min(100, int(percent)))
            cmd = [
                "powershell",
                "(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,{})".format(percent),
            ]
            subprocess.check_call(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            say(f"Brightness set to {percent}%")
        except Exception:
            say("Failed to set brightness (requires WMI support and laptop monitor)")

    import subprocess

def get_wifi_interface():
    result = subprocess.run(
        ["netsh", "interface", "show", "interface"],
        capture_output=True, text=True
    )
    for line in result.stdout.splitlines():
        if "Wireless" in line or "Wi-Fi" in line:
            parts = line.split()
            return parts[-1]  # interface name is last column
    return None

def wifi_toggle(state="off"):
    iface = get_wifi_interface()
    if not iface:
        print("❌ Could not find Wi-Fi interface.")
        return
    cmd = ["netsh", "interface", "set", "interface", iface, "admin=" + state]
    subprocess.run(cmd)
    print(f"✅ Wi-Fi turned {state}")



# ----------------------- Application Controller -----------------------

class ApplicationController:
    """Operate documents, spreadsheets, presentations, images, media, PDFs, and utilities."""

    # ----- Documents -----
    def write_doc(self, text: str, path: str):
        if Document is None:
            say("python-docx not installed.")
            return
        # Path can include aliases like Documents/file.docx
        path = norm(expand_alias(path))
        if not is_probable_path(path):
            # Treat as name; save to Documents by default
            path = os.path.join(ALIAS_MAP.get('documents', os.path.join(HOME, 'Documents')), path)
        ensure_dir(os.path.dirname(path))
        doc = Document()
        doc.add_paragraph(text)
        doc.save(path)
        say(f"Wrote Word document: {path}")

    # ----- Spreadsheets -----
    def write_spreadsheet_sum(self, numbers: List[float], path: str):
        if Workbook is None:
            say("openpyxl not installed.")
            return
        path = norm(expand_alias(path))
        if not is_probable_path(path):
            path = os.path.join(ALIAS_MAP.get('documents', os.path.join(HOME, 'Documents')), path)
        ensure_dir(os.path.dirname(path))
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["Numbers"])
        for n in numbers:
            ws.append([n])
        ws["C1"] = "Sum"
        ws["C2"] = f"=SUM(A2:A{len(numbers)+1})"
        wb.save(path)
        say(f"Saved spreadsheet with SUM formula: {path}")

    # ----- Presentations -----
    def make_presentation(self, title: str, path: str):
        if Presentation is None:
            say("python-pptx not installed.")
            return
        path = norm(expand_alias(path))
        if not is_probable_path(path):
            path = os.path.join(ALIAS_MAP.get('documents', os.path.join(HOME, 'Documents')), path)
        ensure_dir(os.path.dirname(path))
        prs = Presentation()
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = title
        slide.placeholders[1].text = "Generated by Local AI"
        prs.save(path)
        say(f"Created presentation: {path}")

    # ----- Images -----
    def image_resize(self, src_token: str, dst_token: str, width: int, height: int):
        if Image is None:
            say("Pillow not installed.")
            return
        src = resolve_any(src_token)
        if not src:
            say(f"Could not resolve '{src_token}'")
            return
        dst_token = expand_alias(dst_token)
        dst = norm(dst_token)
        if not is_probable_path(dst_token):
            # Save next to source
            base, ext = os.path.splitext(src)
            dst = f"{base}_resized{ext}"
        ensure_dir(os.path.dirname(dst))
        with Image.open(src) as im:
            im2 = im.resize((width, height))
            im2.save(dst)
        say(f"Saved resized image to: {dst}")

    def image_crop(self, src_token: str, dst_token: str, left: int, top: int, right: int, bottom: int):
        if Image is None:
            say("Pillow not installed.")
            return
        src = resolve_any(src_token)
        if not src:
            say(f"Could not resolve '{src_token}'")
            return
        dst_token = expand_alias(dst_token)
        dst = norm(dst_token)
        if not is_probable_path(dst_token):
            base, ext = os.path.splitext(src)
            dst = f"{base}_crop{ext}"
        ensure_dir(os.path.dirname(dst))
        with Image.open(src) as im:
            im2 = im.crop((left, top, right, bottom))
            im2.save(dst)
        say(f"Saved cropped image to: {dst}")

    # ----- Media Playback -----
    def play_media(self, token: str):
        target = resolve_any(token)
        if not target:
            say(f"Could not resolve '{token}'")
            return
        os.startfile(target)
        say(f"Playing media: {target}")

    # ----- PDFs -----
    def open_pdf(self, token: str):
        target = resolve_any(token)
        if not target:
            say(f"Could not resolve '{token}'")
            return
        os.startfile(target)
        say(f"Opened PDF: {target}")

    def search_pdf(self, token: str, query: str):
        if PyPDF2 is None:
            say("PyPDF2 not installed.")
            return
        path = resolve_any(token)
        if not path:
            say(f"Could not resolve '{token}'")
            return
        hits = []
        with open(path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            for i, page in enumerate(reader.pages):
                try:
                    text = page.extract_text() or ""
                except Exception:
                    text = ""
                if query.lower() in text.lower():
                    hits.append(i+1)
        if hits:
            say(f"Query '{query}' found on pages: {hits}")
        else:
            say(f"Query '{query}' not found.")

    # ----- Utilities -----
    def open_notepad(self):
        subprocess.Popen(["notepad.exe"])
        say("Opened Notepad")

    def open_calculator(self):
        subprocess.Popen(["calc.exe"])
        say("Opened Calculator")

    def antivirus_quick_scan(self):
        defender = r"C:\\Program Files\\Windows Defender\\MpCmdRun.exe"
        if os.path.exists(defender):
            subprocess.Popen([defender, "-Scan", "-ScanType", "1"])  # Quick scan
            say("Started Windows Defender quick scan")
        else:
            say("Windows Defender CLI not found.")


# ----------------------- Intent Parser -----------------------

@dataclass
class Intent:
    name: str
    args: Tuple


class IntentParser:
    """Lightweight, rule-based parser for common phrasings (paths OR names)."""
    def __init__(self, sysc: SystemController, appc: ApplicationController):
        self.sys = sysc
        self.app = appc

    def parse(self, text: str) -> Optional[Intent]:
        t_raw = text.strip()
        t = t_raw.lower()

        # Quit & help
        if t in {"quit", "exit", "bye"}:
            return Intent("quit", tuple())
        if t in {"help", "?"}:
            return Intent("help", tuple())

        # Power
        if re.search(r"\bshutdown\b", t):
            return Intent("shutdown", tuple())
        if re.search(r"\brestart\b", t):
            return Intent("restart", tuple())
        if re.search(r"\bsleep\b", t):
            return Intent("sleep", tuple())
        if re.search(r"\block\b", t):
            return Intent("lock", tuple())
        if re.search(r"\blog(out| off)\b", t):
            return Intent("logout", tuple())

        # Storage / battery
        m = re.search(r"storage\s+([a-z]:)", t)
        if m:
            return Intent("storage", (m.group(1),))
        if t.startswith("battery"):
            return Intent("battery", tuple())

        # Volume/brightness
        m = re.search(r"set\s+volume\s+(\d{1,3})", t)
        if m:
            return Intent("set_volume", (int(m.group(1)),))
        if "volume up" in t:
            return Intent("volume_up", tuple())
        if "volume down" in t:
            return Intent("volume_down", tuple())
        m = re.search(r"set\s+brightness\s+(\d{1,3})", t)
        if m:
            return Intent("set_brightness", (int(m.group(1)),))
        if "brightness up" in t:
            return Intent("brightness_up", tuple())
        if "brightness down" in t:
            return Intent("brightness_down", tuple())
        if "wifi on" in t:
            return Intent("wifi_on", tuple())
        if "wifi off" in t:
            return Intent("wifi_off", tuple())

        # ----- Files & Folders (paths or names, quotes optional) -----
        # create folder "Work" in Documents
        m = re.search(r"create\s+folder\s+\"?(.+?)\"?\s+in\s+\"?(.+?)\"?$", t_raw, re.IGNORECASE)
        if m:
            return Intent("create_folder", (m.group(1), m.group(2)))

        # copy/move/rename/delete/open/search
        for intent_name, pat in [
            ("copy", r"copy\s+\"?(.+?)\"?\s+to\s+\"?(.+?)\"?$"),
            ("move", r"move\s+\"?(.+?)\"?\s+to\s+\"?(.+?)\"?$"),
            ("rename", r"rename\s+\"?(.+?)\"?\s+to\s+\"?(.+?)\"?$"),
            ("delete", r"delete\s+\"?(.+?)\"?$"),
            ("open_file", r"open\s+file\s+\"?(.+?)\"?$"),
            ("open_file", r"open\s+\"?(.+?\..+?)\"?$"),  # open notes.txt
            ("search", r"search\s+\"?(.+?)\"?\s+in\s+\"?(.+?)\"?$"),
        ]:
            m = re.search(pat, t_raw, re.IGNORECASE)
            if m:
                if intent_name in {"copy", "move"}:
                    return Intent(intent_name, (m.group(1), m.group(2)))
                if intent_name == "rename":
                    return Intent("rename", (m.group(1), os.path.basename(m.group(2))))
                if intent_name == "delete":
                    return Intent("delete", (m.group(1),))
                if intent_name == "open_file":
                    return Intent("open_file", (m.group(1),))
                if intent_name == "search":
                    # pattern then directory
                    return Intent("search", (m.group(2), m.group(1)))

        # Windows/apps
        m = re.search(r"open\s+([a-z\s]+)$", t)
        if m:
            return Intent("open_app", (m.group(1).strip(),))
        m = re.search(r"close\s+([a-z\s]+)$", t)
        if m:
            return Intent("close_app", (m.group(1).strip(),))
        if t in {"alt tab", "switch app", "switch"}:
            return Intent("alt_tab", tuple())

        # Documents/spreadsheets/presentations
        m = re.search(r"write\s+doc\s+\"(.+?)\"\s+to\s+\"?(.+?)\"?$", t_raw, re.IGNORECASE)
        if m:
            return Intent("write_doc", (m.group(1), m.group(2)))
        m = re.search(r"spreadsheet\s+sum\s+([0-9,\.\s-]+)\s+to\s+\"?(.+?)\"?$", t_raw, re.IGNORECASE)
        if m:
            nums = [float(x) for x in re.split(r"[\s,]+", m.group(1).strip()) if x]
            return Intent("sheet_sum", (nums, m.group(2)))
        m = re.search(r"presentation\s+title\s+\"(.+?)\"\s+to\s+\"?(.+?)\"?$", t_raw, re.IGNORECASE)
        if m:
            return Intent("make_ppt", (m.group(1), m.group(2)))

        # Images
        m = re.search(r"resize\s+image\s+\"?(.+?)\"?\s+to\s+(\d+)x(\d+)\s+save\s+\"?(.+?)\"?$", t_raw, re.IGNORECASE)
        if m:
            return Intent("img_resize", (m.group(1), m.group(4), int(m.group(2)), int(m.group(3))))
        m = re.search(r"crop\s+image\s+\"?(.+?)\"?\s+\((\d+),(\d+),(\d+),(\d+)\)\s+save\s+\"?(.+?)\"?$", t_raw, re.IGNORECASE)
        if m:
            return Intent("img_crop", (m.group(1), m.group(6), int(m.group(2)), int(m.group(3)), int(m.group(4)), int(m.group(5))))

        # Media & PDFs
        m = re.search(r"play\s+(video|audio|media)\s+\"?(.+?)\"?$", t_raw, re.IGNORECASE)
        if m:
            return Intent("play_media", (m.group(2),))
        m = re.search(r"read\s+pdf\s+\"?(.+?)\"?$", t_raw, re.IGNORECASE)
        if m:
            return Intent("open_pdf", (m.group(1),))
        m = re.search(r"search\s+pdf\s+\"(.+?)\"\s+in\s+\"?(.+?)\"?$", t_raw, re.IGNORECASE)
        if m:
            return Intent("search_pdf", (m.group(2), m.group(1)))

        # Utilities
        if t == "notepad":
            return Intent("open_notepad", tuple())
        if t == "calculator":
            return Intent("open_calculator", tuple())
        if "antivirus scan" in t or "defender scan" in t:
            return Intent("antivirus_scan", tuple())

        # Conversational
        if re.search(r"\b(hello|hi|hey|greetings)\b", t):
            return Intent("greet", tuple())
        if re.search(r"\b(how are you|how's it going|how do you do)\b", t):
            return Intent("status", tuple())
        if re.search(r"\b(what's your name|who are you|what are you)\b", t):
            return Intent("introduce", tuple())
        if re.search(r"\b(thank you|thanks|thankyou)\b", t):
            return Intent("thank", tuple())
        if re.search(r"\b(goodbye|bye|see you later|farewell)\b", t):
            return Intent("goodbye", tuple())

        # Science questions
        if "?" in t:
            return Intent("science_question", (t_raw,))

        # Math calculations
        if re.search(r"\b(calculate|what is|compute|math)\b.*[\d\+\-\*/]", t):
            return Intent("math", (t_raw,))

        # Time
        if re.search(r"\b(time|what time)\b", t):
            return Intent("time", tuple())

        # Date
        if re.search(r"\b(date|what date|today)\b", t):
            return Intent("date", tuple())

        # Joke
        if re.search(r"\b(joke|tell.*joke)\b", t):
            return Intent("joke", tuple())

        # Quote
        if re.search(r"\b(quote|inspire|motivate)\b", t):
            return Intent("quote", tuple())

        return None

    # ------ Execute intents ------
    def dispatch(self, intent: Intent):
        name, args = intent.name, intent.args
        s, a = self.sys, self.app
        if name == "quit":
            say("Goodbye!")
            sys.exit(0)
        if name == "help":
            print_help()
            return
        # Power
        if name == "shutdown": s.shutdown(); return
        if name == "restart": s.restart(); return
        if name == "sleep": s.sleep(); return
        if name == "lock": s.lock(); return
        if name == "logout": s.logout(); return
        # Storage/Battery
        if name == "storage": s.storage_usage(*args); return
        if name == "battery": s.battery_status(); return
        # Settings
        if name == "set_volume": s.set_volume(*args); return
        if name == "volume_up": s.volume_up(); return
        if name == "volume_down": s.volume_down(); return
        if name == "set_brightness": s.set_brightness(*args); return
        if name == "brightness_up": s.set_brightness( min(100, 10 + current_brightness_guess()) ); return
        if name == "brightness_down": s.set_brightness( max(0, current_brightness_guess() - 10) ); return
        if name == "wifi_on": s.wifi(True); return
        if name == "wifi_off": s.wifi(False); return
        # Files
        if name == "create_folder": s.create_folder(*args); return
        if name == "copy": s.copy(*args); return
        if name == "move": s.move(*args); return
        if name == "rename": s.rename(*args); return
        if name == "delete": s.delete(*args); return
        if name == "search": s.search(*args); return
        if name == "open_file": s.open_file(*args); return
        # Windows/apps
        if name == "open_app": s.open_app(*args); return
        if name == "close_app": s.close_app_by_name(*args); return
        if name == "alt_tab": s.alt_tab(); return
        # Docs/Sheets/PPT
        if name == "write_doc": a.write_doc(*args); return
        if name == "sheet_sum": a.write_spreadsheet_sum(*args); return
        if name == "make_ppt": a.make_presentation(*args); return
        # Images
        if name == "img_resize": a.image_resize(*args); return
        if name == "img_crop": a.image_crop(*args); return
        # Media & PDFs
        if name == "play_media": a.play_media(*args); return
        if name == "open_pdf": a.open_pdf(*args); return
        if name == "search_pdf": a.search_pdf(*args); return
        # Utilities
        if name == "open_notepad": a.open_notepad(); return
        if name == "open_calculator": a.open_calculator(); return
        if name == "antivirus_scan": a.antivirus_quick_scan(); return

        # Conversational
        if name == "greet":
            say("Hello! Nice to meet you. What can I do for you?")
            return
        if name == "status":
            say("I'm doing great, thank you! How about you?")
            return
        if name == "introduce":
            say("I'm your local AI assistant, designed to help with system tasks and answer science questions.")
            return
        if name == "thank":
            say("You're welcome! Is there anything else I can help with?")
            return
        if name == "goodbye":
            say("Goodbye! Have a wonderful day.")
            sys.exit(0)

        # Science
        if name == "science_question":
            answer_science(args[0])
            return

        # Math
        if name == "math":
            calculate_math(args[0])
            return

        # Time
        if name == "time":
            get_time()
            return

        # Date
        if name == "date":
            get_date()
            return

        # Joke
        if name == "joke":
            tell_joke()
            return

        # Quote
        if name == "quote":
            tell_quote()
            return

        say("Sorry, I didn't understand that. Try 'help' for commands or ask a science question.")


def current_brightness_guess() -> int:
    # We don't read brightness; return a static guess the user can nudge.
    return 50


# ----------------------- UI Loop -----------------------

def print_help():
    print(
        """
Commands you can use (no full paths needed):
  create folder Work in Documents
  copy A.txt to Desktop
  move budget.xlsx to Documents
  rename notes.txt to notes-old.txt
  delete temp.log
  search "*.pdf" in Documents
  open file report.docx
  open notes.txt

  open notepad | open calculator | open paint | open word | open excel | open powerpoint
  close notepad
  alt tab

  write doc "Hello world" to Documents/hello.docx
  spreadsheet sum 1,2,3,4 to Documents/sum.xlsx
  presentation title "My Deck" to Desktop/deck.pptx

  resize image "photo.jpg" to 800x600 save "photo_out.jpg"
  crop image "photo.jpg" (0,0,400,300) save "photo_crop.jpg"

  play media movie.mp4
  read pdf manual.pdf
  search pdf "budget" in report.pdf

  set volume 30 | volume up | volume down
  set brightness 60 | brightness up | brightness down
  wifi on | wifi off

  storage C: | battery
  shutdown | restart | sleep | lock | logout

  help | quit
"""
    )


def main():
    say("Hello! I'm your AI assistant. How can I help?")
    sysc = SystemController()
    appc = ApplicationController()
    parser = IntentParser(sysc, appc)

    # Optional: accept a one-shot command via CLI args
    if len(sys.argv) > 1:
        cmd = " ".join(sys.argv[1:])
        intent = parser.parse(cmd)
        if intent:
            parser.dispatch(intent)
        else:
            say("Sorry, I didn't understand that. Try 'help' for commands.")
        return

    while True:
        try:
            text = input("You: ")
        except (EOFError, KeyboardInterrupt):
            print()
            break
        if not text.strip():
            continue
        intent = parser.parse(text)
        if intent:
            try:
                parser.dispatch(intent)
            except Exception as e:
                say(f"Error: {e}")
        else:
            say("Sorry, I didn't understand. Type 'help' for examples.")


if __name__ == "__main__":
    main()
