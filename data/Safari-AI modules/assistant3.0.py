

# Local Windows AI Assistant (Offline) — Enhanced with Voice Recognition
# =============================================================================

"""
Local Windows AI Assistant (Offline) — Enhanced with Voice Recognition
=====================================================================

This single-file Python app runs locally on Windows and performs two categories of tasks:
1) Navigating Around (system-level operations)
2) Using Software Systems (operate common apps & files)

What's new in this version
--------------------------
- **Voice Recognition**: Talk to your AI assistant hands-free
- **Wake Word Activation**: Say "Hey Safari" to activate listening mode
- **Continuous Listening**: The AI can listen for commands continuously
- **Voice Feedback Toggle**: Choose between voice and text responses
- **Visual Indicators**: See when the AI is listening, processing, or speaking
- **Fallback to Text**: If voice recognition fails, you can still type commands

Voice Commands:
- "Hey Safari" - Wake up the AI
- "Stop listening" - Deactivate continuous listening
- "Type mode" - Switch to text input mode
- "Voice mode" - Switch back to voice input mode

Install (recommended, then run):
    py -m pip install --upgrade pip
    py -m pip install pyautogui pywinauto psutil send2trash pyttsx3 python-docx openpyxl python-pptx pillow PyPDF2 comtypes pycaw pywin32 SpeechRecognition pyaudio

Run:
    py local_ai.py

Note: For voice recognition, you may need to install additional audio drivers:
    - On Windows: PyAudio should work with most systems
    - If you encounter issues, try: pip install pipwin && pipwin install pyaudio

Examples you can say (no full paths needed!):
    "Hey Safari, create folder Work in Documents"
    "Hey Safari, what time is it?"
    "Hey Safari, set volume to 30"
    "Hey Safari, tell me a joke"
    "Hey Safari, open notepad"
"""

from __future__ import annotations
import os
import re
import sys
import time
import shutil
import glob
import subprocess
import datetime
import random
import threading
import queue
from dataclasses import dataclass
from typing import Optional, List, Tuple, Dict, Any


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

# Voice recognition imports
try:
    import speech_recognition as sr
except Exception:
    sr = None

try:
    import pyaudio
except Exception:
    pyaudio = None

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
    import openpyxl
except Exception:
    openpyxl = None

try:
    from pptx import Presentation
except Exception:
    Presentation = None

try:
    from PIL import Image
except Exception:
    Image = None


# ----------------------- Voice Recognition Module -----------------------

class VoiceRecognition:
    """Handles voice recognition and audio input/output."""
    
    def __init__(self, conversation_manager):
        self.conv_manager = conversation_manager
        self.recognizer = sr.Recognizer() if sr else None
        self.microphone = sr.Microphone() if (sr and pyaudio) else None
        self.listening = False
        self.continuous_listening = False
        self.wake_word = "hey assistant"
        self.voice_enabled = True
        self.audio_queue = queue.Queue()
        self.listening_thread = None
        
        # Adjust recognizer settings
        if self.recognizer:
            self.recognizer.energy_threshold = 300
            self.recognizer.dynamic_energy_threshold = True
            self.recognizer.pause_threshold = 0.8
    
    def is_available(self) -> bool:
        """Check if voice recognition is available."""
        return self.recognizer is not None and self.microphone is not None
    
    def start_continuous_listening(self):
        """Start continuous listening in a separate thread."""
        if not self.is_available():
            say("Voice recognition is not available. Please install SpeechRecognition and PyAudio.")
            return False
        
        if self.continuous_listening:
            return True
        
        self.continuous_listening = True
        self.listening_thread = threading.Thread(target=self._continuous_listen_loop, daemon=True)
        self.listening_thread.start()
        say("Voice recognition activated. Say 'Safari' followed by your command.")
        return True
    
    def stop_continuous_listening(self):
        """Stop continuous listening."""
        self.continuous_listening = False
        if self.listening_thread:
            self.listening_thread.join(timeout=1)
        say("Voice recognition deactivated.")
    
    def _continuous_listen_loop(self):
        """Background thread for continuous listening."""
        with self.microphone as source:
            self.recognizer.adjust_for_ambient_noise(source, duration=1)
        
        while self.continuous_listening:
            try:
                # Listen for wake word
                if self._listen_for_wake_word():
                    say("I'm listening...")
                    # Listen for command after wake word
                    command = self._listen_for_command()
                    if command:
                        self.audio_queue.put(command)
            except Exception as e:
                print(f"Voice recognition error: {e}")
                time.sleep(1)
    
    def _listen_for_wake_word(self) -> bool:
        """Listen for the wake word."""
        with self.microphone as source:
            try:
                audio = self.recognizer.listen(source, timeout=1, phrase_time_limit=3)
                text = self.recognizer.recognize_google(audio).lower()
                return self.wake_word in text
            except sr.WaitTimeoutError:
                return False
            except sr.UnknownValueError:
                return False
            except Exception as e:
                print(f"Wake word detection error: {e}")
                return False
    
    def _listen_for_command(self) -> Optional[str]:
        """Listen for a command after wake word."""
        with self.microphone as source:
            try:
                print("Listening for command...")
                audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=10)
                command = self.recognizer.recognize_google(audio)
                print(f"Recognized: {command}")
                return command
            except sr.WaitTimeoutError:
                say("I didn't hear anything. Please try again.")
                return None
            except sr.UnknownValueError:
                say("Sorry, I didn't understand that. Please try again.")
                return None
            except Exception as e:
                print(f"Command recognition error: {e}")
                say("Sorry, I had trouble understanding. Please try again.")
                return None
    
    def listen_once(self) -> Optional[str]:
        """Listen for a single command (non-continuous mode)."""
        if not self.is_available():
            say("Voice recognition is not available.")
            return None
        
        with self.microphone as source:
            self.recognizer.adjust_for_ambient_noise(source, duration=1)
            say("I'm listening...")
            
            try:
                audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=10)
                command = self.recognizer.recognize_google(audio)
                print(f"Recognized: {command}")
                return command
            except sr.WaitTimeoutError:
                say("I didn't hear anything. Please try again.")
                return None
            except sr.UnknownValueError:
                say("Sorry, I didn't understand that. Please try again.")
                return None
            except Exception as e:
                print(f"Voice recognition error: {e}")
                say("Sorry, I had trouble understanding. Please try again.")
                return None
    
    def get_audio_command(self) -> Optional[str]:
        """Get command from audio queue (for continuous mode)."""
        try:
            return self.audio_queue.get_nowait()
        except queue.Empty:
            return None
    
    def toggle_voice_mode(self):
        """Toggle between voice and text input modes."""
        self.voice_enabled = not self.voice_enabled
        if self.voice_enabled:
            say("Voice mode enabled.")
        else:
            say("Text mode enabled. Voice recognition is off.")
            self.stop_continuous_listening()
        return self.voice_enabled

# ----------------------- Helpers -----------------------

def say(text: str, voice_override: Optional[bool] = None):
    """Speak + print feedback (if pyttsx3 is available)."""
    print(text)
    
    # Check if we should use voice
    use_voice = voice_override if voice_override is not None else (
        hasattr(say, 'voice_enabled') and say.voice_enabled
    )
    
    if use_voice and pyttsx3:
        try:
            engine = pyttsx3.init()
            engine.setProperty('rate', 180)  # Normal speaking speed
            engine.setProperty('volume', 0.8)  # Slightly lower volume for smoothness
            voices = engine.getProperty('voices')
            if voices and len(voices) > 1:
                engine.setProperty('voice', voices[1].id)  # Use second voice if available
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
    say(random.choice(jokes))


def tell_quote():
    quotes = [
        "The only way to do great work is to love what you do. - Steve Jobs",
        "Believe you can and you're halfway there. - Theodore Roosevelt",
        "The future belongs to those who believe in the beauty of their dreams. - Eleanor Roosevelt",
        "You miss 100% of the shots you don't take. - Wayne Gretzky",
        "The best way to predict the future is to create it. - Peter Drucker",
    ]
    say(random.choice(quotes))


def get_time():
    say(f"The current time is {datetime.datetime.now().strftime('%I:%M %p')}")


def get_date():
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


# ----------------------- Conversation Manager -----------------------

class ConversationManager:
    """Manages conversation context and history for more natural interactions."""
    
    def __init__(self, max_history=20):
        self.history: List[Dict[str, Any]] = []
        self.max_history = max_history
        self.context: Dict[str, Any] = {}
        self.user_profile: Dict[str, Any] = {
            "name": None,
            "preferences": {}
        }
    
    def add_to_history(self, role: str, message: str):
        """Add a message to the conversation history."""
        self.history.append({
            "role": role, 
            "message": message, 
            "timestamp": time.time()
        })
        # Keep only the most recent messages
        if len(self.history) > self.max_history:
            self.history = self.history[-self.max_history:]
    
    def get_recent_context(self, count: int = 5) -> List[Dict[str, Any]]:
        """Get the most recent conversation context."""
        return self.history[-count:] if self.history else []
    
    def update_context(self, key: str, value: Any):
        """Update the conversation context."""
        self.context[key] = value
    
    def get_context(self, key: str) -> Any:
        """Get a value from the conversation context."""
        return self.context.get(key)
    
    def clear_context(self):
        """Clear the conversation context."""
        self.context = {}
    
    def set_user_preference(self, key: str, value: Any):
        """Set a user preference."""
        self.user_profile["preferences"][key] = value
    
    def get_user_preference(self, key: str, default: Any = None) -> Any:
        """Get a user preference."""
        return self.user_profile["preferences"].get(key, default)


# ----------------------- Response Generator -----------------------

class ResponseGenerator:
    """Generates more natural, conversational responses."""
    
    def __init__(self, conversation_manager: ConversationManager):
        self.conv_manager = conversation_manager
        self.response_templates = {
            "greeting": [
                "Hello! How can I help you today?",
                "Hi there! What can I do for you?",
                "Good to see you! What would you like to work on?",
            ],
            "farewell": [
                "Goodbye! Have a great day!",
                "See you later! Feel free to come back anytime.",
                "Until next time! Take care.",
            ],
            "confirmation": [
                "Got it. I'll {action} right away.",
                "Sure thing. {action} now.",
                "I'll {action} for you.",
            ],
            "error": [
                "I'm sorry, I couldn't {action}. Could you try again?",
                "Hmm, I had trouble {action}. Maybe we could try a different approach?",
                "I wasn't able to {action}. Is there another way I can help?",
            ],
            "unknown": [
                "I'm not sure how to help with that. Could you rephrase?",
                "I don't understand. Can you tell me more?",
                "That's beyond my current capabilities. Is there something else I can do?",
            ]
        }
    
    def generate_response(self, intent: 'Intent', success: bool = True, details: Optional[str] = None) -> str:
        """Generate a contextual response based on intent and execution result."""
        if intent.name == "greet":
            return self._get_random_template("greeting")
        
        if intent.name == "goodbye":
            return self._get_random_template("farewell")
        
        if success:
            if details:
                return f"I've successfully {self._intent_to_phrase(intent)}: {details}"
            else:
                return self._get_random_template("confirmation").format(
                    action=self._intent_to_phrase(intent)
                )
        else:
            return self._get_random_template("error").format(
                action=self._intent_to_phrase(intent)
            )
    
    def _get_random_template(self, category: str) -> str:
        """Get a random response template from a category."""
        return random.choice(self.response_templates.get(category, ["I'm not sure how to respond."]))
    
    def _intent_to_phrase(self, intent: 'Intent') -> str:
        """Convert an intent to a natural language phrase."""
        phrases = {
            "create_folder": "create the folder",
            "copy": "copy the file",
            "move": "move the file",
            "rename": "rename the file",
            "delete": "delete the file",
            "open_file": "open the file",
            "search": "search for files",
            "open_app": "open the application",
            "close_app": "close the application",
            "set_volume": "set the volume",
            "set_brightness": "adjust the brightness",
            "shutdown": "shutdown your computer",
            "restart": "restart your computer",
            "sleep": "put your computer to sleep",
            "lock": "lock your computer",
            "logout": "log you out",
            "write_doc": "create the document",
            "sheet_sum": "create the spreadsheet",
            "make_ppt": "create the presentation",
            "img_resize": "resize the image",
            "img_crop": "crop the image",
            "play_media": "play the media file",
            "open_pdf": "open the PDF",
            "search_pdf": "search the PDF",
            "antivirus_scan": "start the antivirus scan",
            "wifi_on": "turn on Wi-Fi",
            "wifi_off": "turn off Wi-Fi",
        }
        return phrases.get(intent.name, "perform that action")
    
    def generate_conversational_response(self, user_input: str) -> str:
        """Generate a response for purely conversational inputs."""
        user_input_lower = user_input.lower()
        
        # Check for follow-up questions
        if any(word in user_input_lower for word in ["why", "how", "what", "when", "where"]):
            return "That's an interesting question. I'd need more context to give you a good answer."
        
        # Check for expressions of emotion
        if any(word in user_input_lower for word in ["happy", "glad", "excited"]):
            return "I'm glad to hear that! Is there anything I can help you with today?"
        
        if any(word in user_input_lower for word in ["sad", "upset", "frustrated"]):
            return "I'm sorry to hear that. Maybe I can help with something to make your day better?"
        
        # Default conversational response
        return "I see. Is there something specific you'd like me to help you with?"


# ----------------------- Conversational Enhancements -----------------------

class ConversationalEnhancements:
    """Additional conversational capabilities for the AI assistant."""
    
    def __init__(self, conversation_manager: ConversationManager):
        self.conv_manager = conversation_manager
        self.personality_traits = {
            "helpfulness": 0.9,
            "friendliness": 0.8,
            "humor": 0.6,
            "formality": 0.5  # 0 = very casual, 1 = very formal
        }
    
    def adjust_response_tone(self, response: str) -> str:
        """Adjust the tone of responses based on user preferences and conversation context."""
        # Check if user prefers more formal or casual language
        formality = self.conv_manager.get_user_preference("formality", self.personality_traits["formality"])
        
        if formality < 0.3:  # Very casual
            response = response.replace("I am", "I'm")
            response = response.replace("I will", "I'll")
            response = response.replace("I have", "I've")
            response = response.replace("you are", "you're")
            response = response.replace("you will", "you'll")
            response = response.replace("do not", "don't")
        
        elif formality > 0.7:  # More formal
            response = response.replace("I'm", "I am")
            response = response.replace("I'll", "I will")
            response = response.replace("I've", "I have")
            response = response.replace("you're", "you are")
            response = response.replace("you'll", "you will")
            response = response.replace("don't", "do not")
        
        return response
    
    def detect_user_mood(self, user_input: str) -> str:
        """Simple mood detection based on user input."""
        user_input_lower = user_input.lower()
        
        # Positive indicators
        positive_words = ["happy", "glad", "great", "awesome", "fantastic", "love", "like", "enjoy"]
        if any(word in user_input_lower for word in positive_words):
            return "positive"
        
        # Negative indicators
        negative_words = ["sad", "angry", "frustrated", "hate", "dislike", "terrible", "awful", "bad"]
        if any(word in user_input_lower for word in negative_words):
            return "negative"
        
        # Neutral
        return "neutral"
    
    def respond_to_mood(self, mood: str) -> Optional[str]:
        """Generate a response based on detected user mood."""
        if mood == "positive":
            return "I'm glad to hear that! Is there anything I can help you with today?"
        elif mood == "negative":
            return "I'm sorry to hear that. Is there something I can do to help improve your day?"
        else:
            return None  # No special response for neutral mood


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

    def get_current_brightness(self) -> int:
        """Get current brightness level."""
        try:
            cmd = [
                "powershell",
                "(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness).CurrentBrightness"
            ]
            result = subprocess.run(cmd, capture_output=True, text=True)
            return int(result.stdout.strip())
        except Exception:
            return 50  # fallback

    def brightness_up(self, step: int = 10):
        current = self.get_current_brightness()
        new = min(100, current + step)
        self.set_brightness(new)

    def brightness_down(self, step: int = 10):
        current = self.get_current_brightness()
        new = max(0, current - step)
        self.set_brightness(new)

    def wifi_toggle(self, state: str):
        """Toggle Wi-Fi on/off."""
        try:
            # Get Wi-Fi interface name
            result = subprocess.run(
                ["netsh", "interface", "show", "interface"],
                capture_output=True, text=True
            )
            iface = None
            for line in result.stdout.splitlines():
                if "Wireless" in line or "Wi-Fi" in line:
                    parts = line.split()
                    iface = parts[-1]  # interface name is last column
                    break
            
            if not iface:
                say("Could not find Wi-Fi interface.")
                return
            
            cmd = ["netsh", "interface", "set", "interface", iface, f"admin={state}"]
            subprocess.run(cmd)
            say(f"Wi-Fi turned {state}")
        except Exception as e:
            say(f"Failed to toggle Wi-Fi: {str(e)}")

    def wifi_on(self):
        self.wifi_toggle("enabled")

    def wifi_off(self):
        self.wifi_toggle("disabled")


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
        if openpyxl is None:
            say("openpyxl not installed.")
            return
        path = norm(expand_alias(path))
        if not is_probable_path(path):
            path = os.path.join(ALIAS_MAP.get('documents', os.path.join(HOME, 'Documents')), path)
        ensure_dir(os.path.dirname(path))
        wb = openpyxl.Workbook()
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
try:
    import PyPDF2
except ImportError:
    PyPDF2 = None    
class PDFTools:
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
        with open(path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            for i, page in enumerate(reader.pages):
                try:
                    text = page.extract_text() or ""
                except Exception:
                    text = ""
                if query.lower() in text.lower():
                    hits.append(i + 1)

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


class EnhancedIntentParser:
    """Enhanced intent parser with better conversational understanding."""
    
    def __init__(self, sysc: SystemController, appc: ApplicationController, conversation_manager: ConversationManager):
        self.sys = sysc
        self.app = appc
        self.conv_manager = conversation_manager

    def parse(self, text: str) -> Optional[Intent]:
        t_raw = text.strip()
        t = t_raw.lower()

        # First check if this is a follow-up to a previous command
        recent_context = self.conv_manager.get_recent_context(3)
        
        # Check for simple confirmations or denials
        if t in ["yes", "yeah", "sure", "ok", "do it", "please do"]:
            # Check if there's a pending action in context
            pending_action = self.conv_manager.get_context("pending_action")
            if pending_action:
                return pending_action
        
        if t in ["no", "nope", "don't", "stop", "cancel"]:
            # Clear any pending action
            self.conv_manager.update_context("pending_action", None)
            return Intent("cancel", tuple())
        
        # Check for conversational patterns
        if self._is_conversational(t_raw):
            return Intent("conversational", (t_raw,))
        
        # Check for personal questions to the AI
        if self._is_personal_question(t_raw):
            return Intent("personal_question", (t_raw,))
        
        # Voice control commands
        if t in ["stop listening", "stop voice", "voice off"]:
            return Intent("stop_voice", tuple())
        if t in ["start listening", "voice on", "hey assistant"]:
            return Intent("start_voice", tuple())
        if t in ["type mode", "text mode", "keyboard mode"]:
            return Intent("type_mode", tuple())
        if t in ["voice mode", "talk mode"]:
            return Intent("voice_mode", tuple())
        
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

    def _is_conversational(self, text: str) -> bool:
        """Check if the input is purely conversational."""
        conversational_patterns = [
            r"^(how|what|why|when|where|who) (are|is|do|did|does|can|will|would|could|should)",
            r"^(i think|i feel|i believe|i want|i need|i wish)",
            r"^(tell me about|do you know|can you explain)",
            r"^(that's|that is|it's|it is) (good|bad|great|terrible|awesome|horrible)",
        ]
        
        text_lower = text.lower()
        return any(re.search(pattern, text_lower) for pattern in conversational_patterns)
    
    def _is_personal_question(self, text: str) -> bool:
        """Check if the input is a personal question to the AI."""
        personal_patterns = [
            r"(what|who) (are|is) you",
            r"(what|who) (do|does) you (do|be)",
            r"(where|when) (were|was) you (created|made|born)",
            r"(do|can) you (have|feel|think|believe)",
            r"(are|is) you (human|alive|real)",
            r"(do|can) you (learn|remember|know)",
        ]
        
        text_lower = text.lower()
        return any(re.search(pattern, text_lower) for pattern in personal_patterns)

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
        if name == "brightness_up": s.brightness_up(); return
        if name == "brightness_down": s.brightness_down(); return
        if name == "wifi_on": s.wifi_on(); return
        if name == "wifi_off": s.wifi_off(); return
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


# ----------------------- Chat Interface -----------------------

class ChatInterface:
    """A more interactive, chat-like interface for the AI assistant with voice support."""
    
    def __init__(self, sysc: SystemController, appc: ApplicationController):
        self.conv_manager = ConversationManager()
        self.response_generator = ResponseGenerator(self.conv_manager)
        self.parser = EnhancedIntentParser(sysc, appc, self.conv_manager)
        self.sysc = sysc
        self.appc = appc
        self.conv_enhancements = ConversationalEnhancements(self.conv_manager)
        self.voice_recognition = VoiceRecognition(self.conv_manager)
        
        # Set up user preferences
        self._setup_user_preferences()
        
        # Set voice preference globally
        say.voice_enabled = self.voice_recognition.voice_enabled
    
    def _setup_user_preferences(self):
        """Set up initial user preferences."""
        # Check if user has a name
        user_name = self.conv_manager.get_user_preference("name")
        if not user_name:
            print("Hello! I don't believe we've met before. What's your name?")
            name = input("You: ").strip()
            if name:
                self.conv_manager.set_user_preference("name", name)
                self.conv_manager.add_to_history("assistant", f"Nice to meet you, {name}!")
                say(f"Nice to meet you, {name}!")
    
    def _display_typing_indicator(self):
        """Display a typing indicator to simulate thinking."""
        print("AI: ", end="", flush=True)
        for _ in range(3):
            time.sleep(0.5)
            print(".", end="", flush=True)
        print("\rAI: ", end="", flush=True)
    
    def _handle_personal_question(self, question: str) -> str:
        """Handle personal questions about the AI."""
        question_lower = question.lower()
        
        if "name" in question_lower:
            return "I'm your local AI assistant, designed to help you with various tasks on your Windows computer."
        
        if "created" in question_lower or "made" in question_lower:
            return "I was created as a program to assist you with your computer tasks."
        
        if "learn" in question_lower or "remember" in question_lower:
            return "I can remember our conversation during this session, but I don't retain information between sessions."
        
        if "feel" in question_lower or "emotion" in question_lower:
            return "As an AI, I don't have emotions, but I'm designed to be helpful and friendly."
        
        if "help" in question_lower:
            return "I can help you with file operations, system controls, opening applications, and more. Just ask me what you need!"
        
        # Default response for other personal questions
        return "That's an interesting question about me. I'm here to assist you with your computer needs. Is there something specific I can help you with?"
    
    def _show_help(self):
        """Show help information in a conversational way."""
        user_name = self.conv_manager.get_user_preference("name", "there")
        help_text = f"""
Hi {user_name}! I can help you with many tasks:

Voice Commands:
- "Hey Assistant" - Wake me up
- "Stop listening" - Deactivate voice mode
- "Type mode" - Switch to text input
- "Voice mode" - Switch back to voice input

File Operations:
- Create, copy, move, rename, delete files/folders
- Search for files
- Open files with their default applications

System Controls:
- Adjust volume and brightness
- Control Wi-Fi
- Shutdown, restart, sleep, lock, or logout

Applications:
- Open and close applications
- Create documents, spreadsheets, presentations
- Work with images and media files

You can also ask me questions about science, tell me to tell jokes, or just have a conversation!

Just say what you want to do, like "Hey Assistant, create a folder called work in documents" or "what time is it?"
        """
        print(help_text)
        say("I've displayed some information about what I can help you with.")
    
    def start(self):
        """Start the chat interface with voice support."""
        user_name = self.conv_manager.get_user_preference("name", "there")
        
        # Check if voice recognition is available
        if self.voice_recognition.is_available():
            say(f"Hello {user_name}! I'm your AI assistant with voice recognition. Say 'Hey Assistant' to activate voice mode, or type your commands.")
        else:
            say(f"Hello {user_name}! I'm your AI assistant. Voice recognition is not available, but you can type your commands.")
        
        while True:
            try:
                user_input = None
                
                # Check for voice input if continuous listening is enabled
                if self.voice_recognition.continuous_listening:
                    user_input = self.voice_recognition.get_audio_command()
                    if user_input:
                        print(f"\nVoice: {user_input}")
                
                # If no voice input, get text input
                if not user_input:
                    user_input = input("\nYou: ")
                
                if not user_input.strip():
                    continue
                
                # Add user input to history
                self.conv_manager.add_to_history("user", user_input)
                
                # Display typing indicator
                self._display_typing_indicator()
                
                # Parse the intent
                intent = self.parser.parse(user_input)
                
                if not intent:
                    response = self.response_generator.generate_conversational_response(user_input)
                    say(response)
                    self.conv_manager.add_to_history("assistant", response)
                    continue
                
                # Handle special intents
                if intent.name == "quit":
                    response = self.response_generator.generate_response(intent)
                    say(response)
                    break
                
                if intent.name == "help":
                    self._show_help()
                    continue
                
                if intent.name == "cancel":
                    say("Operation cancelled.")
                    continue
                
                if intent.name == "conversational":
                    response = self.response_generator.generate_conversational_response(user_input)
                    say(response)
                    self.conv_manager.add_to_history("assistant", response)
                    continue
                
                if intent.name == "personal_question":
                    response = self._handle_personal_question(user_input)
                    say(response)
                    self.conv_manager.add_to_history("assistant", response)
                    continue
                
                if intent.name == "unknown":
                    response = self.response_generator.generate_response(intent, success=False)
                    say(response)
                    self.conv_manager.add_to_history("assistant", response)
                    continue
                
                # Voice control commands
                if intent.name == "stop_voice":
                    self.voice_recognition.stop_continuous_listening()
                    continue
                
                if intent.name == "start_voice":
                    if self.voice_recognition.is_available():
                        self.voice_recognition.start_continuous_listening()
                    else:
                        say("Voice recognition is not available.")
                    continue
                
                if intent.name == "type_mode":
                    self.voice_recognition.stop_continuous_listening()
                    self.voice_recognition.voice_enabled = False
                    say.voice_enabled = False
                    say("Switched to text mode.")
                    continue
                
                if intent.name == "voice_mode":
                    if self.voice_recognition.is_available():
                        self.voice_recognition.voice_enabled = True
                        say.voice_enabled = True
                        self.voice_recognition.start_continuous_listening()
                    else:
                        say("Voice recognition is not available.")
                    continue
                
                # Execute the intent
                try:
                    self.parser.dispatch(intent)
                    response = self.response_generator.generate_response(intent, success=True)
                except Exception as e:
                    response = self.response_generator.generate_response(intent, success=False, details=str(e))
                
                # Detect user mood and adjust response if needed
                mood = self.conv_enhancements.detect_user_mood(user_input)
                mood_response = self.conv_enhancements.respond_to_mood(mood)
                if mood_response and intent.name == "conversational":
                    response = mood_response
                
                # Adjust tone based on user preferences
                response = self.conv_enhancements.adjust_response_tone(response)
                
                say(response)
                self.conv_manager.add_to_history("assistant", response)
                
            except (EOFError, KeyboardInterrupt):
                print("\nGoodbye!")
                break


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

Voice Commands:
  "Hey Assistant" - Wake up the AI
  "Stop listening" - Deactivate voice mode
  "Type mode" - Switch to text input
  "Voice mode" - Switch back to voice input

  help | quit
"""
    )


def main():
    # Initialize the controllers
    sysc = SystemController()
    appc = ApplicationController()
    
    # Create and start the chat interface
    chat_interface = ChatInterface(sysc, appc)
    
    # Optional: accept a one-shot command via CLI args
    if len(sys.argv) > 1:
        cmd = " ".join(sys.argv[1:])
        intent = chat_interface.parser.parse(cmd)
        if intent:
            try:
                chat_interface.parser.dispatch(intent)
                response = chat_interface.response_generator.generate_response(intent, success=True)
                say(response)
            except Exception as e:
                response = chat_interface.response_generator.generate_response(intent, success=False, details=str(e))
                say(response)
        else:
            say("Sorry, I didn't understand that. Try 'help' for commands.")
        return
    
    # Start the interactive chat interface
    chat_interface.start()


if __name__ == "__main__":
    main()


# Local Windows AI Assistant (Offline) — Enhanced ChatGPT-like Version
# =============================================================================

"""
Local Windows AI Assistant (Offline) — Enhanced ChatGPT-like Version
=====================================================================

This single-file Python app runs locally on Windows and performs two categories of tasks:
1) Navigating Around (system-level operations)
2) Using Software Systems (operate common apps & files)

What's new in this version
--------------------------
- **ChatGPT-like conversational interface** with memory and context
- **Natural language understanding** for more intuitive interactions
- **Personality adaptation** based on user preferences
- **Conversation history** to maintain context across interactions
- **Enhanced error handling** with more user-friendly messages
- **Mood detection** and appropriate responses
- **Typing indicator** for more natural interaction flow
- **User profile** with preferences and name memory

Install (recommended, then run):
    py -m pip install --upgrade pip
    py -m pip install pyautogui pywinauto psutil send2trash pyttsx3 python-docx openpyxl python-pptx pillow PyPDF2 comtypes pycaw pywin32

Run:
    py local_ai.py

Examples you can type (no full paths needed!):
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

Conversational examples:
    How are you today?
    What's your name?
    Tell me a joke
    I'm feeling frustrated
    Can you help me with my files?
    Why did the system restart?
"""

from __future__ import annotations
import os
import re
import sys
import time
import shutil
import glob
import subprocess
import datetime
import random
from dataclasses import dataclass
from typing import Optional, List, Tuple, Dict, Any

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
    import openpyxl
except Exception:
    openpyxl = None

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
    say(random.choice(jokes))


def tell_quote():
    quotes = [
        "The only way to do great work is to love what you do. - Steve Jobs",
        "Believe you can and you're halfway there. - Theodore Roosevelt",
        "The future belongs to those who believe in the beauty of their dreams. - Eleanor Roosevelt",
        "You miss 100% of the shots you don't take. - Wayne Gretzky",
        "The best way to predict the future is to create it. - Peter Drucker",
    ]
    say(random.choice(quotes))


def get_time():
    say(f"The current time is {datetime.datetime.now().strftime('%I:%M %p')}")


def get_date():
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


# ----------------------- Conversation Manager -----------------------

class ConversationManager:
    """Manages conversation context and history for more natural interactions."""
    
    def __init__(self, max_history=20):
        self.history: List[Dict[str, Any]] = []
        self.max_history = max_history
        self.context: Dict[str, Any] = {}
        self.user_profile: Dict[str, Any] = {
            "name": None,
            "preferences": {}
        }
    
    def add_to_history(self, role: str, message: str):
        """Add a message to the conversation history."""
        self.history.append({
            "role": role, 
            "message": message, 
            "timestamp": time.time()
        })
        # Keep only the most recent messages
        if len(self.history) > self.max_history:
            self.history = self.history[-self.max_history:]
    
    def get_recent_context(self, count: int = 5) -> List[Dict[str, Any]]:
        """Get the most recent conversation context."""
        return self.history[-count:] if self.history else []
    
    def update_context(self, key: str, value: Any):
        """Update the conversation context."""
        self.context[key] = value
    
    def get_context(self, key: str) -> Any:
        """Get a value from the conversation context."""
        return self.context.get(key)
    
    def clear_context(self):
        """Clear the conversation context."""
        self.context = {}
    
    def set_user_preference(self, key: str, value: Any):
        """Set a user preference."""
        self.user_profile["preferences"][key] = value
    
    def get_user_preference(self, key: str, default: Any = None) -> Any:
        """Get a user preference."""
        return self.user_profile["preferences"].get(key, default)


# ----------------------- Response Generator -----------------------

class ResponseGenerator:
    """Generates more natural, conversational responses."""
    
    def __init__(self, conversation_manager: ConversationManager):
        self.conv_manager = conversation_manager
        self.response_templates = {
            "greeting": [
                "Hello! How can I help you today?",
                "Hi there! What can I do for you?",
                "Good to see you! What would you like to work on?",
            ],
            "farewell": [
                "Goodbye! Have a great day!",
                "See you later! Feel free to come back anytime.",
                "Until next time! Take care.",
            ],
            "confirmation": [
                "Got it. I'll {action} right away.",
                "Sure thing. {action} now.",
                "I'll {action} for you.",
            ],
            "error": [
                "I'm sorry, I couldn't {action}. Could you try again?",
                "Hmm, I had trouble {action}. Maybe we could try a different approach?",
                "I wasn't able to {action}. Is there another way I can help?",
            ],
            "unknown": [
                "I'm not sure how to help with that. Could you rephrase?",
                "I don't understand. Can you tell me more?",
                "That's beyond my current capabilities. Is there something else I can do?",
            ]
        }
    
    def generate_response(self, intent: 'Intent', success: bool = True, details: Optional[str] = None) -> str:
        """Generate a contextual response based on intent and execution result."""
        if intent.name == "greet":
            return self._get_random_template("greeting")
        
        if intent.name == "goodbye":
            return self._get_random_template("farewell")
        
        if success:
            if details:
                return f"I've successfully {self._intent_to_phrase(intent)}: {details}"
            else:
                return self._get_random_template("confirmation").format(
                    action=self._intent_to_phrase(intent)
                )
        else:
            return self._get_random_template("error").format(
                action=self._intent_to_phrase(intent)
            )
    
    def _get_random_template(self, category: str) -> str:
        """Get a random response template from a category."""
        return random.choice(self.response_templates.get(category, ["I'm not sure how to respond."]))
    
    def _intent_to_phrase(self, intent: 'Intent') -> str:
        """Convert an intent to a natural language phrase."""
        phrases = {
            "create_folder": "create the folder",
            "copy": "copy the file",
            "move": "move the file",
            "rename": "rename the file",
            "delete": "delete the file",
            "open_file": "open the file",
            "search": "search for files",
            "open_app": "open the application",
            "close_app": "close the application",
            "set_volume": "set the volume",
            "set_brightness": "adjust the brightness",
            "shutdown": "shutdown your computer",
            "restart": "restart your computer",
            "sleep": "put your computer to sleep",
            "lock": "lock your computer",
            "logout": "log you out",
            "write_doc": "create the document",
            "sheet_sum": "create the spreadsheet",
            "make_ppt": "create the presentation",
            "img_resize": "resize the image",
            "img_crop": "crop the image",
            "play_media": "play the media file",
            "open_pdf": "open the PDF",
            "search_pdf": "search the PDF",
            "antivirus_scan": "start the antivirus scan",
            "wifi_on": "turn on Wi-Fi",
            "wifi_off": "turn off Wi-Fi",
        }
        return phrases.get(intent.name, "perform that action")
    
    def generate_conversational_response(self, user_input: str) -> str:
        """Generate a response for purely conversational inputs."""
        user_input_lower = user_input.lower()
        
        # Check for follow-up questions
        if any(word in user_input_lower for word in ["why", "how", "what", "when", "where"]):
            return "That's an interesting question. I'd need more context to give you a good answer."
        
        # Check for expressions of emotion
        if any(word in user_input_lower for word in ["happy", "glad", "excited"]):
            return "I'm glad to hear that! Is there anything I can help you with today?"
        
        if any(word in user_input_lower for word in ["sad", "upset", "frustrated"]):
            return "I'm sorry to hear that. Maybe I can help with something to make your day better?"
        
        # Default conversational response
        return "I see. Is there something specific you'd like me to help you with?"


# ----------------------- Conversational Enhancements -----------------------

class ConversationalEnhancements:
    """Additional conversational capabilities for the AI assistant."""
    
    def __init__(self, conversation_manager: ConversationManager):
        self.conv_manager = conversation_manager
        self.personality_traits = {
            "helpfulness": 0.9,
            "friendliness": 0.8,
            "humor": 0.6,
            "formality": 0.5  # 0 = very casual, 1 = very formal
        }
    
    def adjust_response_tone(self, response: str) -> str:
        """Adjust the tone of responses based on user preferences and conversation context."""
        # Check if user prefers more formal or casual language
        formality = self.conv_manager.get_user_preference("formality", self.personality_traits["formality"])
        
        if formality < 0.3:  # Very casual
            response = response.replace("I am", "I'm")
            response = response.replace("I will", "I'll")
            response = response.replace("I have", "I've")
            response = response.replace("you are", "you're")
            response = response.replace("you will", "you'll")
            response = response.replace("do not", "don't")
        
        elif formality > 0.7:  # More formal
            response = response.replace("I'm", "I am")
            response = response.replace("I'll", "I will")
            response = response.replace("I've", "I have")
            response = response.replace("you're", "you are")
            response = response.replace("you'll", "you will")
            response = response.replace("don't", "do not")
        
        return response
    
    def detect_user_mood(self, user_input: str) -> str:
        """Simple mood detection based on user input."""
        user_input_lower = user_input.lower()
        
        # Positive indicators
        positive_words = ["happy", "glad", "great", "awesome", "fantastic", "love", "like", "enjoy"]
        if any(word in user_input_lower for word in positive_words):
            return "positive"
        
        # Negative indicators
        negative_words = ["sad", "angry", "frustrated", "hate", "dislike", "terrible", "awful", "bad"]
        if any(word in user_input_lower for word in negative_words):
            return "negative"
        
        # Neutral
        return "neutral"
    
    def respond_to_mood(self, mood: str) -> Optional[str]:
        """Generate a response based on detected user mood."""
        if mood == "positive":
            return "I'm glad to hear that! Is there anything I can help you with today?"
        elif mood == "negative":
            return "I'm sorry to hear that. Is there something I can do to help improve your day?"
        else:
            return None  # No special response for neutral mood


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

    def get_current_brightness(self) -> int:
        """Get current brightness level."""
        try:
            cmd = [
                "powershell",
                "(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness).CurrentBrightness"
            ]
            result = subprocess.run(cmd, capture_output=True, text=True)
            return int(result.stdout.strip())
        except Exception:
            return 50  # fallback

    def brightness_up(self, step: int = 10):
        current = self.get_current_brightness()
        new = min(100, current + step)
        self.set_brightness(new)

    def brightness_down(self, step: int = 10):
        current = self.get_current_brightness()
        new = max(0, current - step)
        self.set_brightness(new)

    def wifi_toggle(self, state: str):
        """Toggle Wi-Fi on/off."""
        try:
            # Get Wi-Fi interface name
            result = subprocess.run(
                ["netsh", "interface", "show", "interface"],
                capture_output=True, text=True
            )
            iface = None
            for line in result.stdout.splitlines():
                if "Wireless" in line or "Wi-Fi" in line:
                    parts = line.split()
                    iface = parts[-1]  # interface name is last column
                    break
            
            if not iface:
                say("Could not find Wi-Fi interface.")
                return
            
            cmd = ["netsh", "interface", "set", "interface", iface, f"admin={state}"]
            subprocess.run(cmd)
            say(f"Wi-Fi turned {state}")
        except Exception as e:
            say(f"Failed to toggle Wi-Fi: {str(e)}")

    def wifi_on(self):
        self.wifi_toggle("enabled")

    def wifi_off(self):
        self.wifi_toggle("disabled")


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
        if openpyxl is None:
            say("openpyxl not installed.")
            return
        path = norm(expand_alias(path))
        if not is_probable_path(path):
            path = os.path.join(ALIAS_MAP.get('documents', os.path.join(HOME, 'Documents')), path)
        ensure_dir(os.path.dirname(path))
        wb = openpyxl.Workbook()
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


class EnhancedIntentParser:
    """Enhanced intent parser with better conversational understanding."""
    
    def __init__(self, sysc: SystemController, appc: ApplicationController, conversation_manager: ConversationManager):
        self.sys = sysc
        self.app = appc
        self.conv_manager = conversation_manager

    def parse(self, text: str) -> Optional[Intent]:
        t_raw = text.strip()
        t = t_raw.lower()

        # First check if this is a follow-up to a previous command
        recent_context = self.conv_manager.get_recent_context(3)
        
        # Check for simple confirmations or denials
        if t in ["yes", "yeah", "sure", "ok", "do it", "please do"]:
            # Check if there's a pending action in context
            pending_action = self.conv_manager.get_context("pending_action")
            if pending_action:
                return pending_action
        
        if t in ["no", "nope", "don't", "stop", "cancel"]:
            # Clear any pending action
            self.conv_manager.update_context("pending_action", None)
            return Intent("cancel", tuple())
        
        # Check for conversational patterns
        if self._is_conversational(t_raw):
            return Intent("conversational", (t_raw,))
        
        # Check for personal questions to the AI
        if self._is_personal_question(t_raw):
            return Intent("personal_question", (t_raw,))
        
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

    def _is_conversational(self, text: str) -> bool:
        """Check if the input is purely conversational."""
        conversational_patterns = [
            r"^(how|what|why|when|where|who) (are|is|do|did|does|can|will|would|could|should)",
            r"^(i think|i feel|i believe|i want|i need|i wish)",
            r"^(tell me about|do you know|can you explain)",
            r"^(that's|that is|it's|it is) (good|bad|great|terrible|awesome|horrible)",
        ]
        
        text_lower = text.lower()
        return any(re.search(pattern, text_lower) for pattern in conversational_patterns)
    
    def _is_personal_question(self, text: str) -> bool:
        """Check if the input is a personal question to the AI."""
        personal_patterns = [
            r"(what|who) (are|is) you",
            r"(what|who) (do|does) you (do|be)",
            r"(where|when) (were|was) you (created|made|born)",
            r"(do|can) you (have|feel|think|believe)",
            r"(are|is) you (human|alive|real)",
            r"(do|can) you (learn|remember|know)",
        ]
        
        text_lower = text.lower()
        return any(re.search(pattern, text_lower) for pattern in personal_patterns)

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
        if name == "brightness_up": s.brightness_up(); return
        if name == "brightness_down": s.brightness_down(); return
        if name == "wifi_on": s.wifi_on(); return
        if name == "wifi_off": s.wifi_off(); return
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


# ----------------------- Chat Interface -----------------------

class ChatInterface:
    """A more interactive, chat-like interface for the AI assistant."""
    
    def __init__(self, sysc: SystemController, appc: ApplicationController):
        self.conv_manager = ConversationManager()
        self.response_generator = ResponseGenerator(self.conv_manager)
        self.parser = EnhancedIntentParser(sysc, appc, self.conv_manager)
        self.sysc = sysc
        self.appc = appc
        self.conv_enhancements = ConversationalEnhancements(self.conv_manager)
        
        # Set up user preferences
        self._setup_user_preferences()
    
    def _setup_user_preferences(self):
        """Set up initial user preferences."""
        # Check if user has a name
        user_name = self.conv_manager.get_user_preference("name")
        if not user_name:
            print("Hello! I don't believe we've met before. What's your name?")
            name = input("You: ").strip()
            if name:
                self.conv_manager.set_user_preference("name", name)
                self.conv_manager.add_to_history("assistant", f"Nice to meet you, {name}!")
                say(f"Nice to meet you, {name}!")
    
    def _display_typing_indicator(self):
        """Display a typing indicator to simulate thinking."""
        print("AI: ", end="", flush=True)
        for _ in range(3):
            time.sleep(0.5)
            print(".", end="", flush=True)
        print("\rAI: ", end="", flush=True)
    
    def _handle_personal_question(self, question: str) -> str:
        """Handle personal questions about the AI."""
        question_lower = question.lower()
        
        if "name" in question_lower:
            return "I'm your local AI assistant, designed to help you with various tasks on your Windows computer."
        
        if "created" in question_lower or "made" in question_lower:
            return "I was created as a program to assist you with your computer tasks."
        
        if "learn" in question_lower or "remember" in question_lower:
            return "I can remember our conversation during this session, but I don't retain information between sessions."
        
        if "feel" in question_lower or "emotion" in question_lower:
            return "As an AI, I don't have emotions, but I'm designed to be helpful and friendly."
        
        if "help" in question_lower:
            return "I can help you with file operations, system controls, opening applications, and more. Just ask me what you need!"
        
        # Default response for other personal questions
        return "That's an interesting question about me. I'm here to assist you with your computer needs. Is there something specific I can help you with?"
    
    def _show_help(self):
        """Show help information in a conversational way."""
        user_name = self.conv_manager.get_user_preference("name", "there")
        help_text = f"""
Hi {user_name}! I can help you with many tasks:

File Operations:
- Create, copy, move, rename, delete files/folders
- Search for files
- Open files with their default applications

System Controls:
- Adjust volume and brightness
- Control Wi-Fi
- Shutdown, restart, sleep, lock, or logout

Applications:
- Open and close applications
- Create documents, spreadsheets, presentations
- Work with images and media files

You can also ask me questions about science, tell me to tell jokes, or just have a conversation!

Just type what you want to do in natural language, like "create a folder called work in documents" or "what's the weather like?"
        """
        print(help_text)
        say("I've displayed some information about what I can help you with.")
    
    def start(self):
        """Start the chat interface."""
        user_name = self.conv_manager.get_user_preference("name", "there")
        say(f"Hello {user_name}! I'm your AI assistant. How can I help you today?")
        
        while True:
            try:
                user_input = input("\nYou: ")
                if not user_input.strip():
                    continue
                
                # Add user input to history
                self.conv_manager.add_to_history("user", user_input)
                
                # Display typing indicator
                self._display_typing_indicator()
                
                # Parse the intent
                intent = self.parser.parse(user_input)
                
                if not intent:
                    response = self.response_generator.generate_conversational_response(user_input)
                    say(response)
                    self.conv_manager.add_to_history("assistant", response)
                    continue
                
                # Handle special intents
                if intent.name == "quit":
                    response = self.response_generator.generate_response(intent)
                    say(response)
                    break
                
                if intent.name == "help":
                    self._show_help()
                    continue
                
                if intent.name == "cancel":
                    say("Operation cancelled.")
                    continue
                
                if intent.name == "conversational":
                    response = self.response_generator.generate_conversational_response(user_input)
                    say(response)
                    self.conv_manager.add_to_history("assistant", response)
                    continue
                
                if intent.name == "personal_question":
                    response = self._handle_personal_question(user_input)
                    say(response)
                    self.conv_manager.add_to_history("assistant", response)
                    continue
                
                if intent.name == "unknown":
                    response = self.response_generator.generate_response(intent, success=False)
                    say(response)
                    self.conv_manager.add_to_history("assistant", response)
                    continue
                
                # Execute the intent
                try:
                    self.parser.dispatch(intent)
                    response = self.response_generator.generate_response(intent, success=True)
                except Exception as e:
                    response = self.response_generator.generate_response(intent, success=False, details=str(e))
                
                # Detect user mood and adjust response if needed
                mood = self.conv_enhancements.detect_user_mood(user_input)
                mood_response = self.conv_enhancements.respond_to_mood(mood)
                if mood_response and intent.name == "conversational":
                    response = mood_response
                
                # Adjust tone based on user preferences
                response = self.conv_enhancements.adjust_response_tone(response)
                
                say(response)
                self.conv_manager.add_to_history("assistant", response)
                
            except (EOFError, KeyboardInterrupt):
                print("\nGoodbye!")
                break


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
    # Initialize the controllers
    sysc = SystemController()
    appc = ApplicationController()
    
    # Create and start the chat interface
    chat_interface = ChatInterface(sysc, appc)
    
    # Optional: accept a one-shot command via CLI args
    if len(sys.argv) > 1:
        cmd = " ".join(sys.argv[1:])
        intent = chat_interface.parser.parse(cmd)
        if intent:
            try:
                chat_interface.parser.dispatch(intent)
                response = chat_interface.response_generator.generate_response(intent, success=True)
                say(response)
            except Exception as e:
                response = chat_interface.response_generator.generate_response(intent, success=False, details=str(e))
                say(response)
        else:
            say("Sorry, I didn't understand that. Try 'help' for commands.")
        return
    
    # Start the interactive chat interface
    chat_interface.start()


if __name__ == "__main__":
    main()