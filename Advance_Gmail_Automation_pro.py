import os
import sys
import threading
import pickle
import json
import re
import shutil
import time
import base64
import random
import urllib.parse
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, StringVar, Toplevel
from tkinter.scrolledtext import ScrolledText
from datetime import datetime, date

# Dependency checks
missing_deps = []
try:
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
except ImportError as e:
    build = None
    missing_deps.append("google-api-python-client, google-auth-httplib2, google-auth-oauthlib")
    googleapi_error = str(e)
else:
    googleapi_error = None

try:
    from playwright.sync_api import sync_playwright
    from playwright_stealth import stealth_sync
except ImportError as e:
    sync_playwright = None
    stealth_sync = None
    missing_deps.append("playwright, playwright-stealth")
    playwright_error = str(e)
else:
    playwright_error = None

SCOPES = ['https://www.googleapis.com/auth/gmail.send', "https://www.googleapis.com/auth/gmail.readonly"]
CREDENTIALS_DIR = "credentials"
TOKENS_DIR = "tokens"
PROFILES_DIR = "profiles"
ATTACHMENTS_DIR = "attachments"
USAGE_FILE = "gmail_usage.json"
TOKENS_MAP_FILE = os.path.join(TOKENS_DIR, "tokens_map.json")
FORMULAS_FILE = "search_formulas.json"
TEMPLATE_FILE = "email_templates.json"
VERSION = "3.6.1"

for d in [CREDENTIALS_DIR, TOKENS_DIR, PROFILES_DIR, ATTACHMENTS_DIR]:
    os.makedirs(d, exist_ok=True)

def save_json(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, indent=2)

def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def save_json_safe(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, indent=2)

def load_json_safe(path):
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def save_pickle(path, obj):
    with open(path, "wb") as f:
        pickle.dump(obj, f)

def load_pickle(path):
    with open(path, "rb") as f:
        return pickle.load(f)

def is_gmail(email):
    email = email.strip()
    if not re.fullmatch(r"[a-zA-Z0-9._%+-]+@gmail\.com", email):
        return False
    for forbidden in ['"', "'", " ", "+", "%", "=", ",", "/", "\\"]:
        if forbidden in email:
            return False
    return True

def extract_gmails(text):
    return re.findall(r'[a-zA-Z0-9._%+-]+@gmail\.com', text)

def get_today_usage(gmail):
    if not os.path.exists(USAGE_FILE):
        return 0
    usage = load_json(USAGE_FILE)
    today = date.today().isoformat()
    return usage.get(gmail, {}).get(today, 0)

def increment_usage(gmail, count=1):
    today = date.today().isoformat()
    usage = load_json(USAGE_FILE) if os.path.exists(USAGE_FILE) else {}
    if gmail not in usage:
        usage[gmail] = {}
    if today not in usage[gmail]:
        usage[gmail][today] = 0
    usage[gmail][today] += count
    save_json(USAGE_FILE, usage)

def update_tokens_map(email, creds_path):
    m = load_json_safe(TOKENS_MAP_FILE)
    m[email] = creds_path
    save_json_safe(TOKENS_MAP_FILE, m)

def get_creds_path_from_email(email):
    m = load_json_safe(TOKENS_MAP_FILE)
    return m.get(email)

def gmail_authenticate(creds_path, token_path):
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request

    creds = None
    if token_path and os.path.exists(token_path):
        creds = load_pickle(token_path)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
            creds = flow.run_local_server(port=0)
        if token_path:
            save_pickle(token_path, creds)
    return creds

def gmail_send_message(service, sender, to, subject, message_text, attachment_path=None, html_message=None):
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    if html_message:
        message.attach(MIMEText(html_message, 'html'))
    else:
        message.attach(MIMEText(message_text, 'plain'))
    if attachment_path:
        filename = os.path.basename(attachment_path)
        with open(attachment_path, 'rb') as f:
            part = MIMEBase('application', "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
            message.attach(part)
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return service.users().messages().send(userId="me", body={'raw': raw}).execute()

def get_gmail_stats(token_path):
    try:
        creds = load_pickle(token_path)
        if creds and creds.valid:
            from googleapiclient.discovery import build
            service = build('gmail', 'v1', credentials=creds)
            profile = service.users().getProfile(userId="me").execute()
            email = profile.get("emailAddress", "Unknown")
            quota = 500
            used = get_today_usage(email)
            return email, quota, used
    except:
        pass
    return "Unknown", 500, 0

def get_chrome_profiles():
    profiles = []
    win_path = os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data")
    if os.path.exists(win_path):
        for d in os.listdir(win_path):
            profile_path = os.path.join(win_path, d)
            if os.path.isdir(profile_path) and (d == "Default" or d.startswith("Profile")):
                profiles.append((f"Chrome - {d}", profile_path))
    mac_path = os.path.expanduser("~/Library/Application Support/Google/Chrome")
    if os.path.exists(mac_path):
        for d in os.listdir(mac_path):
            profile_path = os.path.join(mac_path, d)
            if os.path.isdir(profile_path) and (d == "Default" or d.startswith("Profile")):
                profiles.append((f"Chrome - {d}", profile_path))
    linux_path = os.path.expanduser("~/.config/google-chrome")
    if os.path.exists(linux_path):
        for d in os.listdir(linux_path):
            profile_path = os.path.join(linux_path, d)
            if os.path.isdir(profile_path) and (d == "Default" or d.startswith("Profile")):
                profiles.append((f"Chrome - {d}", profile_path))
    edge_path = os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\User Data")
    if os.path.exists(edge_path):
        for d in os.listdir(edge_path):
            profile_path = os.path.join(edge_path, d)
            if os.path.isdir(profile_path) and (d == "Default" or d.startswith("Profile")):
                profiles.append((f"Edge - {d}", profile_path))
    brave_path = os.path.expandvars(r"%LOCALAPPDATA%\BraveSoftware\Brave-Browser\User Data")
    if os.path.exists(brave_path):
        for d in os.listdir(brave_path):
            profile_path = os.path.join(brave_path, d)
            if os.path.isdir(profile_path) and (d == "Default" or d.startswith("Profile")):
                profiles.append((f"Brave - {d}", profile_path))
    firefox_path = os.path.expandvars(r"%APPDATA%\Mozilla\Firefox\Profiles")
    if os.path.exists(firefox_path):
        for d in os.listdir(firefox_path):
            profile_path = os.path.join(firefox_path, d)
            if os.path.isdir(profile_path):
                profiles.append((f"Firefox - {d}", profile_path))
    return profiles

class CTAManager:
    def __init__(self, label):
        self.label = label
        self._running = False
        self._step = 0
        self._start_time = None
        self._current_action = ""
        self._action_detail = ""
        self._timer_thread = None
        self._error = False
        self._last_message = ""
        self._lock = threading.Lock()

    def start(self, action, detail=None):
        with self._lock:
            self._running = True
            self._error = False
            self._step = 0
            self._start_time = time.time()
            self._current_action = action
            self._action_detail = detail or ""
            self.label.after(0, lambda: self.label.config(text=f"▶ {action}..."))
            if not self._timer_thread or not self._timer_thread.is_alive():
                self._timer_thread = threading.Thread(target=self._animate, daemon=True)
                self._timer_thread.start()

    def set_action(self, action, detail=None):
        with self._lock:
            self._current_action = action
            self._action_detail = detail or ""

    def error(self, msg="Error Occurred!"):
        with self._lock:
            self._error = True
            self._running = False
            self._last_message = msg
            self.label.after(0, lambda: self.label.config(
                text=f"✖ {msg} (Click 'Ready' to reset)"))

    def stop(self, final_msg="Done."):
        with self._lock:
            self._running = False
            self._error = False
            self._last_message = final_msg
            self.label.after(0, lambda: self.label.config(text=f"✔ {final_msg}"))

    def reset(self):
        with self._lock:
            self._running = False
            self._error = False
            self._last_message = "Ready."
            self.label.after(0, lambda: self.label.config(text="Ready."))

    def _animate(self):
        spinner = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        while True:
            with self._lock:
                if not self._running:
                    break
                if self._error:
                    break
                elapsed = int(time.time() - self._start_time) if self._start_time else 0
                mins = elapsed // 60
                secs = elapsed % 60
                spin = spinner[self._step % len(spinner)]
                msg = f"{spin} {self._current_action}"
                if self._action_detail:
                    msg += f" ({self._action_detail})"
                msg += f" | Time: {mins:02d}:{secs:02d}"
                self.label.after(0, lambda m=msg: self.label.config(text=m))
                self._step += 1
            time.sleep(0.13)

def countdown_console(seconds, log_cb=None):
    for remaining in range(seconds, 0, -1):
        msg = f"[CAPTCHA/Pause] Waiting for {remaining} seconds before retrying..."
        print(msg, end="\r", flush=True)
        if log_cb:
            log_cb(msg, "WARN")
        time.sleep(1)
    print(" " * 80, end="\r")

def load_search_formulas():
    if os.path.exists(FORMULAS_FILE):
        try:
            with open(FORMULAS_FILE, "r", encoding="utf-8") as f:
                formulas = json.load(f)
                if isinstance(formulas, list):
                    return formulas
        except Exception:
            pass
    return [
        'https://www.google.com/search?q=site:{site}+"{niche}"+"{location}"+%22@gmail.com%22&num=100'
    ]

def save_search_formulas(formulas):
    try:
        with open(FORMULAS_FILE, "w", encoding="utf-8") as f:
            json.dump(formulas, f, indent=2)
    except Exception:
        pass

def load_templates():
    DEFAULT_TEMPLATE = {
        "name": "Default",
        "html": """<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>AN Graphic Studio | Premium Design Solutions</title>
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&family=Montserrat:wght@600;700&display=swap" rel="stylesheet">
  <style>
    @keyframes fadeIn {
      from { opacity: 0; transform: translateY(20px); }
      to { opacity: 1; transform: translateY(0); }
    }
    @keyframes pulse {
      0% { transform: scale(1); }
      50% { transform: scale(1.05); }
      100% { transform: scale(1); }
    }
    .bulletin-board {
      background-color: #ffffff;
      background-image: 
        linear-gradient(#f1f1f1 1px, transparent 1px),
        linear-gradient(90deg, #f1f1f1 1px, transparent 1px);
      background-size: 20px 20px;
      background-position: center;
      border-radius: 16px;
      box-shadow: 0 10px 30px rgba(0,0,0,0.1);
    }
    .services-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 15px;
      margin: 25px 0 30px;
    }
    @media only screen and (max-width: 600px) {
      .services-grid {
        grid-template-columns: 1fr !important;
      }
      table {
        width: 100% !important;
      }
    }
  </style>
</head>
<body style="font-family: 'Poppins', sans-serif; background: linear-gradient(135deg, #e6f0ff 0%, #f8faff 100%); padding: 30px; color: #333; margin: 0;">
  <!-- Main Container -->
  <table width="100%" cellpadding="0" cellspacing="0" style="max-width: 680px; margin: auto;" class="bulletin-board">
    <!-- Header -->
    <tr>
      <td align="center" style="padding: 40px 20px; position: relative;">
        <table width="100%" cellpadding="0" cellspacing="0">
          <tr>
            <td width="50%" align="right" style="padding-right: 20px; border-right: 1px solid #f1f1f1;">
              <img src="https://angfxstudio.infy.uk/wp-content/uploads/2025/06/modern-update-logo-transparent.png" alt="AN Graphic Studio" width="160" style="display:block; filter: drop-shadow(0 4px 8px rgba(0,0,0,0.1)); position: relative; animation: fadeIn 0.8s ease-out;">
            </td>
            <td width="50%" align="left" style="padding-left: 20px; vertical-align: middle;">
              <h1 style="color: #0057ff; margin: 0; font-size: 28px; letter-spacing: 1px; font-weight: 700; font-family: 'Montserrat', sans-serif; text-transform: uppercase;">AN GRAPHIC STUDIO</h1>
              <p style="color: #666; font-size: 14px; margin: 8px 0 0; font-weight: 300;">Designs That Speak Before You Do</p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    
    <!-- Main Content -->
    <tr>
      <td style="padding: 0 40px 40px; position: relative;">
        <div style="position: absolute; top: 20px; right: 20px; width: 40px; height: 40px; border-radius: 50%; background: rgba(0,87,255,0.1);"></div>
        <div style="position: absolute; bottom: 30px; left: 30px; width: 20px; height: 20px; border-radius: 50%; background: rgba(0,161,255,0.1);"></div>
        
        <h2 style="color: #0057ff; margin-bottom: 15px; font-size: 24px; font-weight: 600;">Hello! 👋</h2>
        <p style="font-size: 16px; line-height: 1.8; margin-bottom: 25px;">We're <strong style="color: #0057ff;">AN Graphic Studio</strong> – a creative powerhouse crafting <strong>visual identities</strong> and <strong>digital experiences</strong> that make brands unforgettable.</p>
        
        <!-- Services Grid -->
        <div class="services-grid">
          <div style="background: #f5f9ff; border-left: 3px solid #0057ff; padding: 15px; border-radius: 0 8px 8px 0;">
            <h3 style="margin: 0 0 8px 0; color: #0057ff; font-size: 15px;">🎨 Logo Design</h3>
            <p style="margin: 0; font-size: 13px; color: #555;">Unique brand identities</p>
          </div>
          <div style="background: #f5f9ff; border-left: 3px solid #00a1ff; padding: 15px; border-radius: 0 8px 8px 0;">
            <h3 style="margin: 0 0 8px 0; color: #00a1ff; font-size: 15px;">📱 Social Media</h3>
            <p style="margin: 0; font-size: 13px; color: #555;">Engaging content</p>
          </div>
          <div style="background: #f5f9ff; border-left: 3px solid #0057ff; padding: 15px; border-radius: 0 8px 8px 0;">
            <h3 style="margin: 0 0 8px 0; color: #0057ff; font-size: 15px;">📺 YouTube</h3>
            <p style="margin: 0; font-size: 13px; color: #555;">Thumbnails & banners</p>
          </div>
          <div style="background: #f5f9ff; border-left: 3px solid #00a1ff; padding: 15px; border-radius: 0 8px 8px 0;">
            <h3 style="margin: 0 0 8px 0; color: #00a1ff; font-size: 15px;">📢 Marketing</h3>
            <p style="margin: 0; font-size: 13px; color: #555;">Posters & banners</p>
          </div>
        </div>
        
        <p style="font-size: 15px; line-height: 1.8; margin: 25px 0; padding-left: 20px; border-left: 3px solid #e6f0ff;">
          Trusted by clients across <strong style="color: #0057ff;">USA, UK, Canada, Germany</strong> and worldwide — we deliver <span style="background: linear-gradient(135deg, #0057ff, #00a1ff); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-weight: bold;">high-impact visuals that drive results</span>.
        </p>
        
        <!-- CTA Buttons -->
        <div style="text-align: center; margin: 40px 0 30px;">
          <a href="https://www.angraphicstudio.infy.uk" target="_blank" style="display:inline-block; background: linear-gradient(135deg, #0057ff 0%, #00a1ff 100%); color: white; padding: 14px 30px; border-radius: 50px; text-decoration: none; font-weight: 600; margin: 0 10px 15px 0; box-shadow: 0 4px 15px rgba(0,87,255,0.3); transition: all 0.3s; animation: pulse 2s infinite;">🌐 Visit Website</a>
          <a href="https://www.behance.net/AnGraphicStudio" target="_blank" style="display:inline-block; background: linear-gradient(135deg, #0057ff 0%, #00a1ff 100%); color: white; padding: 14px 30px; border-radius: 50px; text-decoration: none; font-weight: 600; margin: 0 0 15px 10px; box-shadow: 0 4px 15px rgba(0,87,255,0.3); transition: all 0.3s;">🎨 View Portfolio</a>
        </div>
        
        <!-- WhatsApp Button -->
        <div style="text-align: center; position: relative;">
          <div style="display: inline-block; position: relative;">
            <a href="https://wa.me/8801902678986" style="background-color: #25D366; color: white; padding: 14px 40px; text-decoration: none; border-radius: 50px; font-weight: 600; display:inline-block; box-shadow: 0 4px 15px rgba(37,211,102,0.3); transition: all 0.3s; position: relative; z-index: 1;">
              💬 Chat on WhatsApp
            </a>
            <div style="position: absolute; top: -5px; left: -5px; right: -5px; bottom: -5px; border: 1px solid rgba(37,211,102,0.3); border-radius: 55px; z-index: 0;"></div>
          </div>
          <p style="font-size: 14px; color: #666; text-align: center; margin-top: 30px;">
            Or simply reply to this email — let's create something extraordinary together!
          </p>
        </div>
      </td>
    </tr>
    
    <!-- Footer -->
    <tr>
      <td style="background: linear-gradient(135deg, #f5f9ff 0%, #e6f0ff 100%); text-align: center; padding: 25px; font-size: 13px; color: #555;">
        <div style="height: 1px; background: linear-gradient(90deg, rgba(0,87,255,0) 0%, rgba(0,87,255,0.3) 50%, rgba(0,87,255,0) 100%); width: 80%; margin: 0 auto 20px;"></div>
        <img src="https://angfxstudio.infy.uk/wp-content/uploads/2025/06/modern-update-logo-transparent.png" alt="Logo" width="80" style="opacity: 0.7; margin-bottom: 15px;">
        <p style="margin: 5px 0;">
          © 2025 AN Graphic Studio. All rights reserved.
        </p>
        <p style="margin: 5px 0; font-size: 12px;">
          Crafted with ❤️ by Alamin | <a href="mailto:angraphicstudio@gmail.com" style="color: #0057ff; text-decoration: none;">angraphicstudio@gmail.com</a>
        </p>
        <div style="margin-top: 15px;">
          <a href="#"><img src="https://cdn-icons-png.flaticon.com/512/2111/2111463.png" width="20" alt="Instagram" style="margin: 0 5px;"></a>
          <a href="#"><img src="https://cdn-icons-png.flaticon.com/512/733/733579.png" width="20" alt="Twitter" style="margin: 0 5px;"></a>
          <a href="#"><img src="https://cdn-icons-png.flaticon.com/512/3536/3536505.png" width="20" alt="LinkedIn" style="margin: 0 5px;"></a>
          <a href="#"><img src="https://cdn-icons-png.flaticon.com/512/2504/2504903.png" width="20" alt="Dribbble" style="margin: 0 5px;"></a>
        </div>
      </td>
    </tr>
  </table>
</body>
</html>
"""
    }
    # Always ensure there's at least the default template.
    if not os.path.exists(TEMPLATE_FILE):
        with open(TEMPLATE_FILE, "w", encoding="utf-8") as f:
            json.dump([DEFAULT_TEMPLATE], f, indent=2)
    try:
        with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
            templates = json.load(f)
    except Exception:
        templates = [DEFAULT_TEMPLATE]
    # Ensure default is there and at the top
    if not any(t.get("name") == "Default" for t in templates):
        templates.insert(0, DEFAULT_TEMPLATE)
    else:
        # Move default to top
        templates = sorted(templates, key=lambda t: 0 if t.get("name") == "Default" else 1)
    # Remove duplicates by name, keeping the first
    seen = set()
    unique_templates = []
    for t in templates:
        if t["name"] not in seen:
            unique_templates.append(t)
            seen.add(t["name"])
    return unique_templates

def save_templates(templates):
    with open(TEMPLATE_FILE, "w", encoding="utf-8") as f:
        json.dump(templates, f, indent=2)

def add_or_update_template(name, html):
    templates = load_templates()
    found = False
    for t in templates:
        if t["name"] == name:
            t["html"] = html
            found = True
            break
    if not found:
        templates.append({"name": name, "html": html})
    save_templates(templates)

def playwright_gmail_leads(site, niche, location, quantity, log_cb, progress_cb, cta_mgr,
                           gui_parent=None, formula=None, abort_event=None, two_min_captcha=True):
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        log_cb("Playwright not installed.", "ERROR")
        cta_mgr.error("Playwright missing!")
        return []

    if not formula:
        formula = 'https://www.google.com/search?q=site:{site}+"{niche}"+"{location}"+%22@gmail.com%22&num=100'

    try:
        search_url = formula.format(site=site, niche=niche, location=location)
    except Exception:
        search_url = f'https://www.google.com/search?q=site:{site}+"{niche}"+"{location}"+%22@gmail.com%22&num=100'

    emails = []
    seen = set()
    page_num = 1
    user_closed = False

    user_agent = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/123.0.0.0 Safari/537.36"
    )

    cta_mgr.set_action("Launching browser", None)
    log_cb("[✔] Browser opened successfully", "INFO")

    total_attempts = 0
    max_attempts = 20

    while len(emails) < quantity and total_attempts < max_attempts and not user_closed \
            and (not abort_event or not abort_event.is_set()):
        total_attempts += 1
        try:
            with sync_playwright() as p:
                profiles = get_chrome_profiles()
                profile_dir = None

                def pick_profile_dialog():
                    dlg = Toplevel(gui_parent) if gui_parent else Toplevel()
                    dlg.title("Select Browser Profile (Auto/Manual)")
                    dlg.geometry("490x330")
                    ttk.Label(dlg, text="Select a detected browser profile:").pack(pady=(10, 2))
                    profile_list = tk.Listbox(dlg, width=68, height=10)
                    for name, path in profiles:
                        profile_list.insert(tk.END, f"{name}: {path}")
                    profile_list.pack(pady=(3, 5))
                    ttk.Label(dlg, text="Or click 'Browse' to select any folder.").pack(pady=(2, 2))

                    selected_profile = []
                    def select_profile():
                        idx = profile_list.curselection()
                        if idx:
                            selected_profile.append(profiles[idx[0]][1])
                            dlg.destroy()
                    def browse_profile():
                        folder = filedialog.askdirectory(title="Select a browser profile folder")
                        if folder:
                            selected_profile.append(folder)
                            dlg.destroy()
                    ttk.Button(dlg, text="Select", command=select_profile).pack(side="left", padx=(40,10), pady=6)
                    ttk.Button(dlg, text="Browse", command=browse_profile).pack(side="left", padx=10, pady=6)
                    ttk.Button(dlg, text="Default (headless/no profile)", command=dlg.destroy).pack(side="right", padx=(10,30), pady=6)
                    dlg.grab_set()
                    dlg.wait_window()
                    return selected_profile[0] if selected_profile else None

                profile_dir = pick_profile_dialog()
                if profile_dir and os.path.exists(profile_dir):
                    browser = p.chromium.launch_persistent_context(
                        user_data_dir=profile_dir,
                        headless=False,
                        args=[
                            "--disable-blink-features=AutomationControlled",
                            "--disable-infobars",
                            "--start-maximized"
                        ]
                    )
                    page = browser.new_page()
                    log_cb(f"[INFO] Using selected browser profile: {profile_dir}", "INFO")
                else:
                    browser = p.chromium.launch(
                        headless=False,
                        args=[
                            "--disable-blink-features=AutomationControlled",
                            "--disable-infobars",
                            "--start-maximized"
                        ]
                    )
                    context = browser.new_context(
                        user_agent=user_agent,
                        viewport={"width": 1366, "height": 768},
                        locale="en-US",
                        timezone_id="America/New_York"
                    )
                    page = context.new_page()
                    log_cb(f"[INFO] Using default browser (no custom profile).", "INFO")

                cta_mgr.set_action("Opening Google with formula...", None)
                log_cb(f"[🕒] Navigating to Google with formula: {search_url}", "INFO")
                try:
                    page.goto(search_url, timeout=60000)
                except Exception as e:
                    log_cb(f"[❌] Failed to open search URL: {e}", "ERROR")
                    cta_mgr.error("Failed to open search URL")
                    try:
                        browser.close()
                    except Exception:
                        pass
                    raise

                if two_min_captcha:
                    cta_mgr.set_action("Captcha solving: Wait 2 minutes", None)
                    log_cb("[🔒] Waiting for manual CAPTCHA solving (2 minutes)...", "WARN")
                    for remaining in range(120, 0, -1):
                        if abort_event and abort_event.is_set():
                            user_closed = True
                            break
                        try:
                            if "search" in page.url:
                                pass
                        except Exception as e:
                            if "closed" in str(e).lower():
                                log_cb("[USER] Browser closed manually during captcha. Stopping...", "ERROR")
                                user_closed = True
                                if abort_event:
                                    abort_event.set()
                                break
                        if remaining % 10 == 0 or remaining < 11:
                            log_cb(f"[⏳] Time left: {remaining} seconds", "INFO")
                        time.sleep(1)
                    if user_closed:
                        try:
                            browser.close()
                        except Exception:
                            pass
                        break

                cta_mgr.set_action("Google loaded. Scraping starts now", None)
                log_cb("[✅] Google search loaded, starting scraping.", "SUCCESS")

                def scroll_page():
                    try:
                        page.evaluate(
                            "() => { window.scrollBy(0, document.body.scrollHeight); }"
                        )
                        time.sleep(1.0)
                    except Exception:
                        pass

                next_selectors_aria = [
                    "a[aria-label='Next']",
                    "a[aria-label='পরবর্তী']",
                ]
                next_selectors_text = [
                    "span:has-text('Next')",
                    "span:has-text('পরবর্তী')",
                ]
                next_selectors_numeric = [
                    lambda n: f"a:has-text('{n}')",
                    lambda n: f"a[href*='start={(n-1)*10}']",
                    lambda n: f"button:has-text('{n}')"
                ]

                while len(emails) < quantity and not user_closed and (not abort_event or not abort_event.is_set()):
                    try:
                        log_cb(f"[DEBUG] Scraping emails from page {page_num}, URL: {page.url}", "INFO")
                        for try_num in range(5):
                            try:
                                page.wait_for_load_state("networkidle", timeout=30000)
                                content = page.content()
                                break
                            except Exception as e:
                                if "navigating" in str(e).lower():
                                    time.sleep(2)
                                    continue
                                elif "has been closed" in str(e).lower() or "browser has been closed" in str(e).lower() or "closed" in str(e).lower():
                                    log_cb("[USER] Browser closed manually during scraping (inner). Stopping...", "ERROR")
                                    user_closed = True
                                    if abort_event:
                                        abort_event.set()
                                    break
                                else:
                                    raise
                        else:
                            log_cb("Failed to get page content after retries.", "ERROR")
                            break

                        if user_closed or (abort_event and abort_event.is_set()):
                            break

                        page_emails = re.findall(r"[a-zA-Z0-9._%+-]+@gmail\.com", content)
                        new_emails = [e for e in page_emails if e not in seen]
                        emails.extend(new_emails)
                        for e in new_emails:
                            seen.add(e)
                        emails = list(sorted(set(emails)))
                        found_now = len(new_emails)

                        log_cb(
                            f"[📨] Found {found_now} Gmail(s) in Page {page_num} (Total: {len(emails)})",
                            "INFO",
                        )
                        progress_cb(len(emails), quantity)

                        if len(emails) >= quantity:
                            break

                        scroll_page()

                        fallback_attempted = False

                        for selector in next_selectors_aria:
                            try:
                                btn = page.query_selector(selector)
                                if btn and btn.is_visible():
                                    log_cb(f"[➡] Clicking 'Next' via aria-label: {selector}", "INFO")
                                    btn.click()
                                    fallback_attempted = True
                                    break
                            except Exception:
                                continue
                        if not fallback_attempted:
                            for selector in next_selectors_text:
                                try:
                                    btn = page.query_selector(selector)
                                    if btn and btn.is_visible():
                                        log_cb(f"[➡] Clicking next via text selector: {selector}", "INFO")
                                        btn.click()
                                        fallback_attempted = True
                                        break
                                except Exception:
                                    continue
                        if not fallback_attempted:
                            num = page_num + 1
                            for get_selector in next_selectors_numeric:
                                selector = get_selector(num)
                                try:
                                    btn = page.query_selector(selector)
                                    if btn and btn.is_visible():
                                        log_cb(f"[➡] Clicking numeric page button: {selector}", "INFO")
                                        btn.click()
                                        fallback_attempted = True
                                        break
                                except Exception:
                                    continue
                        if not fallback_attempted:
                            log_cb("[⚠] No pagination element found. Stopping pagination.", "WARN")
                            log_cb("[INFO] Pausing scraping for 3 minutes (180 seconds).", "INFO")
                            countdown_console(180, log_cb)
                            log_cb("[INFO] Resuming scraping after pause.", "INFO")
                            continue
                        time.sleep(2)
                        page_num += 1

                    except Exception as e:
                        err_msg = str(e).lower()
                        if "has been closed" in err_msg or "browser has been closed" in err_msg or "targetclosederror" in err_msg or "browser closed" in err_msg or "closed" in err_msg:
                            log_cb("[USER] Browser closed manually during scraping. Stopping all automation...", "ERROR")
                            user_closed = True
                            if abort_event:
                                abort_event.set()
                            break
                        log_cb(f"[❌] Exception during scraping (likely CAPTCHA or browser closed): {e}", "ERROR")
                        print("\n[!] CAPTCHA or browser error detected. Pausing scraping for 3 minutes (180 seconds).", flush=True)
                        countdown_console(180, log_cb)
                        log_cb("[INFO] Resuming scraping after pause.", "INFO")
                        try:
                            if browser:
                                browser.close()
                        except Exception:
                            pass
                        break

                try:
                    browser.close()
                except Exception:
                    pass

            if len(emails) >= quantity or user_closed or (abort_event and abort_event.is_set()):
                break

        except Exception as outer_e:
            err_msg = str(outer_e).lower()
            if "has been closed" in err_msg or "browser has been closed" in err_msg or "targetclosederror" in err_msg or "browser closed" in err_msg or "closed" in err_msg:
                log_cb("[USER] Browser closed manually during scraping (outer). Stopping all automation...", "ERROR")
                user_closed = True
                if abort_event:
                    abort_event.set()
                break
            log_cb(f"[❌] Outer exception during Playwright session: {outer_e}", "ERROR")
            print("\n[!] Outer CAPTCHA or browser error detected. Pausing scraping for 3 minutes (180 seconds).", flush=True)
            countdown_console(180, log_cb)
            log_cb("[INFO] Resuming scraping after pause.", "INFO")
            continue

    emails = list(sorted(set(emails)))
    log_cb(
        f"[✅] Gmail extraction complete. Total: {len(emails)}", "SUCCESS"
    )
    return emails[:quantity]

def playwright_launch_gmail(profile_path, log_cb, cta_mgr):
    cta_mgr.set_action("Launching browser for login", profile_path)
    log_cb(f"Launching browser for login at {profile_path}", "INFO")
    time.sleep(2)
    cta_mgr.stop("Login Completed")

def playwright_open_profile(profile_path, log_cb, cta_mgr):
    cta_mgr.set_action("Opening saved profile", profile_path)
    log_cb(f"Opening saved browser profile at {profile_path}", "INFO")
    time.sleep(2)
    cta_mgr.stop("Profile Opened")

class GmailProApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gmail Automation Pro v1.0")
        self.root.geometry("1020x650")
        self.root.configure(bg="#f9f9f9")
        self.status_var = StringVar(value="Ready.")
        self.progress = tk.IntVar(value=0)
        self.start_time = None
        self.task_running = False

        # Copyright header
        self.header_copyright = ttk.Label(
            self.root,
            text="Copyright © AnGraphicStudio 2025. All rights reserved.",
            font=("Segoe UI", 8),
            foreground="gray",
            anchor='center'
        )
        self.header_copyright.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(4, 0))

        self.default_formula = 'https://www.google.com/search?q=site:{site}+"{niche}"+"{location}"+%22@gmail.com%22&num=100'
        self.formulas = load_search_formulas()
        if self.default_formula not in self.formulas:
            self.formulas.insert(0, self.default_formula)
        self.formulas_used = list(self.formulas)
        self.selected_formula = tk.StringVar(value=self.formulas_used[0])
        self.abort_event = threading.Event()

        # Email templates
        self.templates = load_templates()
        self.selected_template_var = tk.StringVar(value=self.templates[0]["name"] if self.templates else "Default")

        # Main grid
        self.left_frame = ttk.Frame(self.root, padding=(10, 4))
        self.left_frame.grid(row=1, column=0, sticky="nsew", padx=(12, 10), pady=8)
        self.right_frame = ttk.Frame(self.root, padding=(7, 2))
        self.right_frame.grid(row=1, column=1, sticky="nsew", padx=(2, 10), pady=8)
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=0)
        self.root.rowconfigure(1, weight=1)

        self.cta_label = ttk.Label(self.left_frame, text="Ready.", font=("Segoe UI", 10, "bold"), foreground="#003366")
        self.cta_label.grid(row=99, column=0, sticky="ew", pady=(7, 0), columnspan=1)
        self.cta_mgr = CTAManager(self.cta_label)

        self.api_group = ttk.LabelFrame(self.left_frame, text="API Account")
        self.api_group.grid(row=0, column=0, sticky="ew", pady=(0, 7))
        self.api_group.columnconfigure(0, weight=1)
        ttk.Button(self.api_group, text="Upload credentials.json", command=self.upload_credentials).grid(row=0, column=0, sticky="w", pady=1)
        ttk.Button(self.api_group, text="Authorize & Save Token", command=self.add_gmail_api_account).grid(row=0, column=1, sticky="w", padx=3, pady=1)
        self.api_accounts_box = ttk.Treeview(self.api_group, columns=("Account", "Used", "Quota"), show="headings", height=1)
        for col in ("Account", "Used", "Quota"):
            self.api_accounts_box.heading(col, text=col)
            self.api_accounts_box.column(col, width=92)
        self.api_accounts_box.grid(row=1, column=0, columnspan=2, pady=2, sticky="ew")
        self.refresh_api_accounts()

        self.profile_group = ttk.LabelFrame(self.left_frame, text="Profile (Browser)")
        self.profile_group.grid(row=1, column=0, sticky="ew", pady=(0, 7))
        ttk.Button(self.profile_group, text="Login & Save Profile", command=lambda: self.add_gmail_profile()).grid(row=0, column=0, sticky="w")
        ttk.Button(self.profile_group, text="Open Saved Profile", command=lambda: self.open_gmail_profile()).grid(row=0, column=1, sticky="w", padx=3)

        self.email_group = ttk.LabelFrame(self.left_frame, text="Email List")
        self.email_group.grid(row=2, column=0, sticky="ew", pady=(0, 7))
        ttk.Label(self.email_group, text="Comma separated:").grid(row=0, column=0, sticky="w")
        self.email_entry = ttk.Entry(self.email_group, width=35)
        self.email_entry.grid(row=0, column=1, sticky="w")
        ttk.Button(self.email_group, text="Add", command=self.add_email_list).grid(row=0, column=2, sticky="w", padx=2)
        self.email_list_text = ScrolledText(self.email_group, height=1, width=36)
        self.email_list_text.grid(row=1, column=0, columnspan=3, pady=2)
        self.gmails_list = []

        self.send_group = ttk.LabelFrame(self.left_frame, text="Send Automation")
        self.send_group.grid(row=3, column=0, sticky="ew", pady=(0, 7))
        ttk.Label(self.send_group, text="API Account:").grid(row=0, column=0, sticky="w")
        self.accounts_listbox = tk.Listbox(self.send_group, selectmode=tk.MULTIPLE, width=18, height=1)
        self.accounts_listbox.grid(row=0, column=1, sticky="w")
        ttk.Button(self.send_group, text="Refresh", command=self.refresh_accounts_listbox).grid(row=0, column=2, sticky="w", padx=1)

        # HTML Template Controls
        ttk.Label(self.send_group, text="HTML Template:").grid(row=1, column=0, sticky="w", pady=(3,0))
        self.template_combo = ttk.Combobox(self.send_group, textvariable=self.selected_template_var, values=[t["name"] for t in self.templates], width=24, state="readonly")
        self.template_combo.grid(row=1, column=1, sticky="w", pady=(3,0))
        ttk.Button(self.send_group, text="Add/Edit Template", command=self.open_template_dialog).grid(row=1, column=2, sticky="w", padx=1, pady=(3,0))

        ttk.Label(self.send_group, text="Message:").grid(row=2, column=0, sticky="nw", pady=(3,0))
        self.message_text = ScrolledText(self.send_group, height=1, width=28)
        self.message_text.grid(row=2, column=1, pady=(3,0), sticky="w")
        ttk.Button(self.send_group, text="Attachment", command=self.browse_attachment).grid(row=2, column=2, sticky="nw", padx=1, pady=(3,0))
        self.attachment_label = ttk.Label(self.send_group, text="None")
        self.attachment_label.grid(row=3, column=1, sticky="w")
        ttk.Button(self.send_group, text="Send", command=self.start_automation).grid(row=3, column=2, sticky="e", pady=2)
        self.automation_stats = ttk.Label(self.send_group, text="Sent: 0 | Skipped: 0 | Errors: 0")
        self.automation_stats.grid(row=4, column=1, sticky="w")
        self.attachment_path = None

        self.finder_group = ttk.LabelFrame(self.left_frame, text="Gmail Finder")
        self.finder_group.grid(row=4, column=0, sticky="ew", pady=(0, 7))

        formula_panel = ttk.Frame(self.finder_group)
        formula_panel.grid(row=0, column=0, columnspan=4, sticky="ew", pady=(1,3))
        ttk.Label(formula_panel, text="Formula:").grid(row=0, column=0, sticky="w")
        self.formula_combo = ttk.Combobox(formula_panel, textvariable=self.selected_formula, values=self.formulas_used, width=25)
        self.formula_combo.grid(row=0, column=1, sticky="ew", padx=(2,3))
        formula_panel.columnconfigure(1, weight=1)
        self.new_formula_entry = ttk.Entry(formula_panel, width=20)
        self.new_formula_entry.grid(row=0, column=2, padx=(2,1))
        ttk.Button(formula_panel, text="+ Add", width=7, command=self.add_formula).grid(row=0, column=3, padx=(2,1))
        ttk.Button(formula_panel, text="Default", width=7, command=self.set_default_formula).grid(row=0, column=4, padx=(2,1))

        ttk.Label(self.finder_group, text="Site:").grid(row=1, column=0, sticky="w")
        self.site_entry = ttk.Combobox(
            self.finder_group,
            values=["instagram.com", "facebook.com", "x.com", "youtube.com"],
            width=12
        )
        self.site_entry.set("instagram.com")
        self.site_entry.grid(row=1, column=1, sticky="w")
        ttk.Label(self.finder_group, text="Niche:").grid(row=1, column=2, sticky="w")
        self.niche_entry = ttk.Combobox(
            self.finder_group,
            values=["Photographer", "Designer", "Salon", "Agency", "Marketing", "Content Creator"],
            width=12
        )
        self.niche_entry.set("Photographer")
        self.niche_entry.grid(row=1, column=3, sticky="w")
        ttk.Label(self.finder_group, text="Location:").grid(row=2, column=0, sticky="w")
        self.location_entry = ttk.Entry(self.finder_group, width=13)
        self.location_entry.insert(0, "Los Angeles")
        self.location_entry.grid(row=2, column=1, sticky="w")
        ttk.Label(self.finder_group, text="Qty:").grid(row=2, column=2, sticky="w")
        self.quantity_entry = ttk.Entry(self.finder_group, width=5)
        self.quantity_entry.insert(0, "20")
        self.quantity_entry.grid(row=2, column=3, sticky="w")
        ttk.Button(self.finder_group, text="Find", command=self.find_gmails_leads).grid(row=3, column=1, sticky="w", pady=2)
        ttk.Button(self.finder_group, text="Copy All", command=self.copy_leads).grid(row=3, column=2, sticky="w", pady=2)
        ttk.Button(self.finder_group, text="Stop All", command=self.abort_all).grid(row=3, column=3, sticky="ew", pady=2)

        leads_box_frame = ttk.Frame(self.finder_group)
        leads_box_frame.grid(row=4, column=0, columnspan=4, pady=2, sticky="ew")
        leads_box_frame.columnconfigure(0, weight=1)
        leads_box_frame.columnconfigure(1, weight=0)

        self.leads_text = ScrolledText(leads_box_frame, height=8, width=54, wrap=tk.WORD)
        self.leads_text.grid(row=0, column=0, sticky="ewns")
        self.leads_text.bind("<<Modified>>", self.update_leads_count_event)

        side_panel = ttk.Frame(leads_box_frame)
        side_panel.grid(row=0, column=1, sticky="nsw", padx=(7,0))

        self.leads_count_label = ttk.Label(side_panel, text="Count: 0", foreground="#0055aa", font=("Segoe UI", 9, "bold"))
        self.leads_count_label.pack(anchor="ne", pady=(2,2))

        self.check_gmail_btn = ttk.Button(side_panel, text="Check Real Gmails", width=15, command=self.check_real_gmails)
        self.check_gmail_btn.pack(anchor="ne", pady=(2,2))

        self.refresh_leads_btn = ttk.Button(side_panel, text="Refresh", width=15, command=self.refresh_leads_section)
        self.refresh_leads_btn.pack(anchor="ne", pady=(2,2))

        self.stats_group = ttk.LabelFrame(self.left_frame, text="Usage Stats")
        self.stats_group.grid(row=5, column=0, sticky="ew")
        self.stats_box = ScrolledText(self.stats_group, height=1, width=38)
        self.stats_box.grid(row=0, column=0, columnspan=2, pady=2)
        ttk.Button(self.stats_group, text="Refresh", command=self.refresh_stats).grid(row=1, column=1, sticky="e")

        ttk.Label(self.right_frame, text="Console Output", font=("Segoe UI", 9, "bold")).pack(anchor="nw", pady=(0,2))
        self.console_panel = ScrolledText(self.right_frame, height=26, width=41, state="normal", font=("Consolas", 9))
        self.console_panel.pack(fill=tk.BOTH, expand=True, padx=(0,0), pady=(0,0))
        ttk.Label(self.right_frame, textvariable=self.status_var, font=("Segoe UI", 8)).pack(anchor="sw", pady=(5, 0))
        self.progress_bar = ttk.Progressbar(self.right_frame, orient="horizontal", mode="determinate", length=278, variable=self.progress)
        self.progress_bar.pack(anchor="sw", pady=(1, 0))

        self.factory_reset_btn = tk.Button(
            self.root, text="Factory Reset", font=("Segoe UI", 8), bg="#b22222", fg="white",
            command=self.factory_reset, relief="raised", padx=3, pady=1
        )
        self.factory_reset_btn.place(relx=0.99, rely=0.02, anchor="ne")
        ttk.Button(self.left_frame, text="FACTORY RESET", command=self.factory_reset, style="Danger.TButton").grid(row=101, column=0, sticky="ew", pady=(17,0))
        s = ttk.Style()
        s.configure("Danger.TButton", foreground="white", background="#b22222")
        s.map("Danger.TButton",
              foreground=[('active', 'white')],
              background=[('active', '#ff5555')])

        if missing_deps:
            self.log("Missing dependencies: " + ", ".join(missing_deps), "ERROR")
            self.set_status("Some dependencies missing! See console.", 0)
            self.cta_mgr.error("Dependencies missing!")

        self.update_leads_count()

    def set_default_formula(self):
        self.selected_formula.set(self.default_formula)
        self.formula_combo.set(self.default_formula)

    def add_formula(self):
        formula = self.new_formula_entry.get().strip()
        if not formula:
            messagebox.showerror("Error", "Please type a formula to add.")
            return
        if formula in self.formulas_used:
            messagebox.showinfo("Already Exists", "This formula is already in the list.")
            return
        self.formulas.append(formula)
        save_search_formulas(self.formulas)
        self.formulas_used.append(formula)
        self.formula_combo['values'] = self.formulas_used
        self.selected_formula.set(formula)
        self.formula_combo.set(formula)
        self.new_formula_entry.delete(0, tk.END)

    def update_leads_count_event(self, event=None):
        self.leads_text.edit_modified(0)
        self.update_leads_count()

    def update_leads_count(self):
        gmails = [g for g in extract_gmails(self.leads_text.get("1.0", tk.END)) if is_gmail(g)]
        self.leads_count_label.config(text=f"Count: {len(gmails)}")

    def check_real_gmails(self):
        raw = self.leads_text.get("1.0", tk.END)
        gmails = re.findall(r'\b[a-zA-Z0-9._%+-]+@gmail\.com\b', raw)
        valid = []
        seen = set()
        for g in gmails:
            g2 = g.strip().lower()
            if is_gmail(g2) and g2 not in seen:
                valid.append(g2)
                seen.add(g2)
        valid.sort()
        self.leads_text.delete("1.0", tk.END)
        self.leads_text.insert("1.0", ", ".join(valid))
        self.update_leads_count()
        self.log(f"Checked and cleaned leads: {len(valid)} valid gmails.")

    def refresh_leads_section(self):
        self.leads_text.delete("1.0", tk.END)
        self.update_leads_count()
        self.site_entry.set("instagram.com")
        self.niche_entry.set("Photographer")
        self.location_entry.delete(0, tk.END)
        self.location_entry.insert(0, "Los Angeles")
        self.quantity_entry.delete(0, tk.END)
        self.quantity_entry.insert(0, "20")
        self.selected_formula.set(self.formulas_used[0])
        self.formula_combo['values'] = self.formulas_used
        self.formula_combo.set(self.formulas_used[0])
        self.log("Gmail finder section refreshed.")

    def abort_all(self):
        self.abort_event.set()
        self.cta_mgr.error("All automation stopped by user.")

    def factory_reset(self):
        if not messagebox.askyesno("Factory Reset", "Are you sure you want to delete ALL tool data and reset app?\n\nThis will delete all credentials, tokens, profiles, attachments, usage data, formulas, templates, and mapping!\n\nContinue?"):
            return
        for folder in [CREDENTIALS_DIR, TOKENS_DIR, PROFILES_DIR, ATTACHMENTS_DIR]:
            try:
                shutil.rmtree(folder)
                os.makedirs(folder, exist_ok=True)
            except:
                pass
        for file in [USAGE_FILE, FORMULAS_FILE, TEMPLATE_FILE]:
            try:
                if os.path.exists(file):
                    os.remove(file)
            except:
                pass
        for d in [CREDENTIALS_DIR, TOKENS_DIR, PROFILES_DIR, ATTACHMENTS_DIR]:
            os.makedirs(d, exist_ok=True)
        self.email_entry.delete(0, tk.END)
        self.email_list_text.delete("1.0", tk.END)
        self.gmails_list = []
        self.accounts_listbox.delete(0, tk.END)
        self.message_text.delete("1.0", tk.END)
        self.attachment_path = None
        self.attachment_label.config(text="None")
        self.automation_stats.config(text="Sent: 0 | Skipped: 0 | Errors: 0")
        self.leads_text.delete("1.0", tk.END)
        self.site_entry.set("instagram.com")
        self.niche_entry.set("Photographer")
        self.location_entry.delete(0, tk.END)
        self.location_entry.insert(0, "Los Angeles")
        self.quantity_entry.delete(0, tk.END)
        self.quantity_entry.insert(0, "20")
        self.stats_box.delete("1.0", tk.END)
        self.status_var.set("Ready.")
        self.progress.set(0)
        self.cta_mgr.reset()
        self.refresh_api_accounts()
        self.refresh_accounts_listbox()
        self.formulas = load_search_formulas()
        if self.default_formula not in self.formulas:
            self.formulas.insert(0, self.default_formula)
        self.formulas_used = list(self.formulas)
        self.selected_formula.set(self.formulas_used[0])
        self.formula_combo['values'] = self.formulas_used
        self.formula_combo.set(self.formulas_used[0])
        self.abort_event.clear()
        self.update_leads_count()
        # Template reset
        self.templates = load_templates()
        self.selected_template_var.set(self.templates[0]["name"] if self.templates else "Default")
        self.template_combo['values'] = [t["name"] for t in self.templates]
        self.log("FACTORY RESET: All data and settings deleted. Tool is now in default state.", "INFO")

    def log(self, msg, msg_type="INFO"):
        now = datetime.now().strftime("%H:%M:%S")
        self.console_panel.insert(tk.END, f"{now} [{msg_type}] {msg}\n")
        self.console_panel.see(tk.END)
        self.console_panel.update()

    def set_status(self, msg, prog=0):
        self.status_var.set(msg)
        self.progress.set(prog)
        self.root.update_idletasks()

    def upload_credentials(self):
        path = filedialog.askopenfilename(filetypes=[('JSON Files', '*.json')])
        if not path: return
        dest = os.path.join(CREDENTIALS_DIR, os.path.basename(path))
        shutil.copy2(path, dest)
        self.log(f"credentials.json uploaded to {dest}")
        self.set_status("credentials.json uploaded.", 5)

    def add_gmail_api_account(self):
        if "google" in " ".join(missing_deps):
            self.log("Google API libraries missing.", "ERROR")
            self.set_status("Missing dependencies!", 0)
            self.cta_mgr.error("Google API missing!")
            return
        creds_file = filedialog.askopenfilename(initialdir=CREDENTIALS_DIR, filetypes=[('JSON Files', '*.json')])
        if not creds_file: return
        def run():
            try:
                self.cta_mgr.start("Authorizing Gmail API", os.path.basename(creds_file))
                creds = gmail_authenticate(creds_file, None)
                from googleapiclient.discovery import build
                service = build('gmail', 'v1', credentials=creds)
                profile = service.users().getProfile(userId="me").execute()
                email = profile["emailAddress"]
                token_file = os.path.join(TOKENS_DIR, f"{email}.pickle")
                save_pickle(token_file, creds)
                update_tokens_map(email, creds_file)
                self.log(f"Gmail API added for {email}")
                self.set_status(f"Gmail API authorized: {email}", 10)
                self.refresh_api_accounts()
                self.cta_mgr.stop("API Added")
            except Exception as e:
                self.log(str(e), "ERROR")
                self.set_status("Gmail API auth failed.", 0)
                self.cta_mgr.error(f"Failed: {e}")
        threading.Thread(target=run, daemon=True).start()

    def refresh_api_accounts(self):
        self.api_accounts_box.delete(*self.api_accounts_box.get_children())
        for token_file in os.listdir(TOKENS_DIR):
            if token_file.endswith(".pickle"):
                token_path = os.path.join(TOKENS_DIR, token_file)
                email, quota, used = get_gmail_stats(token_path)
                self.api_accounts_box.insert("", "end", values=(email, used, quota))

    def add_gmail_profile(self):
        profiles = get_chrome_profiles()
        profile_dir = None

        def pick_profile_dialog():
            dlg = Toplevel(self.root)
            dlg.title("Select Browser Profile (Auto/Manual)")
            dlg.geometry("440x330")
            ttk.Label(dlg, text="Select a detected browser profile:").pack(pady=(10, 2))
            profile_list = tk.Listbox(dlg, width=60, height=10)
            for name, path in profiles:
                profile_list.insert(tk.END, f"{name}: {path}")
            profile_list.pack(pady=(3, 5))
            ttk.Label(dlg, text="Or click 'Browse' to select any folder.").pack(pady=(2, 2))

            selected_profile = []
            def select_profile():
                idx = profile_list.curselection()
                if idx:
                    selected_profile.append(profiles[idx[0]][1])
                    dlg.destroy()
            def browse_profile():
                folder = filedialog.askdirectory(title="Select a browser profile folder")
                if folder:
                    selected_profile.append(folder)
                    dlg.destroy()
            ttk.Button(dlg, text="Select", command=select_profile).pack(side="left", padx=(40,10), pady=6)
            ttk.Button(dlg, text="Browse", command=browse_profile).pack(side="left", padx=10, pady=6)
            ttk.Button(dlg, text="Cancel", command=dlg.destroy).pack(side="right", padx=(10,30), pady=6)
            dlg.grab_set()
            self.root.wait_window(dlg)
            return selected_profile[0] if selected_profile else None

        profile_dir = pick_profile_dialog()
        if not profile_dir or not os.path.isdir(profile_dir):
            messagebox.showerror("Error", "No valid profile directory selected.")
            return

        profiles_map_path = os.path.join(PROFILES_DIR, "profiles_map.json")
        profiles_map = {}
        if os.path.exists(profiles_map_path):
            with open(profiles_map_path, "r", encoding="utf-8") as f:
                profiles_map = json.load(f)
        profiles_map["last_selected_profile"] = profile_dir
        with open(profiles_map_path, "w", encoding="utf-8") as f:
            json.dump(profiles_map, f, indent=2)
        self.cta_mgr.reset()
        threading.Thread(target=lambda: playwright_launch_gmail(profile_dir, self.log, self.cta_mgr)).start()
        self.log(f"Launching browser for login using selected profile: {profile_dir}")

    def open_gmail_profile(self):
        profiles_map_path = os.path.join(PROFILES_DIR, "profiles_map.json")
        profiles_map = {}
        if os.path.exists(profiles_map_path):
            with open(profiles_map_path, "r", encoding="utf-8") as f:
                profiles_map = json.load(f)
        profile_dir = profiles_map.get("last_selected_profile", "")
        if not profile_dir or not os.path.isdir(profile_dir):
            messagebox.showerror("Error", "No saved profile found. Please add one first.")
            return
        self.cta_mgr.reset()
        threading.Thread(target=lambda: playwright_open_profile(profile_dir, self.log, self.cta_mgr)).start()
        self.log(f"Launching browser with saved profile: {profile_dir}")

    def add_email_list(self):
        emails = [e.strip() for e in self.email_entry.get().split(",") if is_gmail(e.strip())]
        if not emails:
            messagebox.showerror("Error", "No valid Gmail addresses found.")
            return
        self.gmails_list = emails
        self.email_list_text.delete("1.0", tk.END)
        self.email_list_text.insert(tk.END, ", ".join(emails))
        self.log(f"Added {len(emails)} emails to list.")

    def refresh_accounts_listbox(self):
        self.accounts_listbox.delete(0, tk.END)
        for token_file in os.listdir(TOKENS_DIR):
            if token_file.endswith(".pickle"):
                token_path = os.path.join(TOKENS_DIR, token_file)
                email, quota, used = get_gmail_stats(token_path)
                if used < quota:
                    self.accounts_listbox.insert(tk.END, f"{email} ({used}/{quota})")

    def browse_attachment(self):
        path = filedialog.askopenfilename(filetypes=[('All Files', '*.*')])
        if path:
            self.attachment_path = path
            self.attachment_label.config(text=os.path.basename(path))
            self.log(f"Attachment selected: {path}")
        else:
            self.attachment_label.config(text="None")

    def open_template_dialog(self):
        dlg = Toplevel(self.root)
        dlg.title("Add/Edit HTML Template")
        dlg.geometry("730x540")
        ttk.Label(dlg, text="Template Name:").pack(pady=(10,2))
        name_entry = ttk.Entry(dlg, width=40)
        name_entry.pack()
        ttk.Label(dlg, text="Paste HTML code below:").pack(pady=(10,2))
        html_text = ScrolledText(dlg, width=85, height=22)
        html_text.pack(expand=True, fill=tk.BOTH, padx=8, pady=4)
        # Pre-fill if editing
        selected_name = self.selected_template_var.get()
        current_html = ""
        if selected_name:
            for t in self.templates:
                if t["name"] == selected_name:
                    name_entry.insert(0, t["name"])
                    html_text.insert("1.0", t["html"])
                    current_html = t["html"]
                    break

        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(pady=10)
        def save_template():
            name = name_entry.get().strip()
            if not name:
                messagebox.showerror("Error", "Template name required.")
                return
            if name == "Default":
                messagebox.showerror("Error", "You cannot overwrite the Default template.")
                return
            html = html_text.get("1.0", tk.END).strip()
            if not html:
                messagebox.showerror("Error", "HTML code required.")
                return
            add_or_update_template(name, html)
            self.templates = load_templates()
            self.template_combo['values'] = [t["name"] for t in self.templates]
            self.selected_template_var.set(name)
            self.template_combo.set(name)
            messagebox.showinfo("Saved", f"Template '{name}' saved.")
            dlg.destroy()
        ttk.Button(btn_frame, text="Save", command=save_template).pack(side="left", padx=14)
        ttk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(side="left", padx=14)
        dlg.grab_set()
        self.root.wait_window(dlg)

    def start_automation(self):
        if not build or not sync_playwright:
            self.log("Required dependencies missing for automation.", "ERROR")
            if not build and googleapi_error:
                self.log(f"Google API ImportError: {googleapi_error}", "ERROR")
            if not sync_playwright and playwright_error:
                self.log(f"Playwright ImportError: {playwright_error}", "ERROR")
            self.set_status("Missing dependencies!", 0)
            self.cta_mgr.error("Dependencies missing!")
            return
        selected = [self.accounts_listbox.get(i) for i in self.accounts_listbox.curselection()]
        if not selected:
            messagebox.showerror("Error", "Select at least one Gmail API account.")
            return
        if not self.gmails_list:
            messagebox.showerror("Error", "No emails in the list.")
            return
        msg = self.message_text.get("1.0", tk.END).strip()
        if not msg and not self.attachment_path and not self.selected_template_var.get():
            messagebox.showerror("Error", "Message is empty, no attachment and no template selected.")
            return
        attachment = self.attachment_path
        self.cta_mgr.reset()
        threading.Thread(target=self.gmail_campaign_thread, args=(selected, msg, attachment), daemon=True).start()

    def gmail_campaign_thread(self, selected_accounts, msg, attachment):
        sent = skipped = errors = 0
        total = len(self.gmails_list) * len(selected_accounts)
        self.set_status("Starting Gmail campaign...", 5)
        self.start_time = time.time()
        self.task_running = True
        self.cta_mgr.start("Starting Gmail campaign")
        try:
            for idx_account, account in enumerate(selected_accounts):
                email = account.split()[0]
                creds_file = get_creds_path_from_email(email)
                token_file = os.path.join(TOKENS_DIR, f"{email}.pickle")
                if not creds_file or not os.path.exists(token_file):
                    self.log(f"No credentials/token mapping for {email}!", "ERROR")
                    errors += len(self.gmails_list)
                    continue
                try:
                    self.cta_mgr.set_action("Authenticating", email)
                    try:
                        creds = gmail_authenticate(creds_file, token_file)
                        from googleapiclient.discovery import build
                        service = build('gmail', 'v1', credentials=creds)
                    except Exception as e:
                        errors += len(self.gmails_list)
                        self.log(f"Dependency/Authentication error for {email}: {e}", "ERROR")
                        self.cta_mgr.error(f"Dependency/Authentication error: {e}")
                        self.update_automation_stats(sent, skipped, errors, total)
                        return
                    # Get template HTML (if selected)
                    template_html = ""
                    sel_template = self.selected_template_var.get()
                    if sel_template:
                        for t in self.templates:
                            if t["name"] == sel_template:
                                template_html = t["html"]
                                break
                    for idx_gmail, to in enumerate(self.gmails_list):
                        try:
                            self.cta_mgr.set_action("Sending Email", f"Account {idx_account+1}/{len(selected_accounts)}, Email {idx_gmail+1}/{len(self.gmails_list)}")
                            self.show_timer(sent, skipped, errors, total, idx_account, len(selected_accounts), idx_gmail, len(self.gmails_list), "Sending email")
                            if template_html:
                                gmail_send_message(service, email, to, "Campaign Message", msg, attachment, html_message=template_html)
                            else:
                                gmail_send_message(service, email, to, "Campaign Message", msg, attachment)
                            increment_usage(email)
                            sent += 1
                            self.log(f"Sent to {to} from {email}")
                        except Exception as e:
                            skipped += 1
                            self.log(f"Skipped {to}: {e}", "ERROR")
                            self.cta_mgr.set_action("Send failed", f"{to}: {e}")
                        self.update_automation_stats(sent, skipped, errors, total)
                except Exception as e:
                    self.log(f"Failed to send from {email}: {e}", "ERROR")
                    errors += len(self.gmails_list)
                    self.cta_mgr.set_action("Error Sending", str(e))
                self.update_automation_stats(sent, skipped, errors, total)
            if errors > 0:
                self.cta_mgr.error("Automation finished with errors.")
                self.set_status("Automation finished with errors.", 100)
            else:
                self.log("Gmail campaign finished.")
                self.set_status("Automation Complete", 100)
                self.cta_mgr.stop("Gmail campaign finished.")
        except Exception as e:
            self.log(f"Automation error: {e}", "ERROR")
            self.cta_mgr.error(f"Automation failed: {e}")
        self.task_running = False

    def update_automation_stats(self, sent, skipped, errors, total):
        self.automation_stats.config(text=f"Sent: {sent} | Skipped: {skipped} | Errors: {errors}")
        pct = int(100 * (sent + skipped + errors) / (total if total else 1))
        self.progress.set(pct)

    def show_timer(self, sent, skipped, errors, total, idx_account, n_accounts, idx_gmail, n_gmails, action):
        now = time.time()
        elapsed = now - self.start_time if self.start_time else 0
        done = sent + skipped + errors
        left = total - done
        eta = "--:--"
        if done > 0 and left > 0:
            per_item = elapsed / done
            eta_sec = int(left * per_item)
            eta = f"{eta_sec//60:02d}:{eta_sec%60:02d}"
        status_msg = f"[{action}] {sent+skipped+errors}/{total} | Current: Account {idx_account+1}/{n_accounts}, Email {idx_gmail+1}/{n_gmails} | ETA: {eta}"
        self.log(status_msg, "INFO")
        self.set_status(status_msg, self.progress.get())

    def find_gmails_leads(self):
        if not sync_playwright:
            self.log("Playwright is not installed! Install with: pip install playwright playwright-stealth", "ERROR")
            self.set_status("Missing dependencies!", 0)
            self.cta_mgr.error("Playwright missing!")
            return
        site = self.site_entry.get().strip()
        niche = self.niche_entry.get().strip()
        location = self.location_entry.get().strip()
        try:
            desired_quantity = int(self.quantity_entry.get().strip())
        except:
            desired_quantity = 20
        search_formula = self.selected_formula.get().strip()
        if search_formula not in self.formulas_used:
            self.formulas_used.append(search_formula)
            self.formula_combo['values'] = self.formulas_used
        self.abort_event.clear()
        self.set_status("Starting Gmail scraping...", 10)
        self.start_time = time.time()
        self.task_running = True
        self.cta_mgr.reset()
        self.cta_mgr.start("Scraping Gmail (auto-loop)")

        existing_gmails = set(
            g.strip().lower() for g in extract_gmails(self.leads_text.get("1.0", tk.END)) if is_gmail(g.strip())
        )
        all_gmails = set(existing_gmails)
        attempt = 0

        def update_fn(progress, total):
            self.set_status(f"[Finding] {progress}/{total}", int(progress/total*100) if total else 0)
            self.cta_mgr.set_action("Extracting Gmail leads", f"{progress}/{total}")

        def run():
            nonlocal all_gmails, attempt
            while len(all_gmails) < desired_quantity and not self.abort_event.is_set():
                attempt += 1
                self.log(f"--- Scraping Attempt #{attempt} ---", "INFO")
                gmails = playwright_gmail_leads(
                    site, niche, location, desired_quantity, self.log, update_fn, self.cta_mgr, self.root,
                    formula=search_formula, abort_event=self.abort_event, two_min_captcha=True
                )
                prev_count = len(all_gmails)
                all_gmails.update(g.strip().lower() for g in gmails if is_gmail(g.strip()))
                self.leads_text.delete("1.0", tk.END)
                self.leads_text.insert("1.0", "\n".join(sorted(all_gmails)))
                self.update_leads_count()
                self.log(f"Total unique gmails so far: {len(all_gmails)}", "INFO")
                if len(all_gmails) == prev_count or self.abort_event.is_set():
                    self.log("No new gmails found in this attempt or stopped by user. Stopping loop.", "INFO")
                    break
                if len(all_gmails) >= desired_quantity:
                    break
                time.sleep(random.randint(7, 17))
            if self.abort_event.is_set():
                self.cta_mgr.error("Stopped by user/browser close")
                self.set_status("Finder stopped by user or browser close.", 100)
            else:
                self.cta_mgr.stop(f"Gmail finding loop done. Found {len(all_gmails)}")
                self.set_status(f"Finder done. Total {len(all_gmails)} Gmail(s)", 100)
            self.task_running = False

        threading.Thread(target=run, daemon=True).start()

    def copy_leads(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.leads_text.get("1.0", tk.END).strip())
        self.set_status("Leads copied to clipboard.", 0)

    def refresh_stats(self):
        stats = []
        for token_file in os.listdir(TOKENS_DIR):
            if token_file.endswith(".pickle"):
                token_path = os.path.join(TOKENS_DIR, token_file)
                email, quota, used = get_gmail_stats(token_path)
                stats.append(f"{email}  {used}/{quota} used today")
        self.stats_box.delete("1.0", tk.END)
        self.stats_box.insert(tk.END, "\n".join(stats) if stats else "No Gmail API accounts found.")

    def simple_input_dialog(self, title, prompt):
        d = Toplevel(self.root)
        d.title(title)
        ttk.Label(d, text=prompt).pack(pady=6)
        e = ttk.Entry(d, width=28)
        e.pack(padx=8)
        e.focus()
        val = []
        def ok():
            val.append(e.get())
            d.destroy()
        ttk.Button(d, text="OK", command=ok).pack(pady=3)
        d.grab_set()
        self.root.wait_window(d)
        return val[0] if val else None

if __name__ == "__main__":
    root = tk.Tk()
    app = GmailProApp(root)
    app.log(f"Gmail Automation Pro v{VERSION} started.")
    if missing_deps:
        app.log("Some dependencies are missing. See above for details.", "ERROR")
        app.cta_mgr.error("Dependencies missing!")
    root.mainloop()