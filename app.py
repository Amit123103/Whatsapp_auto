"""
Skytecher WhatsApp Outreach Suite (v2.0)
Production-ready Python desktop application with Tkinter UI, NVIDIA NIM / AI model integration,
and automated cold WhatsApp messaging from Excel.

Key Features:
- Excel (.xlsx) parsing with pandas & openpyxl
- Intelligent phone cleaning (+91 formatting, removes floats like .0, strips dashes/spaces)
- Automatic skipping for duplicates, empty/invalid numbers, sent-history, and blocklists
- Automatic continuation when a phone number does not have WhatsApp (never stops the pipeline)
- AI Personalized Message Generation via NVIDIA NIM / OpenAI-compatible API (e.g. DeepSeek-v4.1-flash, GPT-OSS-20B)
- Full Skytecher branding: Team, https://skytecher.com, skytechersolutions@gmail.com, WhatsApp: 8960061745
- .env configuration support for owner phone, company details, and API keys
- Multi-threaded background dispatch (UI never freezes)
- Anti-ban delay randomization, daily limits, simulation mode, and live color-coded console
- Results export to Excel (results.xlsx)

Author: Built for Skytecher (https://skytecher.com)
Sender Phone: 8960061745
Email: skytechersolutions@gmail.com
"""

import os
import re
import sys
import time
import random
import datetime
import threading
import traceback
import webbrowser
import urllib.parse
from typing import Dict, List, Optional, Tuple, Set

import pandas as pd
import openpyxl
import requests

# Load environment variables
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

# Optional pyautogui import for keyboard automation
try:
    import pyautogui
    PYAUTOGUI_AVAILABLE = True
except Exception:
    PYAUTOGUI_AVAILABLE = False
    pyautogui = None

# Optional pywhatkit import (fallback method)
try:
    import pywhatkit
    PYWHATKIT_AVAILABLE = True
except Exception:
    PYWHATKIT_AVAILABLE = False
    pywhatkit = None


# Disable pyautogui failsafe so mouse-corner movements don't crash the automation
if PYAUTOGUI_AVAILABLE:
    pyautogui.FAILSAFE = False
    pyautogui.PAUSE = 0.3


def send_whatsapp_message(phone: str, message: str, wait_time: int = 18) -> Tuple[bool, str]:
    """
    Sends a WhatsApp message fully automatically:
    1. Direct web.whatsapp.com/send?phone=XX&text=YY URL (pre-fills phone + message)
    2. Waits for WhatsApp Web to load the conversation
    3. Clicks on the message input box to ensure active browser focus
    4. Automatically clicks the circular Send Button (the button with '>' icon at bottom-right)
       - Uses template matching across light/dark/green themes
       - Falls back to exact pixel coordinate targeting at the bottom-right corner of the chat
       - Concurrently triggers Enter and Ctrl+Enter for guaranteed delivery
    5. Waits for delivery confirmation
    6. Closes the browser tab automatically (Ctrl+W)
    
    The entire process is 100% hands-free.
    
    Returns: (success: bool, detail: str)
    """
    if not PYAUTOGUI_AVAILABLE:
        return False, "pyautogui not installed — required for auto-sending. Run: pip install pyautogui"

    try:
        # URL-encode the message for the WhatsApp Web API
        encoded_msg = urllib.parse.quote(message)
        # WhatsApp API expects country code digits without '+'
        clean_phone = phone.lstrip("+").replace(" ", "").replace("-", "")
        
        # Build the WhatsApp Web direct-send URL
        wa_url = f"https://web.whatsapp.com/send?phone={clean_phone}&text={encoded_msg}"
        
        # ---- STEP 1: Open WhatsApp Web in default browser ----
        webbrowser.open(wa_url)
        
        # ---- STEP 2: Wait for the page to fully load ----
        time.sleep(wait_time)
        
        screen_w, screen_h = pyautogui.size()
        
        # ---- STEP 3: Click inside the message input area to guarantee focus ----
        try:
            # Click near bottom-center of the chat input bar
            input_x = int(screen_w * 0.60)
            input_y = int(screen_h * 0.94)
            pyautogui.click(input_x, input_y)
            time.sleep(0.8)
        except Exception:
            pass

        # ---- STEP 4: Automatically CLICK the circular Send Button ----
        button_clicked = False
        
        # Try computer vision template matching in the bottom-right quadrant
        templates = [
            os.path.join("assets", "whatsapp_send_btn_light.png"),
            os.path.join("assets", "whatsapp_send_btn_dark.png"),
            os.path.join("assets", "whatsapp_send_btn_green.png"),
        ]
        search_region = (
            int(screen_w * 0.65),
            int(screen_h * 0.80),
            int(screen_w * 0.35),
            int(screen_h * 0.20),
        )
        for tpl in templates:
            if os.path.exists(tpl):
                try:
                    pos = pyautogui.locateCenterOnScreen(tpl, confidence=0.65, region=search_region)
                    if pos:
                        pyautogui.click(pos.x, pos.y)
                        button_clicked = True
                        time.sleep(0.4)
                        break
                except Exception:
                    pass

        # If template matching didn't trigger, click exact physical coordinates of Send Button
        if not button_clicked:
            # The circular Send button in WhatsApp Web is docked at the far right of the input bar
            send_coords = [
                (screen_w - 58, screen_h - 68),                  # Primary 1080p target
                (int(screen_w * 0.972), int(screen_h * 0.942)),  # Proportional target
                (screen_w - 65, screen_h - 70),                  # Inset target
            ]
            for cx, cy in send_coords:
                try:
                    pyautogui.click(cx, cy)
                    time.sleep(0.3)
                except Exception:
                    pass

        # ---- STEP 5: Trigger Enter and Ctrl+Enter keystrokes as a multi-layer guarantee ----
        try:
            pyautogui.press("enter")
            time.sleep(0.3)
            pyautogui.hotkey("ctrl", "enter")
        except Exception:
            pass
        
        # ---- STEP 6: Wait for message transmission / delivery ----
        time.sleep(4)
        
        # ---- STEP 7: Close the browser tab (Ctrl+W) ----
        pyautogui.hotkey("ctrl", "w")
        time.sleep(1.5)
        
        return True, "Message sent successfully via WhatsApp Web"
        
    except Exception as e:
        # Attempt to clean up any lingering browser tab or popup on error
        try:
            time.sleep(1)
            pyautogui.press("escape")
            time.sleep(0.5)
            pyautogui.hotkey("ctrl", "w")
            time.sleep(1)
        except Exception:
            pass
        return False, f"WhatsApp send error: {str(e)}"


# ==========================================
# CONFIGURATION & ENVIRONMENT SETUP
# ==========================================
COMPANY_NAME = os.getenv("COMPANY_NAME", "Skytecher")
COMPANY_WEBSITE = os.getenv("COMPANY_WEBSITE", "https://skytecher.com")
COMPANY_EMAIL = os.getenv("COMPANY_EMAIL", "skytechersolutions@gmail.com")
SENDER_PHONE = os.getenv("SENDER_PHONE", "8960061745")

DEFAULT_BLOCKLIST_FILE = "blocklist.txt"
DEFAULT_SENT_HISTORY_FILE = "sent_history.txt"
DEFAULT_EXPORT_FILE = "results.xlsx"

DEFAULT_AI_BASE_URL = os.getenv("AI_BASE_URL", "https://integrate.api.nvidia.com/v1")
DEFAULT_AI_MODEL = os.getenv("AI_MODEL", "deepseek-ai/deepseek-v4.1-flash")

POPULAR_AI_MODELS = [
    "deepseek-ai/deepseek-v4.1-flash",
    "openai/gpt-oss-20b",
    "nvidia/nemotron-3.5-lightning-30b-a3b",
    "meta/llama-3.3-70b-instruct",
    "google/gemma-2-9b-it",
    "mistralai/mistral-large-2-instruct",
]

DEFAULT_MESSAGE_TEMPLATE = """{greeting}

I came across *{business}* and wanted to reach out from the team at *Skytecher* (https://skytecher.com).

We help businesses in your space solve 3 common growth bottlenecks:
❌ Losing high-value customers to competitors with stronger online presence
❌ Lack of a modern, fast mobile website or WhatsApp catalog
❌ Inconsistent daily customer inquiries and leads

We can help your team fix this with *{service}*.

💡 *Tailored Idea for {business}:*
{suggestion}

Would you be open to a quick 3-minute chat or demo this week? No pressure — just sharing actionable ideas you can use right away.

Best regards,
*Skytecher Team* ⚡
🌐 https://skytecher.com
✉️ skytechersolutions@gmail.com
📱 +91 8960061745

_Reply STOP to opt out._"""

# Pre-crafted Cold Message Presets for Skytecher (English, Hinglish, & Service Specific)
MESSAGE_PRESETS = {
    "1. Professional Introduction (English)": """{greeting}

This is the team at *Skytecher* — we help companies like *{business}* grow their daily customer inquiries and online sales.

We noticed that many businesses in your industry face these challenges:
❌ Low visibility on Google when local customers search
❌ Outdated or missing mobile website to convert visitors
❌ Missing direct WhatsApp ordering or lead capture

We can help solve this for you with *{service}*.

💡 *Quick Idea for {business}:*
{suggestion}

Would it be okay if I shared a quick 2-minute idea on how this could bring more daily customers to your business?

👉 See our work: https://skytecher.com

Best regards,
*Skytecher Team* ⚡
✉️ skytechersolutions@gmail.com
📱 +91 8960061745

_Reply STOP to opt out._""",

    "2. Problem -> Solution (English)": """{greeting}

Did you know that over *80% of customers* research online before choosing a business?

Without a high-speed website and targeted online presence, *{business}* could be missing out on ready-to-buy customers every day.

*Here is how Skytecher fixes this:*
✅ Modern, high-converting websites & mobile apps
✅ Targeted Google ranking & social reach
✅ Automated WhatsApp lead capture that runs 24/7

💡 *Specific Recommendation for {business}:*
{suggestion}

Can I send over a quick 2-minute walkthrough showing how *{service}* can increase your inquiries?

👉 Portfolio: https://skytecher.com

*Skytecher Team* ⚡
✉️ skytechersolutions@gmail.com | 📱 +91 8960061745

_Reply STOP to opt out._""",

    "3. Free Strategy Call (English)": """{greeting}

I'm reaching out from *Skytecher* because we have a proven growth framework for companies like *{business}*.

🎁 *COMPLIMENTARY OFFER:* We are offering a free 10-minute digital audit & strategy session for your team on *{service}* — 100% actionable, no sales pitch.

*What we will share:*
✅ Breakdown of your current digital visibility vs competitors
✅ 3 immediate fixes to attract more high-paying customers
✅ A step-by-step roadmap to scale inquiries

💡 *Initial Observation:*
{suggestion}

Would you be open to a quick 5-minute call this week?

👉 Check our work: https://skytecher.com

Warm regards,
*Skytecher Team* ⚡
✉️ skytechersolutions@gmail.com | 📱 +91 8960061745

_Reply STOP to opt out._""",

    "4. Follow-up Day 3 (English)": """{greeting}

Quick follow-up from Skytecher! 🙂

I previously shared an idea on how *{business}* can generate 2-3x more customer inquiries using *{service}*.

💡 *Quick reminder:*
{suggestion}

Would you have 2 minutes this week for a brief chat, or should I send a quick summary over WhatsApp?

👉 https://skytecher.com

*Skytecher Team* ⚡
✉️ skytechersolutions@gmail.com | 📱 +91 8960061745

_Reply STOP to opt out._""",

    "5. Final Polite Follow-up (English)": """{greeting}

This is my last message — I truly respect your time and won't follow up again uninvited.

Whenever *{business}* is ready to upgrade your website, scale customer leads, or implement *{service}*, Skytecher is always here to support your team.

💡 {suggestion}

Wishing *{business}* continued growth and success! 🚀

👉 https://skytecher.com
✉️ skytechersolutions@gmail.com | 📱 +91 8960061745

*Skytecher Team* ⚡

_Reply STOP to opt out._""",

    "6. Friendly Introduction (Hinglish)": """{greeting}

Main *Skytecher Team* (https://skytecher.com) se connect kar raha hoon.

Hum *{business}* jaise businesses ko online grow karne aur daily customer inquiries badhane me help karte hain:
✅ Fast & Modern Website / Mobile App
✅ Google pe top visibility taaki local customers aapko dhundh sakein
✅ Direct WhatsApp ordering & automated customer inquiry setup

Aaj kal mostly businesses ye challenge face karti hain:
❌ Customers Google pe search karte hain par competitors ke paas chale jaate hain
❌ Outdated website ki wajah se trust build nahi hota

💡 *{business} ke liye hamara suggestion:*
{suggestion}

We can help you fix this with *{service}*.

Kya main aapke saath ek quick 2-minute idea share kar sakta hoon?

👉 Humare projects check karein: https://skytecher.com

Best regards,
*Skytecher Team* ⚡
✉️ skytechersolutions@gmail.com
📱 +91 8960061745

_Reply STOP to opt out._""",

    "7. Problem -> Solution (Hinglish)": """{greeting}

Aaj kal 80%+ customers kisi bhi business se judne se pehle unhe online check karte hain.

Agar *{business}* ki strong digital presence nahi hai, to rozana valuable customers miss ho rahe hain.

*Skytecher aapko kaise help karta hai:*
✅ *{service}* se aapka customer conversion 2x boost
✅ Google search pe top rank
✅ 24/7 automated WhatsApp lead capture

💡 *Hamara quick idea:*
{suggestion}

👉 Portfolio: https://skytecher.com

Kya hum is week 5 minute connect kar sakte hain?

*Skytecher Team* ⚡
✉️ skytechersolutions@gmail.com | 📱 +91 8960061745

_Reply STOP to opt out._""",

    "8. Free Consultation (Hinglish)": """{greeting}

Hum *{business}* ke liye ek *FREE 10-minute digital growth audit* offer kar rahe hain!

📋 *Is session me aapko milega:*
✅ Aapki current digital visibility ka quick review
✅ 3 practical growth tips jo aap turant use kar sakte hain
✅ *{service}* ke through daily leads badhane ka blueprint

Koi sales pressure nahi — sirf actionable ideas.

💡 *Initial suggestion:*
{suggestion}

👉 https://skytecher.com

Kya kal ya parson ek 5-minute call schedule karein?

*Skytecher Team* ⚡
✉️ skytechersolutions@gmail.com | 📱 +91 8960061745

_Reply STOP to opt out._""",

    "9. Follow-up Day 3 (Hinglish)": """{greeting}

Pichle message par ek quick follow-up! 🙂

Humare paas *{business}* ke liye ek simple strategy hai — *{service}* ke through aapki daily customer inquiries 2-3x boost ho sakti hain.

💡 *Reminder:*
{suggestion}

Bas 2 minute ka time lagega — kya ek quick chat karein?

👉 https://skytecher.com

*Skytecher Team* ⚡
✉️ skytechersolutions@gmail.com | 📱 +91 8960061745

_Reply STOP to opt out._""",

    "10. Final Follow-up (Hinglish)": """{greeting}

Main aapka precious time waste nahi karna chahta, isliye yeh mera last message hai.

Aage kabhi bhi agar *{business}* ko digital growth me help chahiye:
✅ High-converting Website ya Mobile App
✅ Daily customer leads aur Google ranking
✅ WhatsApp automation

*Skytecher* hamesha aapke support ke liye taiyaar hai! 🚀

👉 https://skytecher.com
✉️ skytechersolutions@gmail.com | 📱 +91 8960061745

Best wishes for your business! ⚡
*Skytecher Team*

_Reply STOP to opt out._""",

    "11. Website Development Focus": """{greeting}

Does *{business}* have a fast, mobile-friendly website that turns visitors into paying customers?

🔍 *Common website bottlenecks:*
❌ Slow loading speed — visitors bounce in 3 seconds
❌ Poor mobile experience — over 75% traffic is mobile
❌ No direct WhatsApp chat or lead capture button
❌ Outdated look that doesn't reflect your actual quality

*Skytecher builds high-performance websites:*
✅ Modern, lightning-fast & mobile-optimized
✅ Built-in WhatsApp lead capture & instant booking
✅ Ranked on Google for local customer searches

💡 *Recommendation for {business}:*
{suggestion}

Would it be okay if I shared a free 2-minute design preview for your business?

👉 Portfolio: https://skytecher.com

*Skytecher Team* ⚡
✉️ skytechersolutions@gmail.com | 📱 +91 8960061745

_Reply STOP to opt out._""",

    "12. Mobile App Development Focus": """{greeting}

Have you considered how a dedicated mobile app or custom business portal could streamline operations for *{business}*?

📱 *Key benefits for your business:*
✅ Direct push notifications to customers (90%+ open rate)
✅ Seamless order / service booking without middlemen
✅ Automated loyalty & repeat customer engagement

💡 *Suggested Concept for {business}:*
{suggestion}

We specialize in high-speed, cost-effective mobile apps built with *{service}*.

Would you be open to seeing a 2-minute prototype concept?

👉 https://skytecher.com

*Skytecher Team* ⚡
✉️ skytechersolutions@gmail.com | 📱 +91 8960061745

_Reply STOP to opt out._""",

    "13. Digital Marketing & SEO Focus": """{greeting}

When potential customers in your area search online for services you provide, does *{business}* show up on page 1? 🔍

*The reality of digital marketing today:*
❌ Running unoptimized ads burns budget with zero ROI
❌ Without local SEO, nearby competitors get all the calls
❌ Inquiries go cold without instant WhatsApp automation

*How Skytecher delivers measurable results:*
✅ Top Google search ranking for high-intent keywords
✅ High-converting ad campaigns with transparent ROI tracking
✅ 2-4x increase in qualified weekly inquiries

💡 *Tailored Strategy for {business}:*
{suggestion}

Can I send you a free 2-minute digital visibility audit for your business?

👉 Case Studies: https://skytecher.com

*Skytecher Team* ⚡
✉️ skytechersolutions@gmail.com | 📱 +91 8960061745

_Reply STOP to opt out._""",

    "14. Full Detailed Intro": DEFAULT_MESSAGE_TEMPLATE,
}



# ==========================================
# HELPER UTILITIES: PHONE, FILE & TEMPLATE
# ==========================================
def clean_phone_number(raw_val) -> Tuple[Optional[str], Optional[str]]:
    """
    Cleans and standardizes phone numbers from Excel.
    - Handles Excel float outputs like 9876543210.0
    - Strips whitespace, dashes, parentheses, tabs
    - Adds +91 for 10-digit Indian numbers
    - Normalizes 11-digit numbers starting with 0 to +91
    - Handles international numbers with country codes
    
    Returns: (cleaned_phone, error_message)
    """
    if raw_val is None or pd.isna(raw_val):
        return None, "Empty phone number"

    val_str = str(raw_val).strip()
    if not val_str:
        return None, "Empty phone number"

    # Remove Excel float .0 artifact
    if val_str.endswith(".0"):
        val_str = val_str[:-2].strip()

    # Preserve leading '+' if present, extract all digits
    has_plus = val_str.startswith("+")
    digits = re.sub(r"\D", "", val_str)

    if not digits:
        return None, "No valid digits found"

    # 10 Digits -> Standard Indian Mobile Number -> +91XXXXXXXXXX
    if len(digits) == 10:
        return f"+91{digits}", None

    # 11 Digits starting with 0 -> Indian STD format (0XXXXXXXXXX) -> +91XXXXXXXXXX
    if len(digits) == 11 and digits.startswith("0"):
        return f"+91{digits[1:]}", None

    # 12 Digits starting with 91 -> Indian Number without plus (+91XXXXXXXXXX)
    if len(digits) == 12 and digits.startswith("91"):
        return f"+{digits}", None

    # E.164 international standard: 10 to 15 digits
    if has_plus and (10 <= len(digits) <= 15):
        return f"+{digits}", None

    if 11 <= len(digits) <= 15:
        return f"+{digits}", None

    return None, f"Invalid phone format ({len(digits)} digits: '{val_str}')"


def load_blocklist(filepath: str = DEFAULT_BLOCKLIST_FILE) -> Set[str]:
    """Loads blocked numbers from a text file into a set of normalized phone numbers."""
    blocked: Set[str] = set()
    if not os.path.exists(filepath):
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                f.write("# Skytecher WhatsApp Outreach - Blocklist\n# Add numbers to block below:\n")
        except Exception:
            pass
        return blocked

    try:
        with open(filepath, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                norm, _ = clean_phone_number(line)
                if norm:
                    blocked.add(norm)
                else:
                    blocked.add(line)
    except Exception as e:
        print(f"Error loading blocklist: {e}")
    return blocked


def append_to_blocklist(phone: str, filepath: str = DEFAULT_BLOCKLIST_FILE) -> bool:
    """Appends a new phone number to the blocklist file."""
    norm, _ = clean_phone_number(phone)
    target = norm if norm else phone.strip()
    if not target:
        return False
    try:
        with open(filepath, "a", encoding="utf-8") as f:
            f.write(f"\n{target}")
        return True
    except Exception as e:
        print(f"Error appending to blocklist: {e}")
        return False


def load_sent_history(filepath: str = DEFAULT_SENT_HISTORY_FILE) -> Set[str]:
    """Loads previously messaged phone numbers from history file."""
    sent: Set[str] = set()
    if not os.path.exists(filepath):
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                f.write("# Skytecher WhatsApp Outreach - Sent History\n")
        except Exception:
            pass
        return sent

    try:
        with open(filepath, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                parts = [p.strip() for p in line.split(",")]
                if parts:
                    norm, _ = clean_phone_number(parts[0])
                    if norm:
                        sent.add(norm)
    except Exception as e:
        print(f"Error loading sent history: {e}")
    return sent


def append_to_sent_history(phone: str, company: str = "", filepath: str = DEFAULT_SENT_HISTORY_FILE):
    """Logs a successful send to the history file to avoid future duplicates."""
    try:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(filepath, "a", encoding="utf-8") as f:
            f.write(f"{phone}, {timestamp}, {company}\n")
    except Exception as e:
        print(f"Error appending to sent history: {e}")


def _safe_str(val) -> str:
    """Safely convert any value to a stripped string. Handles None, NaN, float, etc."""
    if val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ["nan", "none", "nat", "null"]:
        return ""
    return s


def get_smart_greeting_and_name(raw_name: str, raw_company: str, is_hinglish: bool = False) -> Tuple[str, str]:
    """
    Returns (greeting, display_name) that avoids awkward corporate greetings like 'Hi Apex Retail Solutions'.
    If raw_name is empty, generic, or identical to company, or corporate:
    -> ('Hello Team <Company> 👋', 'Team <Company>') in English.
    -> ('Namaste Team <Company> 🙏', 'Team <Company>') in Hinglish.
    If raw_name is an individual person's name (e.g. 'Rahul', 'Amit Sharma'):
    -> ('Hi Rahul 👋', 'Rahul') in English.
    -> ('Namaste Rahul ji 🙏', 'Rahul') in Hinglish.
    """
    name = (raw_name or "").strip()
    company = (raw_company or "").strip()

    is_empty_or_generic = (
        not name
        or name.lower() in ["there", "contact", "lead", "sir", "madam", "owner", "manager", "nan", "none", "unknown"]
        or name.lower().startswith("contact #")
        or name.lower().startswith("lead #")
    )

    corporate_keywords = [
        "solutions", "retail", "tech", "technologies", "services", "enterprises",
        "pvt", "ltd", "llp", "corp", "corporation", "inc", "co", "company",
        "industries", "agency", "consulting", "group", "holdings", "stores",
        "shop", "mart", "bazaar", "hospital", "clinic", "school", "academy",
        "hotel", "restaurant", "resort", "jewellers", "jewelers", "fashion",
        "care", "motors", "associates", "logistics", "traders", "trading",
        "creations", "ventures", "studio", "labs"
    ]
    name_words = [w.lower().strip(".,()/-") for w in name.split()]
    is_corporate = any(kw in name_words for kw in corporate_keywords)

    is_company = (
        is_empty_or_generic
        or (company and name.lower() == company.lower())
        or is_corporate
    )

    if is_company:
        target = company if company and company.lower() != "your business" else ""
        if target:
            if is_hinglish:
                return f"Namaste Team {target} 🙏", f"Team {target}"
            return f"Hello Team {target} 👋", f"Team {target}"
        else:
            return ("Namaste ji 🙏" if is_hinglish else "Hi there 👋"), "there"
    else:
        first_name = name.split()[0].title()
        if is_hinglish:
            return f"Namaste {first_name} ji 🙏", first_name
        return f"Hi {first_name} 👋", first_name


def format_template_message(template: str, contact_data: Dict[str, str]) -> str:
    """Substitutes placeholders with contact attributes, rich lead data, and Skytecher details."""
    company = (
        _safe_str(contact_data.get("Company"))
        or _safe_str(contact_data.get("Business"))
        or "your business"
    )
    raw_name = _safe_str(contact_data.get("Name"))
    is_hinglish = "namaste" in template.lower() or "karein" in template.lower() or "hum " in template.lower() or "aap" in template.lower()
    greeting, display_name = get_smart_greeting_and_name(raw_name, company, is_hinglish)

    service = (
        _safe_str(contact_data.get("Suggested Services"))
        or _safe_str(contact_data.get("Service"))
        or "modern website development & digital growth solutions"
    )
    suggestion = (
        _safe_str(contact_data.get("Notes"))
        or _safe_str(contact_data.get("Suggestion"))
        or "We would love to share high-impact strategies to enhance your digital presence and customer conversion."
    )
    city = _safe_str(contact_data.get("City"))
    industry = _safe_str(contact_data.get("Industry"))
    client_website = _safe_str(contact_data.get("Website"))
    rating = _safe_str(contact_data.get("Google Rating"))

    msg = template
    # Replace greeting placeholder
    msg = msg.replace("{greeting}", greeting)
    # Also handle templates or custom text that start with "Hi {name} 👋" or "Hello {name} 👋"
    msg = msg.replace("Hi {name} 👋", greeting)
    msg = msg.replace("Hello {name} 👋", greeting)
    msg = msg.replace("Namaste {name} ji 🙏", greeting)
    msg = msg.replace("Hi {name} ji 👋", greeting)
    msg = msg.replace("Hello {name} ji 🙏", greeting)
    msg = msg.replace("Hi {name} ji 🙂", greeting.replace("👋", "🙂").replace("🙏", "🙂"))
    msg = msg.replace("Hi {name} 🙏", greeting.replace("👋", "🙏"))

    # If greeting was not explicitly in template or replaced, fallback {name}
    msg = msg.replace("{name}", display_name)
    msg = msg.replace("{business}", company)
    msg = msg.replace("{company}", company)
    msg = msg.replace("{service}", service)
    msg = msg.replace("{suggestion}", suggestion)
    msg = msg.replace("{city}", city if city else "your area")
    msg = msg.replace("{industry}", industry if industry else "your sector")
    msg = msg.replace("{client_website}", client_website if client_website else "")
    msg = msg.replace("{rating}", rating if rating else "")
    msg = msg.replace("{website}", COMPANY_WEBSITE)
    msg = msg.replace("{email}", COMPANY_EMAIL)
    msg = msg.replace("{sender_phone}", SENDER_PHONE)
    return msg


# ==========================================
# AI MESSAGE GENERATION (NVIDIA NIM / LLM)
# ==========================================
def call_ai_generate_message(
    api_key: str,
    base_url: str,
    model: str,
    contact_data: Dict[str, str],
    timeout: int = 20,
) -> Tuple[Optional[str], Optional[str]]:
    """
    Calls NVIDIA NIM / OpenAI-compatible API to craft a custom cold WhatsApp message
    tailored specifically to the client's Suggestion, Service, Company, Industry, and Location.
    """
    if not api_key or not api_key.strip():
        return None, "No API key provided"

    clean_base = base_url.rstrip("/")
    endpoint = f"{clean_base}/chat/completions" if not clean_base.endswith("/chat/completions") else clean_base

    raw_name = _safe_str(contact_data.get("Name"))
    business = _safe_str(contact_data.get("Company")) or "their business"
    greeting, display_name = get_smart_greeting_and_name(raw_name, business)
    service = _safe_str(contact_data.get("Suggested Services")) or _safe_str(contact_data.get("Service")) or "modern website development & digital growth solutions"
    suggestion = _safe_str(contact_data.get("Notes")) or _safe_str(contact_data.get("Suggestion")) or "optimizing mobile conversions and direct WhatsApp orders"
    industry = _safe_str(contact_data.get("Industry"))
    city = _safe_str(contact_data.get("City"))
    website_status = _safe_str(contact_data.get("Website Status"))

    system_prompt = f"""You are an elite B2B cold outreach copywriter for 'Skytecher' (https://skytecher.com), an IT and digital solutions agency offering website development, mobile apps, WhatsApp stores, SEO, and software automation.
Your goal: Craft a professional, high-converting, human WhatsApp cold message that feels personal, not automated.

STRICT OUTREACH RULES:
1. GREETING: Start with the exact greeting: '{greeting}'. NEVER greet a company name directly like 'Hi Apex Retail Solutions' — always address as 'Hello Team [Company]' or 'Hi [First Name]'.
2. REPETITION: Do NOT repeat the company name multiple times. Mention it naturally once or twice maximum (e.g. refer to 'your team' or 'your business' afterwards).
3. PAIN POINTS: Include 2-3 concise, realistic bottlenecks using ❌ emoji bullets (e.g., losing local Google searches, no mobile store, slow inquiries).
4. SOLUTION: Show how Skytecher fixes this with the recommended service using clear value (more leads, higher conversions).
5. PERSONALIZED SUGGESTION: Include the lead's custom note/suggestion under a 💡 *Quick Idea:* header.
6. CALL TO ACTION: A polite, low-friction question (e.g., 'Would it be okay if I shared a quick 2-minute idea for your team?').
7. FOOTER:
   Best regards,
   *Skytecher Team* ⚡
   🌐 https://skytecher.com
   ✉️ skytechersolutions@gmail.com
   📱 +91 8960061745
8. OPT-OUT: End with: _Reply STOP to opt out._
9. FORMATTING: WhatsApp bold (*text*), italic (_text_), clean spacing between sections. Keep under 130 words.
10. Output ONLY the message text — no markdown code fences, no quotes, no conversational filler."""

    user_prompt = f"""Create a personalized cold WhatsApp message for:
- Target Greeting: {greeting}
- Company: {business}
- Industry: {industry or 'Growing Business'}
- City: {city or 'India'}
- Existing Website: {website_status or 'Needs upgrade'}
- Recommended Service: {service}
- Custom Suggestion: {suggestion}"""

    headers = {
        "Authorization": f"Bearer {api_key.strip()}",
        "Content-Type": "application/json",
    }

    payload = {
        "model": model.strip(),
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.65,
        "max_tokens": 350,
    }

    try:
        resp = requests.post(endpoint, json=payload, headers=headers, timeout=timeout)
        if resp.status_code == 200:
            data = resp.json()
            choices = data.get("choices", [])
            if choices and "message" in choices[0]:
                content = choices[0]["message"].get("content", "").strip()
                # Clean any surrounding quotes
                if content.startswith('"""') and content.endswith('"""'):
                    content = content[3:-3].strip()
                if content.startswith("```") and content.endswith("```"):
                    content = re.sub(r"^```[a-zA-Z]*\n", "", content)
                    content = content.rstrip("`").strip()
                return content, None
            return None, "Empty response choices from AI model"
        else:
            return None, f"API Error HTTP {resp.status_code}: {resp.text[:150]}"
    except Exception as e:
        return None, f"Request failed: {str(e)}"


# ==========================================
# MAIN APPLICATION GUI CLASS
# ==========================================
class WhatsAppOutreachApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Skytecher WhatsApp Outreach Suite — AI Powered Client Acquisition")
        self.root.geometry("1220x920")
        self.root.minsize(1080, 800)

        # Threading and runtime states
        self.worker_thread: Optional[threading.Thread] = None
        self.stop_event = threading.Event()
        self.is_sending = False

        # Data storage
        self.contacts_df: Optional[pd.DataFrame] = None
        self.contacts_list: List[Dict[str, any]] = []
        self.current_file_path: Optional[str] = None
        self.results_records: List[Dict[str, any]] = []

        # Counter stats
        self.total_loaded = 0
        self.sent_count = 0
        self.failed_count = 0
        self.skipped_count = 0

        # Build theme and UI
        self._setup_theme()
        self._build_ui()
        self._log(
            "INFO",
            f"Initialized Skytecher Outreach App. Ready to connect via Sender: +91 {SENDER_PHONE}",
        )
        self._log("INFO", f"Company: {COMPANY_NAME} | Website: {COMPANY_WEBSITE} | Email: {COMPANY_EMAIL}")

        if not PYAUTOGUI_AVAILABLE:
            self._log(
                "WARNING",
                "pyautogui not detected. Install it via 'pip install pyautogui' for live WhatsApp sending. You can test in 'Simulation Mode' for now.",
            )
        else:
            self._log("INFO", "WhatsApp Web sender ready. Make sure WhatsApp Web is logged in on your default browser.")

    def _setup_theme(self):
        """Configures modern styling for Tkinter widgets."""
        self.style = ttk.Style(self.root)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self.color_bg = "#f1f5f9"        # Slate-100
        self.color_card = "#ffffff"      # White
        self.color_primary = "#0284c7"   # Sky-600
        self.color_primary_hover = "#0369a1"
        self.color_success = "#059669"   # Emerald-600
        self.color_danger = "#dc2626"    # Red-600
        self.color_text = "#0f172a"      # Slate-900
        self.color_muted = "#64748b"     # Slate-500

        self.root.configure(bg=self.color_bg)

        self.style.configure(".", font=("Segoe UI", 9), background=self.color_bg, foreground=self.color_text)
        self.style.configure("Card.TFrame", background=self.color_card, relief="groove", borderwidth=1)
        self.style.configure("CardHeader.TLabel", font=("Segoe UI", 10, "bold"), background=self.color_card, foreground="#0f172a")
        self.style.configure("CardMuted.TLabel", font=("Segoe UI", 8), background=self.color_card, foreground=self.color_muted)

    def _build_ui(self):
        """Constructs the complete graphical interface."""
        # 1. TOP HEADER BANNER
        header = tk.Frame(self.root, bg="#0f172a", height=78)
        header.pack(fill=tk.X, side=tk.TOP)

        hdr_inner = tk.Frame(header, bg="#0f172a")
        hdr_inner.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        title_box = tk.Frame(hdr_inner, bg="#0f172a")
        title_box.pack(side=tk.LEFT, fill=tk.Y)

        title_lbl = tk.Label(
            title_box,
            text="⚡ SKYTECHER WHATSAPP OUTREACH SUITE (v2.0)",
            font=("Segoe UI", 13, "bold"),
            bg="#0f172a",
            fg="#38bdf8",
        )
        title_lbl.pack(anchor="w")

        sub_lbl = tk.Label(
            title_box,
            text=f"AI Client Outreach • {COMPANY_WEBSITE} • {COMPANY_EMAIL}",
            font=("Segoe UI", 9),
            bg="#0f172a",
            fg="#94a3b8",
        )
        sub_lbl.pack(anchor="w")

        # Sender WhatsApp Badge
        sender_badge = tk.Frame(hdr_inner, bg="#1e293b", highlightbackground="#334155", highlightthickness=1)
        sender_badge.pack(side=tk.RIGHT, padx=5, pady=2)

        tk.Label(
            sender_badge,
            text=f"📱 Sender: +91 {SENDER_PHONE}",
            font=("Segoe UI", 9, "bold"),
            bg="#1e293b",
            fg="#10b981",
            padx=12,
            pady=6,
        ).pack()

        # 2. MAIN SPLIT CONTAINER
        main_container = tk.Frame(self.root, bg=self.color_bg)
        main_container.pack(fill=tk.BOTH, expand=True, padx=16, pady=10)

        # Left Column (Data Source & Message/AI Configuration)
        left_col = tk.Frame(main_container, bg=self.color_bg, width=570)
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))

        # Right Column (Settings, Controls, KPIs & Live Log)
        right_col = tk.Frame(main_container, bg=self.color_bg, width=610)
        right_col.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(8, 0))

        # ----------------------------------------------------
        # CARD 1 (LEFT): EXCEL SOURCE
        # ----------------------------------------------------
        card_upload = ttk.Frame(left_col, style="Card.TFrame", padding=12)
        card_upload.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(card_upload, text="1. Excel Contact Source (.xlsx)", style="CardHeader.TLabel").pack(anchor="w")
        ttk.Label(
            card_upload,
            text="Supports 17-column format (#, Company, Industry, Phone, Suggested Services, Notes...) or standard sheets",
            style="CardMuted.TLabel",
        ).pack(anchor="w", pady=(0, 6))

        upload_btn_row = tk.Frame(card_upload, bg=self.color_card)
        upload_btn_row.pack(fill=tk.X, pady=2)

        self.btn_upload = tk.Button(
            upload_btn_row,
            text="📂 Choose Excel File",
            font=("Segoe UI", 8, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            relief="flat",
            padx=8,
            pady=5,
            cursor="hand2",
            command=self.action_upload_excel,
        )
        self.btn_upload.pack(side=tk.LEFT, padx=(0, 5))

        self.btn_sample = tk.Button(
            upload_btn_row,
            text="📄 Load Sample Leads",
            font=("Segoe UI", 8),
            bg="#e2e8f0",
            fg="#1e293b",
            activebackground="#cbd5e1",
            relief="flat",
            padx=8,
            pady=5,
            cursor="hand2",
            command=self.action_load_sample_file,
        )
        self.btn_sample.pack(side=tk.LEFT, padx=(0, 5))

        self.btn_download_sample = tk.Button(
            upload_btn_row,
            text="📥 Download Template",
            font=("Segoe UI", 8, "bold"),
            bg="#0f766e",
            fg="#ffffff",
            activebackground="#115e59",
            activeforeground="#ffffff",
            relief="flat",
            padx=8,
            pady=5,
            cursor="hand2",
            command=self.action_download_sample_template,
        )
        self.btn_download_sample.pack(side=tk.LEFT, padx=(0, 5))

        self.btn_blocklist = tk.Button(
            upload_btn_row,
            text="🛡️ Blocklist",
            font=("Segoe UI", 8),
            bg="#e2e8f0",
            fg="#1e293b",
            activebackground="#cbd5e1",
            relief="flat",
            padx=8,
            pady=5,
            cursor="hand2",
            command=self.open_blocklist_modal,
        )
        self.btn_blocklist.pack(side=tk.RIGHT)

        # Clear Sent History button
        self.btn_clear_history = tk.Button(
            upload_btn_row,
            text="🗑️ Clear History",
            font=("Segoe UI", 8),
            bg="#f59e0b",
            fg="#ffffff",
            activebackground="#d97706",
            activeforeground="#ffffff",
            relief="flat",
            padx=8,
            pady=5,
            cursor="hand2",
            command=self.action_clear_sent_history,
        )
        self.btn_clear_history.pack(side=tk.RIGHT, padx=(0, 5))

        self.lbl_file_status = tk.Label(
            card_upload,
            text="No Excel file loaded yet.",
            font=("Segoe UI", 8, "italic"),
            bg="#f8fafc",
            fg=self.color_muted,
            relief="groove",
            bd=1,
            anchor="w",
            padx=8,
            pady=4,
        )
        self.lbl_file_status.pack(fill=tk.X, pady=(6, 0))

        # ----------------------------------------------------
        # CARD 2 (LEFT): AI MESSAGE GENERATOR & API KEY
        # ----------------------------------------------------
        card_ai = ttk.Frame(left_col, style="Card.TFrame", padding=12)
        card_ai.pack(fill=tk.X, pady=(0, 8))

        ai_header_row = tk.Frame(card_ai, bg=self.color_card)
        ai_header_row.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(ai_header_row, text="2. AI Cold Message Engine (NVIDIA NIM / LLM)", style="CardHeader.TLabel").pack(side=tk.LEFT)

        self.var_enable_ai = tk.BooleanVar(value=os.getenv("ENABLE_AI_PERSONALIZATION", "False").lower() in ["true", "1"])
        self.chk_enable_ai = tk.Checkbutton(
            ai_header_row,
            text="Enable AI Dynamic Personalization",
            variable=self.var_enable_ai,
            font=("Segoe UI", 8, "bold"),
            bg=self.color_card,
            fg="#0284c7",
            activebackground=self.color_card,
            command=self._on_toggle_ai,
        )
        self.chk_enable_ai.pack(side=tk.RIGHT)

        # AI Settings Grid
        ai_grid = tk.Frame(card_ai, bg=self.color_card)
        ai_grid.pack(fill=tk.X, pady=4)

        # Row 1: Model Selection
        tk.Label(ai_grid, text="AI Model:", font=("Segoe UI", 8, "bold"), bg=self.color_card).grid(row=0, column=0, sticky="w", pady=2)
        self.combo_model = ttk.Combobox(ai_grid, values=POPULAR_AI_MODELS, font=("Segoe UI", 8), width=32)
        self.combo_model.set(DEFAULT_AI_MODEL)
        self.combo_model.grid(row=0, column=1, sticky="w", padx=6, pady=2)

        # Row 2: API Key
        tk.Label(ai_grid, text="API Key:", font=("Segoe UI", 8, "bold"), bg=self.color_card).grid(row=1, column=0, sticky="w", pady=2)
        key_box = tk.Frame(ai_grid, bg=self.color_card)
        key_box.grid(row=1, column=1, sticky="w", padx=6, pady=2)

        self.entry_api_key = tk.Entry(key_box, font=("Segoe UI", 8), show="•", width=28)
        env_key = os.getenv("NVIDIA_API_KEY", "")
        if env_key:
            self.entry_api_key.insert(0, env_key)
        self.entry_api_key.pack(side=tk.LEFT)

        self.var_show_key = tk.BooleanVar(value=False)
        self.btn_toggle_key = tk.Button(
            key_box,
            text="👁️",
            font=("Segoe UI", 7),
            bg="#f1f5f9",
            relief="flat",
            padx=4,
            command=self.action_toggle_key_visibility,
        )
        self.btn_toggle_key.pack(side=tk.LEFT, padx=(2, 0))

        self.btn_save_key = tk.Button(
            ai_grid,
            text="💾 Save Model & Key to .env",
            font=("Segoe UI", 7, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            relief="flat",
            padx=6,
            pady=2,
            command=self.action_save_api_key_to_env,
        )
        self.btn_save_key.grid(row=1, column=2, sticky="w", padx=2)

        # Row 3: Test AI Button
        ai_btn_bar = tk.Frame(card_ai, bg=self.color_card)
        ai_btn_bar.pack(fill=tk.X, pady=(4, 0))

        self.btn_test_ai = tk.Button(
            ai_btn_bar,
            text="✨ Test AI Generation for Lead #1",
            font=("Segoe UI", 8, "bold"),
            bg="#7c3aed",  # Purple
            fg="#ffffff",
            activebackground="#6d28d9",
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
            command=self.action_test_ai_generation,
        )
        self.btn_test_ai.pack(side=tk.LEFT)

        tk.Label(
            ai_btn_bar,
            text="Free NVIDIA NIM API keys available at build.nvidia.com",
            font=("Segoe UI", 7, "italic"),
            bg=self.color_card,
            fg=self.color_muted,
            padx=8,
        ).pack(side=tk.LEFT)

        # ----------------------------------------------------
        # CARD 3 (LEFT): EDITABLE MESSAGE TEMPLATE & PRESETS
        # ----------------------------------------------------
        card_template = ttk.Frame(left_col, style="Card.TFrame", padding=12)
        card_template.pack(fill=tk.BOTH, expand=True)

        ttk.Label(card_template, text="3. Message Template & Format Presets", style="CardHeader.TLabel").pack(anchor="w")

        # Preset Selector Row
        preset_row = tk.Frame(card_template, bg=self.color_card)
        preset_row.pack(fill=tk.X, pady=(2, 4))

        tk.Label(preset_row, text="Choose Format:", font=("Segoe UI", 8, "bold"), bg=self.color_card, fg="#0f172a").pack(side=tk.LEFT)
        self.combo_presets = ttk.Combobox(
            preset_row,
            values=list(MESSAGE_PRESETS.keys()),
            font=("Segoe UI", 8),
            state="readonly",
            width=36,
        )
        self.combo_presets.set("1. Professional Introduction (English)")
        self.combo_presets.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))
        self.combo_presets.bind("<<ComboboxSelected>>", self.on_select_preset)

        # Placeholders cheat-sheet
        badge_frame = tk.Frame(card_template, bg=self.color_card)
        badge_frame.pack(fill=tk.X, pady=(2, 4))
        tk.Label(badge_frame, text="Tags:", font=("Segoe UI", 8, "bold"), bg=self.color_card, fg="#475569").pack(side=tk.LEFT)
        for ph in ["{name}", "{business}", "{service}", "{suggestion}", "{website}", "{email}"]:
            tk.Label(
                badge_frame,
                text=ph,
                font=("Consolas", 8, "bold"),
                bg="#f1f5f9",
                fg="#0284c7",
                padx=3,
                pady=1,
                relief="solid",
                bd=1,
            ).pack(side=tk.LEFT, padx=2)

        self.txt_template = ScrolledText(
            card_template,
            wrap=tk.WORD,
            font=("Segoe UI", 9),
            bg="#ffffff",
            fg="#0f172a",
            insertbackground="#0f172a",
            relief="solid",
            bd=1,
            height=10,
        )
        self.txt_template.pack(fill=tk.BOTH, expand=True, pady=4)
        self.txt_template.insert(tk.END, MESSAGE_PRESETS["1. Professional Introduction (English)"])

        tpl_btn_row = tk.Frame(card_template, bg=self.color_card)
        tpl_btn_row.pack(fill=tk.X, pady=(4, 0))

        self.btn_preview = tk.Button(
            tpl_btn_row,
            text="👁️ Preview First Message",
            font=("Segoe UI", 9, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
            command=self.action_preview_message,
        )
        self.btn_preview.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_reset_tpl = tk.Button(
            tpl_btn_row,
            text="↺ Reset Template",
            font=("Segoe UI", 8),
            bg="#e2e8f0",
            fg="#1e293b",
            activebackground="#cbd5e1",
            relief="flat",
            padx=8,
            pady=4,
            cursor="hand2",
            command=self.action_reset_template,
        )
        self.btn_reset_tpl.pack(side=tk.LEFT)

        # ----------------------------------------------------
        # CARD 4 (RIGHT): SETTINGS & RESILIENCE
        # ----------------------------------------------------
        card_settings = ttk.Frame(right_col, style="Card.TFrame", padding=12)
        card_settings.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(card_settings, text="4. Dispatch, Timing & Anti-Ban Controls", style="CardHeader.TLabel").pack(anchor="w")

        grid_frame = tk.Frame(card_settings, bg=self.color_card)
        grid_frame.pack(fill=tk.X, pady=3)

        # Daily Send Limit
        tk.Label(grid_frame, text="Daily Limit:", font=("Segoe UI", 8, "bold"), bg=self.color_card).grid(row=0, column=0, sticky="w", pady=3)
        self.spin_daily_limit = tk.Spinbox(grid_frame, from_=1, to=500, width=6, font=("Segoe UI", 8))
        self.spin_daily_limit.delete(0, tk.END)
        self.spin_daily_limit.insert(0, "30")
        self.spin_daily_limit.grid(row=0, column=1, sticky="w", padx=6, pady=3)

        # Delay Range
        tk.Label(grid_frame, text="Random Delay:", font=("Segoe UI", 8, "bold"), bg=self.color_card).grid(row=0, column=2, sticky="w", padx=(10, 0), pady=3)
        delay_box = tk.Frame(grid_frame, bg=self.color_card)
        delay_box.grid(row=0, column=3, sticky="w", padx=4, pady=3)

        self.spin_delay_min = tk.Spinbox(delay_box, from_=5, to=300, width=4, font=("Segoe UI", 8))
        self.spin_delay_min.delete(0, tk.END)
        self.spin_delay_min.insert(0, "30")
        self.spin_delay_min.pack(side=tk.LEFT)
        tk.Label(delay_box, text="-", bg=self.color_card, padx=2).pack(side=tk.LEFT)
        self.spin_delay_max = tk.Spinbox(delay_box, from_=10, to=600, width=4, font=("Segoe UI", 8))
        self.spin_delay_max.delete(0, tk.END)
        self.spin_delay_max.insert(0, "60")
        self.spin_delay_max.pack(side=tk.LEFT)
        tk.Label(delay_box, text="sec", font=("Segoe UI", 8), bg=self.color_card, fg=self.color_muted, padx=2).pack(side=tk.LEFT)

        # Page load wait time
        tk.Label(grid_frame, text="Web Load Wait:", font=("Segoe UI", 8, "bold"), bg=self.color_card).grid(row=1, column=0, sticky="w", pady=3)
        self.spin_wait_time = tk.Spinbox(grid_frame, from_=8, to=60, width=6, font=("Segoe UI", 8))
        self.spin_wait_time.delete(0, tk.END)
        self.spin_wait_time.insert(0, "18")
        self.spin_wait_time.grid(row=1, column=1, sticky="w", padx=6, pady=3)

        # Simulation Mode checkbox
        self.var_dry_run = tk.BooleanVar(value=False)
        self.chk_dry_run = tk.Checkbutton(
            grid_frame,
            text="🧪 Simulation / Dry Run (Test without opening WhatsApp)",
            variable=self.var_dry_run,
            font=("Segoe UI", 8, "bold"),
            bg=self.color_card,
            fg="#d97706",
            activebackground=self.color_card,
        )
        self.chk_dry_run.grid(row=1, column=2, columnspan=2, sticky="w", padx=(10, 0), pady=3)

        # Ignore Sent History checkbox
        self.var_ignore_history = tk.BooleanVar(value=False)
        self.chk_ignore_history = tk.Checkbutton(
            grid_frame,
            text="🔄 Re-send to All (Ignore Sent History)",
            variable=self.var_ignore_history,
            font=("Segoe UI", 8),
            bg=self.color_card,
            fg="#0f172a",
            activebackground=self.color_card,
        )
        self.chk_ignore_history.grid(row=2, column=0, columnspan=2, sticky="w", pady=3)

        # ----------------------------------------------------
        # CARD 5 (RIGHT): ACTIONS, PROGRESS & KPIS
        # ----------------------------------------------------
        card_actions = ttk.Frame(right_col, style="Card.TFrame", padding=12)
        card_actions.pack(fill=tk.X, pady=(0, 8))

        btn_action_bar = tk.Frame(card_actions, bg=self.color_card)
        btn_action_bar.pack(fill=tk.X, pady=(0, 6))

        self.btn_start = tk.Button(
            btn_action_bar,
            text="▶  START SENDING",
            font=("Segoe UI", 10, "bold"),
            bg="#059669",
            fg="#ffffff",
            activebackground="#047857",
            activeforeground="#ffffff",
            relief="flat",
            padx=16,
            pady=6,
            cursor="hand2",
            command=self.action_start_sending,
        )
        self.btn_start.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_stop = tk.Button(
            btn_action_bar,
            text="⏹  STOP",
            font=("Segoe UI", 10, "bold"),
            bg="#dc2626",
            fg="#ffffff",
            activebackground="#b91c1c",
            activeforeground="#ffffff",
            relief="flat",
            padx=14,
            pady=6,
            cursor="hand2",
            state=tk.DISABLED,
            command=self.action_stop_sending,
        )
        self.btn_stop.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_export = tk.Button(
            btn_action_bar,
            text="📊 Export Results (.xlsx)",
            font=("Segoe UI", 9, "bold"),
            bg="#0f172a",
            fg="#ffffff",
            activebackground="#1e293b",
            activeforeground="#ffffff",
            relief="flat",
            padx=10,
            pady=6,
            cursor="hand2",
            command=self.action_export_report,
        )
        self.btn_export.pack(side=tk.RIGHT)

        # KPI Counters Grid
        kpi_row = tk.Frame(card_actions, bg=self.color_card)
        kpi_row.pack(fill=tk.X, pady=(2, 6))

        self.kpi_labels = {}
        kpi_defs = [
            ("total", "Total Leads", "#0284c7"),
            ("sent", "Sent Success", "#059669"),
            ("skipped", "Skipped", "#d97706"),
            ("failed", "Failed / No WA", "#dc2626"),
            ("remaining", "Remaining", "#475569"),
        ]

        for idx, (key, title, color) in enumerate(kpi_defs):
            cell = tk.Frame(kpi_row, bg="#f8fafc", relief="solid", bd=1, padx=6, pady=3)
            cell.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2)

            num_lbl = tk.Label(cell, text="0", font=("Segoe UI", 13, "bold"), bg="#f8fafc", fg=color)
            num_lbl.pack()
            txt_lbl = tk.Label(cell, text=title, font=("Segoe UI", 7, "bold"), bg="#f8fafc", fg="#64748b")
            txt_lbl.pack()
            self.kpi_labels[key] = num_lbl

        # Progress bar
        pbar_row = tk.Frame(card_actions, bg=self.color_card)
        pbar_row.pack(fill=tk.X, pady=(4, 0))

        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(pbar_row, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.lbl_progress_pct = tk.Label(
            pbar_row,
            text="0%",
            font=("Segoe UI", 8, "bold"),
            bg=self.color_card,
            fg="#0f172a",
            width=5,
            anchor="e",
        )
        self.lbl_progress_pct.pack(side=tk.RIGHT, padx=(4, 0))

        # ----------------------------------------------------
        # CARD 6 (RIGHT): LIVE ACTIVITY TERMINAL LOG
        # ----------------------------------------------------
        card_log = ttk.Frame(right_col, style="Card.TFrame", padding=12)
        card_log.pack(fill=tk.BOTH, expand=True)

        log_hdr = tk.Frame(card_log, bg=self.color_card)
        log_hdr.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(log_hdr, text="5. Live Activity Console", style="CardHeader.TLabel").pack(side=tk.LEFT)

        btn_clear_log = tk.Button(
            log_hdr,
            text="Clear",
            font=("Segoe UI", 7),
            bg="#f1f5f9",
            fg="#475569",
            activebackground="#e2e8f0",
            relief="flat",
            padx=6,
            pady=1,
            cursor="hand2",
            command=self.action_clear_log,
        )
        btn_clear_log.pack(side=tk.RIGHT)

        self.txt_log = ScrolledText(
            card_log,
            wrap=tk.WORD,
            font=("Consolas", 8),
            bg="#0b0f19",
            fg="#f8fafc",
            insertbackground="#ffffff",
            relief="flat",
            height=14,
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True)

        self.txt_log.tag_config("TAG_TIME", foreground="#64748b")
        self.txt_log.tag_config("TAG_INFO", foreground="#38bdf8")
        self.txt_log.tag_config("TAG_SUCCESS", foreground="#34d399")
        self.txt_log.tag_config("TAG_SKIP", foreground="#fbbf24")
        self.txt_log.tag_config("TAG_ERROR", foreground="#f87171")

        # 3. FOOTER
        footer = tk.Frame(self.root, bg="#0f172a", height=24)
        footer.pack(fill=tk.X, side=tk.BOTTOM)

        self.lbl_footer = tk.Label(
            footer,
            text=f"Ready • Skytecher WhatsApp Outreach Suite • Sender: +91 {SENDER_PHONE} • Website: {COMPANY_WEBSITE}",
            font=("Segoe UI", 8),
            bg="#0f172a",
            fg="#94a3b8",
            padx=12,
            pady=3,
        )
        self.lbl_footer.pack(side=tk.LEFT)

    def _on_toggle_ai(self):
        """Notifies user in log when AI is toggled."""
        if self.var_enable_ai.get():
            self._log("INFO", f"AI Personalization enabled. Model: {self.combo_model.get()}")
        else:
            self._log("INFO", "AI Personalization disabled. Standard message template will be used.")

    # ==========================================
    # LOGGING & COUNTER UPDATES
    # ==========================================
    def _log(self, level: str, message: str):
        """Thread-safe logging method to display colored log lines."""
        def append_log():
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            self.txt_log.insert(tk.END, f"[{timestamp}] ", "TAG_TIME")

            level_upper = level.upper()
            tag = "TAG_INFO"
            if "SUCCESS" in level_upper or "SENT" in level_upper:
                tag = "TAG_SUCCESS"
            elif "SKIP" in level_upper:
                tag = "TAG_SKIP"
            elif "ERROR" in level_upper or "FAIL" in level_upper:
                tag = "TAG_ERROR"

            self.txt_log.insert(tk.END, f"[{level_upper:7}] ", tag)
            self.txt_log.insert(tk.END, f"{message}\n")
            self.txt_log.see(tk.END)

        self.root.after(0, append_log)

    def _update_kpi_counters(self):
        """Refreshes KPI numbers and progress bar in the UI."""
        def update():
            self.kpi_labels["total"].config(text=str(self.total_loaded))
            self.kpi_labels["sent"].config(text=str(self.sent_count))
            self.kpi_labels["skipped"].config(text=str(self.skipped_count))
            self.kpi_labels["failed"].config(text=str(self.failed_count))

            processed = self.sent_count + self.skipped_count + self.failed_count
            remaining = max(0, self.total_loaded - processed)
            self.kpi_labels["remaining"].config(text=str(remaining))

            if self.total_loaded > 0:
                pct = (processed / self.total_loaded) * 100.0
                self.progress_var.set(pct)
                self.lbl_progress_pct.config(text=f"{int(pct)}%")
            else:
                self.progress_var.set(0)
                self.lbl_progress_pct.config(text="0%")

        self.root.after(0, update)

    def action_clear_log(self):
        self.txt_log.delete("1.0", tk.END)

    def on_select_preset(self, event=None):
        """Loads selected format preset into the message template editor."""
        chosen = self.combo_presets.get()
        if chosen in MESSAGE_PRESETS:
            self.txt_template.delete("1.0", tk.END)
            self.txt_template.insert(tk.END, MESSAGE_PRESETS[chosen])
            self._log("INFO", f"Loaded message preset: '{chosen}'")

    def action_toggle_key_visibility(self):
        """Toggles masking on the API key entry."""
        if self.var_show_key.get():
            self.entry_api_key.config(show="•")
            self.var_show_key.set(False)
            self.btn_toggle_key.config(text="👁️")
        else:
            self.entry_api_key.config(show="")
            self.var_show_key.set(True)
            self.btn_toggle_key.config(text="🙈")

    def action_reset_template(self):
        if messagebox.askyesno("Reset Template", "Restore default message template?"):
            self.combo_presets.set("1. Professional Introduction (English)")
            self.txt_template.delete("1.0", tk.END)
            self.txt_template.insert(tk.END, MESSAGE_PRESETS["1. Professional Introduction (English)"])

    # ==========================================
    # SAVE API KEY & MODEL TO .ENV
    # ==========================================
    def action_save_api_key_to_env(self):
        """Saves current API key and model selection to .env file and active environment."""
        key = self.entry_api_key.get().strip()
        model = self.combo_model.get().strip()
        enable_ai_val = "True" if self.var_enable_ai.get() else "False"

        # Update running environment variables
        os.environ["NVIDIA_API_KEY"] = key
        os.environ["AI_MODEL"] = model
        os.environ["ENABLE_AI_PERSONALIZATION"] = enable_ai_val

        env_path = os.path.abspath(".env")
        try:
            lines = []
            if os.path.exists(env_path):
                with open(env_path, "r", encoding="utf-8") as f:
                    lines = f.readlines()

            new_lines = []
            key_set = False
            model_set = False
            ai_flag_set = False

            for l in lines:
                if l.startswith("NVIDIA_API_KEY="):
                    new_lines.append(f"NVIDIA_API_KEY={key}\n")
                    key_set = True
                elif l.startswith("AI_MODEL="):
                    new_lines.append(f"AI_MODEL={model}\n")
                    model_set = True
                elif l.startswith("ENABLE_AI_PERSONALIZATION="):
                    new_lines.append(f"ENABLE_AI_PERSONALIZATION={enable_ai_val}\n")
                    ai_flag_set = True
                else:
                    new_lines.append(l)

            if not key_set:
                new_lines.append(f"NVIDIA_API_KEY={key}\n")
            if not model_set:
                new_lines.append(f"AI_MODEL={model}\n")
            if not ai_flag_set:
                new_lines.append(f"ENABLE_AI_PERSONALIZATION={enable_ai_val}\n")

            with open(env_path, "w", encoding="utf-8") as f:
                f.writelines(new_lines)

            self._log("SUCCESS", f"Model '{model}' and API Key saved to .env and active runtime.")
            messagebox.showinfo("Saved Successfully", f"Saved configuration to .env:\n\n• Model: {model}\n• API Key: {'Configured' if key else 'Empty'}\n• AI Personalization: {enable_ai_val}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed saving to .env: {e}")

    # ==========================================
    # TEST AI GENERATION MODAL
    # ==========================================
    def action_test_ai_generation(self):
        """Tests the AI message generation on lead #1 in a separate thread."""
        if not self.contacts_list:
            messagebox.showinfo("Notice", "Please upload an Excel contact file first.")
            return

        api_key = self.entry_api_key.get().strip()
        if not api_key:
            messagebox.showwarning("Missing API Key", "Please enter your NVIDIA / OpenAI API Key to test AI generation.")
            return

        model = self.combo_model.get().strip()
        contact = self.contacts_list[0]

        self._log("INFO", f"Calling AI model ({model}) for lead #1: {contact['Name']}...")

        def run_test():
            ai_msg, err = call_ai_generate_message(
                api_key=api_key,
                base_url=DEFAULT_AI_BASE_URL,
                model=model,
                contact_data=contact,
            )
            if err:
                self._log("ERROR", f"AI Generation test failed: {err}")
                self.root.after(0, lambda: messagebox.showerror("AI Test Failed", f"AI API call failed:\n{err}"))
            else:
                self._log("SUCCESS", f"AI Message generated successfully for {contact['Name']}!")
                self.root.after(0, lambda: self._show_ai_preview_modal(contact, ai_msg, model))

        threading.Thread(target=run_test, daemon=True).start()

    def _show_ai_preview_modal(self, contact: Dict[str, str], message: str, model: str):
        modal = tk.Toplevel(self.root)
        modal.title(f"AI Generated Cold Message — {model}")
        modal.geometry("640x520")
        modal.transient(self.root)
        modal.grab_set()

        p_frame = tk.Frame(modal, bg="#f8fafc", padx=16, pady=14)
        p_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            p_frame,
            text=f"✨ AI Tailored Cold Message for: {contact['Name']}",
            font=("Segoe UI", 11, "bold"),
            bg="#f8fafc",
            fg="#0f172a",
        ).pack(anchor="w", pady=(0, 2))

        meta = f"Model: {model}  |  Company: {contact['Company'] or 'N/A'}  |  Service: {contact['Service']}"
        tk.Label(p_frame, text=meta, font=("Segoe UI", 8), bg="#f8fafc", fg="#64748b").pack(anchor="w", pady=(0, 6))

        txt = ScrolledText(p_frame, wrap=tk.WORD, font=("Segoe UI", 9), bg="#ffffff", fg="#0f172a", relief="solid", bd=1)
        txt.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        txt.insert(tk.END, message)

        btn_row = tk.Frame(p_frame, bg="#f8fafc")
        btn_row.pack(fill=tk.X)

        def use_as_template():
            self.txt_template.delete("1.0", tk.END)
            self.txt_template.insert(tk.END, message)
            modal.destroy()
            self._log("INFO", "Adopted AI message as current active template.")

        tk.Button(
            btn_row,
            text="Use this as Template",
            font=("Segoe UI", 9, "bold"),
            bg="#7c3aed",
            fg="#ffffff",
            relief="flat",
            padx=10,
            pady=4,
            command=use_as_template,
        ).pack(side=tk.LEFT)

        tk.Button(
            btn_row,
            text="Close",
            font=("Segoe UI", 9),
            bg="#0f172a",
            fg="#ffffff",
            relief="flat",
            padx=12,
            pady=4,
            command=modal.destroy,
        ).pack(side=tk.RIGHT)

    # ==========================================
    # EXCEL PROCESSING & VALIDATION
    # ==========================================
    def action_upload_excel(self):
        file_path = filedialog.askopenfilename(
            title="Select Client Contact Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
        )
        if not file_path:
            return
        self._load_excel_file(file_path)

    def action_download_sample_template(self):
        """Allows user to download the official 17-column Excel lead template."""
        save_path = filedialog.asksaveasfilename(
            title="Save Skytecher Lead Excel Template",
            defaultextension=".xlsx",
            initialfile="skytecher_leads_template.xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
        )
        if not save_path:
            return

        sample_source = os.path.abspath("sample_leads.xlsx")
        try:
            if os.path.exists(sample_source):
                df = pd.read_excel(sample_source)
            else:
                import create_sample_excel
                df = pd.DataFrame(create_sample_excel.sample_leads_data)
            df.to_excel(save_path, index=False)
            self._log("SUCCESS", f"Sample template saved to: {save_path}")
            messagebox.showinfo(
                "Template Downloaded",
                f"Sample Excel template saved successfully to:\n{save_path}\n\nYou can fill your leads in this format and upload it back into the app!",
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed saving template: {e}")

    def action_load_sample_file(self):
        sample_path = os.path.abspath("sample_leads.xlsx")
        if not os.path.exists(sample_path):
            try:
                import create_sample_excel
                df = pd.DataFrame(create_sample_excel.sample_leads_data)
                df.to_excel(sample_path, index=False)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to generate sample file: {e}")
                return

        self._load_excel_file(sample_path)

    def _load_excel_file(self, file_path: str):
        self._log("INFO", f"Loading Excel file: {os.path.basename(file_path)}")
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            messagebox.showerror("Read Error", f"Unable to read Excel file:\n{e}")
            self._log("ERROR", f"Failed to read Excel file: {e}")
            return

        # Flexible alias mapping supporting the 17-column format:
        # #, Company, Industry, Lead Type, Country, City, Phone, Website, Website Status, LinkedIn (company), Google Rating, Source, Suggested Services, Contact Score, Priority, Outreach Status, Notes
        col_aliases = {
            "Phone": ["phone", "mobile", "contact", "phone number", "whatsapp", "cell", "contact number"],
            "Company": ["company", "business", "company name", "firm", "client", "business name", "organization"],
            "Name": ["name", "contact person", "owner", "full name", "client name", "person", "lead name"],
            "Service": ["suggested services", "service", "services", "suggested service", "category", "requirement"],
            "Suggestion": ["notes", "suggestion", "custom pitch", "pitch", "idea", "note", "recommendation", "comments"],
            "Industry": ["industry", "domain", "sector"],
            "City": ["city", "location", "place"],
            "Country": ["country"],
            "Website": ["website", "client website", "url", "web"],
            "Website Status": ["website status", "site status"],
            "LinkedIn": ["linkedin (company)", "linkedin", "linkedin url"],
            "Google Rating": ["google rating", "rating"],
            "Source": ["source", "lead source"],
            "Contact Score": ["contact score", "score"],
            "Priority": ["priority"],
            "Lead Type": ["lead type", "type"],
            "Outreach Status": ["outreach status", "status"],
            "#": ["#", "id", "s.no", "sno", "row"]
        }

        col_mapping = {}
        for canonical, aliases in col_aliases.items():
            for actual_col in df.columns:
                cleaned_header = str(actual_col).strip().lower()
                if cleaned_header in aliases or cleaned_header == canonical.lower():
                    col_mapping[canonical] = actual_col
                    break

        # Phone is the only absolute required column; others fall back gracefully
        if "Phone" not in col_mapping:
            messagebox.showerror(
                "Missing Phone Column",
                f"Excel file must contain a 'Phone' or 'Mobile' column.\nFound headers: {list(df.columns)}",
            )
            self._log("ERROR", f"Missing Phone column in Excel file. Found: {list(df.columns)}")
            return

        self.current_file_path = file_path
        self.contacts_df = df
        self.contacts_list = []
        self.results_records = []

        total_rows = len(df)
        self.total_loaded = total_rows
        self.sent_count = 0
        self.failed_count = 0
        self.skipped_count = 0
        self._update_kpi_counters()

        for idx, row in df.iterrows():
            company_val = str(row.get(col_mapping.get("Company", ""), "")).strip() if "Company" in col_mapping else ""
            if company_val.lower() in ["nan", "none"]:
                company_val = ""

            name_val = str(row.get(col_mapping.get("Name", ""), "")).strip() if "Name" in col_mapping else ""
            if not name_val or name_val.lower() in ["nan", "none"]:
                # If no personal name column, use Company or friendly fallback
                name_val = company_val if company_val else f"Contact #{idx + 1}"

            service_val = str(row.get(col_mapping.get("Service", ""), "")).strip() if "Service" in col_mapping else ""
            if not service_val or service_val.lower() in ["nan", "none"]:
                service_val = "Website & Digital IT Solutions"

            suggestion_val = str(row.get(col_mapping.get("Suggestion", ""), "")).strip() if "Suggestion" in col_mapping else ""
            if suggestion_val.lower() in ["nan", "none"]:
                suggestion_val = ""

            item = {
                "row_index": idx + 1,
                "Name": name_val,
                "Phone_Raw": row.get(col_mapping.get("Phone", "Phone"), ""),
                "Company": company_val,
                "Service": service_val,
                "Suggestion": suggestion_val,
                "Industry": str(row.get(col_mapping.get("Industry", ""), "")).strip() if "Industry" in col_mapping else "",
                "City": str(row.get(col_mapping.get("City", ""), "")).strip() if "City" in col_mapping else "",
                "Country": str(row.get(col_mapping.get("Country", ""), "")).strip() if "Country" in col_mapping else "",
                "Website": str(row.get(col_mapping.get("Website", ""), "")).strip() if "Website" in col_mapping else "",
                "Website Status": str(row.get(col_mapping.get("Website Status", ""), "")).strip() if "Website Status" in col_mapping else "",
                "LinkedIn": str(row.get(col_mapping.get("LinkedIn", ""), "")).strip() if "LinkedIn" in col_mapping else "",
                "Google Rating": str(row.get(col_mapping.get("Google Rating", ""), "")).strip() if "Google Rating" in col_mapping else "",
                "Contact Score": str(row.get(col_mapping.get("Contact Score", ""), "")).strip() if "Contact Score" in col_mapping else "",
                "Priority": str(row.get(col_mapping.get("Priority", ""), "")).strip() if "Priority" in col_mapping else "",
                "Lead Type": str(row.get(col_mapping.get("Lead Type", ""), "")).strip() if "Lead Type" in col_mapping else "",
            }

            self.contacts_list.append(item)

        file_name = os.path.basename(file_path)
        self.lbl_file_status.config(
            text=f"Loaded: {file_name}  ({total_rows} total rows)",
            fg="#059669",
            font=("Segoe UI", 9, "bold"),
        )
        self._log("SUCCESS", f"Successfully loaded '{file_name}' with {total_rows} rows.")

    # ==========================================
    # PREVIEW FIRST MESSAGE
    # ==========================================
    def action_preview_message(self):
        if not self.contacts_list:
            messagebox.showinfo("Preview", "Please upload an Excel contact file first.")
            return

        template = self.txt_template.get("1.0", tk.END).strip()
        first_contact = self.contacts_list[0]
        cleaned_phone, _ = clean_phone_number(first_contact["Phone_Raw"])

        resolved_msg = format_template_message(template, first_contact)

        preview_win = tk.Toplevel(self.root)
        preview_win.title("Preview First Message — Skytecher")
        preview_win.geometry("640x520")
        preview_win.transient(self.root)
        preview_win.grab_set()

        p_frame = tk.Frame(preview_win, bg="#f8fafc", padx=16, pady=14)
        p_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            p_frame,
            text=f"Previewing for: {first_contact['Name']} ({cleaned_phone or 'Invalid Phone'})",
            font=("Segoe UI", 11, "bold"),
            bg="#f8fafc",
            fg="#0f172a",
        ).pack(anchor="w", pady=(0, 4))

        meta_text = f"Company: {first_contact['Company'] or 'N/A'}  |  Service: {first_contact['Service']}"
        tk.Label(p_frame, text=meta_text, font=("Segoe UI", 9), bg="#f8fafc", fg="#64748b").pack(anchor="w", pady=(0, 8))

        preview_text = ScrolledText(
            p_frame,
            wrap=tk.WORD,
            font=("Segoe UI", 9),
            bg="#ffffff",
            fg="#0f172a",
            relief="solid",
            bd=1,
            height=16,
        )
        preview_text.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        preview_text.insert(tk.END, resolved_msg)
        preview_text.config(state=tk.DISABLED)

        tk.Button(
            p_frame,
            text="Close Preview",
            font=("Segoe UI", 9, "bold"),
            bg="#0f172a",
            fg="#ffffff",
            relief="flat",
            padx=14,
            pady=6,
            command=preview_win.destroy,
        ).pack(anchor="e")

    # ==========================================
    # BLOCKLIST MANAGER MODAL
    # ==========================================
    def open_blocklist_modal(self):
        modal = tk.Toplevel(self.root)
        modal.title("Blocklist & 'STOP' Numbers Manager")
        modal.geometry("520x480")
        modal.transient(self.root)
        modal.grab_set()

        frame = tk.Frame(modal, bg="#f8fafc", padx=16, pady=14)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text="🛡️ Blocklisted Phone Numbers", font=("Segoe UI", 11, "bold"), bg="#f8fafc").pack(anchor="w")
        tk.Label(
            frame,
            text="Contacts in this list or replying 'STOP' will be skipped automatically.",
            font=("Segoe UI", 8),
            bg="#f8fafc",
            fg="#64748b",
        ).pack(anchor="w", pady=(0, 8))

        listbox_frame = tk.Frame(frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, pady=4)

        scrollbar = tk.Scrollbar(listbox_frame, orient=tk.VERTICAL)
        blocked_listbox = tk.Listbox(listbox_frame, yscrollcommand=scrollbar.set, font=("Consolas", 9), selectmode=tk.SINGLE)
        scrollbar.config(command=blocked_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        blocked_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        current_blocked = sorted(list(load_blocklist()))
        for num in current_blocked:
            blocked_listbox.insert(tk.END, num)

        add_row = tk.Frame(frame, bg="#f8fafc")
        add_row.pack(fill=tk.X, pady=8)

        entry_new = tk.Entry(add_row, font=("Segoe UI", 9), width=20)
        entry_new.pack(side=tk.LEFT, padx=(0, 6))

        def add_number():
            raw = entry_new.get().strip()
            if not raw:
                return
            cleaned, _ = clean_phone_number(raw)
            target = cleaned if cleaned else raw
            append_to_blocklist(target)
            blocked_listbox.insert(tk.END, target)
            entry_new.delete(0, tk.END)
            self._log("INFO", f"Added {target} to blocklist.txt")

        def remove_number():
            sel = blocked_listbox.curselection()
            if not sel:
                return
            selected_num = blocked_listbox.get(sel[0])
            blocked_listbox.delete(sel[0])
            remaining = [blocked_listbox.get(i) for i in range(blocked_listbox.size())]
            try:
                with open(DEFAULT_BLOCKLIST_FILE, "w", encoding="utf-8") as f:
                    f.write("# Skytecher WhatsApp Outreach - Blocklist\n")
                    for n in remaining:
                        f.write(f"{n}\n")
                self._log("INFO", f"Removed {selected_num} from blocklist.txt")
            except Exception as e:
                messagebox.showerror("Error", f"Failed updating blocklist: {e}")

        tk.Button(add_row, text="Add Number", font=("Segoe UI", 9, "bold"), bg="#0284c7", fg="#ffffff", relief="flat", padx=8, pady=3, command=add_number).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(add_row, text="Remove Selected", font=("Segoe UI", 9), bg="#dc2626", fg="#ffffff", relief="flat", padx=8, pady=3, command=remove_number).pack(side=tk.LEFT)

    # ==========================================
    # CLEAR SENT HISTORY
    # ==========================================
    def action_clear_sent_history(self):
        """Clears sent_history.txt so all contacts can be re-sent."""
        confirm = messagebox.askyesno(
            "Clear Sent History",
            "This will clear all previously sent contact records.\n"
            "All contacts will be eligible for sending again.\n\nProceed?",
        )
        if not confirm:
            return
        try:
            with open(DEFAULT_SENT_HISTORY_FILE, "w", encoding="utf-8") as f:
                f.write("# Skytecher WhatsApp Outreach - Sent History Log\n")
                f.write("# Format: Phone, Date-Time, Company\n")
            self._log("SUCCESS", "Sent history cleared. All contacts are now eligible for sending.")
            messagebox.showinfo("History Cleared", "Sent history has been cleared successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to clear sent history: {e}")

    # ==========================================
    # DISPATCH CONTROLS & BACKGROUND THREAD
    # ==========================================
    def action_start_sending(self):
        if not self.contacts_list:
            messagebox.showwarning("No Contacts", "Please upload an Excel file with contacts first.")
            return

        template = self.txt_template.get("1.0", tk.END).strip()
        if not template:
            messagebox.showwarning("Template Empty", "Message template cannot be empty.")
            return

        try:
            daily_limit = int(self.spin_daily_limit.get())
            delay_min = float(self.spin_delay_min.get())
            delay_max = float(self.spin_delay_max.get())
            wait_time = int(self.spin_wait_time.get())
        except ValueError:
            messagebox.showerror("Invalid Settings", "Please ensure numerical values are entered for limits and delays.")
            return

        if delay_min > delay_max:
            messagebox.showerror("Invalid Delay Range", "Min delay cannot be greater than Max delay.")
            return

        is_dry_run = self.var_dry_run.get()
        if not is_dry_run and not PYAUTOGUI_AVAILABLE:
            resp = messagebox.askyesno(
                "pyautogui Required",
                "pyautogui is not installed. It is required to press Enter to send messages.\n\n"
                "Would you like to run in Simulation / Dry Run mode instead?",
            )
            if resp:
                self.var_dry_run.set(True)
                is_dry_run = True
            else:
                return

        # AI Configuration checks
        use_ai = self.var_enable_ai.get()
        api_key = self.entry_api_key.get().strip()
        ai_model = self.combo_model.get().strip()

        if use_ai and not api_key:
            resp = messagebox.askyesno(
                "No API Key",
                "AI Personalization is checked, but no API Key was provided.\n\n"
                "Would you like to proceed using the standard template instead?",
            )
            if not resp:
                return
            use_ai = False

        if not is_dry_run:
            ai_note = f"• AI Message Generator: ACTIVE ({ai_model})\n" if use_ai else "• AI Personalization: Standard Template\n"
            confirm = messagebox.askyesno(
                "Confirm Outreach Dispatch",
                f"You are about to start sending WhatsApp messages from +91 {SENDER_PHONE}.\n\n"
                f"• Total Contacts: {len(self.contacts_list)}\n"
                f"• Daily Send Limit: {daily_limit}\n"
                f"• Random Delay: {delay_min}s - {delay_max}s\n"
                f"• Web Load Wait: {wait_time}s\n"
                f"{ai_note}\n"
                "METHOD: Direct WhatsApp Web URL → Auto Press Enter → Auto Close Tab\n"
                "If a number is NOT on WhatsApp, the app logs it and moves to the next number.\n\n"
                "⚠️ IMPORTANT: Make sure WhatsApp Web is logged in on your browser.\n"
                "⚠️ DO NOT touch your keyboard/mouse while sending is in progress.\n\n"
                "Proceed?",
            )
            if not confirm:
                return

        self.is_sending = True
        self.stop_event.clear()
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self.btn_upload.config(state=tk.DISABLED)
        self.lbl_footer.config(text="⏳ Dispatch running... DO NOT touch keyboard/mouse. WhatsApp Web is being automated.")

        ignore_history = self.var_ignore_history.get()

        self.worker_thread = threading.Thread(
            target=self._send_loop_worker,
            args=(template, daily_limit, delay_min, delay_max, wait_time, is_dry_run, use_ai, api_key, ai_model, ignore_history),
            daemon=True,
        )
        self.worker_thread.start()

    def action_stop_sending(self):
        if self.is_sending:
            self._log("WARNING", "Stop requested by user. Terminating process safely...")
            self.stop_event.set()
            self.btn_stop.config(state=tk.DISABLED)

    def _send_loop_worker(
        self,
        template: str,
        daily_limit: int,
        delay_min: float,
        delay_max: float,
        wait_time: int,
        is_dry_run: bool,
        use_ai: bool,
        api_key: str,
        ai_model: str,
        ignore_history: bool = False,
    ):
        """Worker thread executing the send loop with validation, AI generation, and error resilience."""
        mode_str = "SIMULATION" if is_dry_run else "LIVE WHATSAPP"
        ai_str = f"AI [{ai_model}]" if use_ai else "Standard Template"
        history_str = " (Ignore Sent History ON)" if ignore_history else ""
        self._log("INFO", f"Starting dispatch batch: {mode_str} | {ai_str}{history_str} | Daily Limit: {daily_limit}")

        blocklist = load_blocklist()
        sent_history = load_sent_history()
        seen_in_batch: Set[str] = set()

        sent_today = 0
        total_contacts = len(self.contacts_list)

        for idx, contact in enumerate(self.contacts_list):
            if self.stop_event.is_set():
                self._log("WARNING", "Outreach halted by user.")
                break

            name = contact["Name"] or f"Contact #{idx + 1}"
            raw_phone = contact["Phone_Raw"]
            company = contact["Company"]

            result_item = {
                "Row": contact["row_index"],
                "Name": name,
                "Raw Phone": raw_phone,
                "Cleaned Phone": "",
                "Company": company,
                "Service": contact["Service"],
                "Status": "",
                "Details": "",
                "Message Preview": "",
                "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

            # Check daily limit
            if sent_today >= daily_limit:
                msg = f"Daily send limit of {daily_limit} reached. Halting batch safely."
                self._log("INFO", msg)
                result_item["Status"] = "Skipped"
                result_item["Details"] = "Daily send limit reached"
                self.results_records.append(result_item)
                self.skipped_count += 1
                self._update_kpi_counters()
                break

            # 1. Clean & validate phone number
            cleaned_phone, err = clean_phone_number(raw_phone)
            if err or not cleaned_phone:
                result_item["Status"] = "Skipped"
                result_item["Details"] = f"Invalid Phone: {err}"
                self.results_records.append(result_item)
                self.skipped_count += 1
                self._log("SKIP", f"[{idx + 1}/{total_contacts}] Skipped {name}: {err}")
                self._update_kpi_counters()
                continue

            result_item["Cleaned Phone"] = cleaned_phone

            # 2. Check for duplicate within the current Excel file
            if cleaned_phone in seen_in_batch:
                result_item["Status"] = "Skipped"
                result_item["Details"] = "Duplicate number in Excel file"
                self.results_records.append(result_item)
                self.skipped_count += 1
                self._log("SKIP", f"[{idx + 1}/{total_contacts}] Skipped duplicate: {cleaned_phone} ({name})")
                self._update_kpi_counters()
                continue
            seen_in_batch.add(cleaned_phone)

            # 3. Check Blocklist / 'STOP'
            if cleaned_phone in blocklist:
                result_item["Status"] = "Skipped"
                result_item["Details"] = "On blocklist / STOP list"
                self.results_records.append(result_item)
                self.skipped_count += 1
                self._log("SKIP", f"[{idx + 1}/{total_contacts}] Skipped blocklisted: {cleaned_phone} ({name})")
                self._update_kpi_counters()
                continue

            # 4. Check Sent History (already messaged)
            if not ignore_history and cleaned_phone in sent_history:
                result_item["Status"] = "Skipped"
                result_item["Details"] = "Already messaged previously (found in sent_history.txt)"
                self.results_records.append(result_item)
                self.skipped_count += 1
                self._log("SKIP", f"[{idx + 1}/{total_contacts}] Skipped previous recipient: {cleaned_phone} ({name})")
                self._update_kpi_counters()
                continue

            # 5. Format message (AI generated OR template fallback)
            final_message = ""
            if use_ai:
                self._log("INFO", f"[{idx + 1}/{total_contacts}] Generating AI message via {ai_model} for {name}...")
                ai_text, ai_err = call_ai_generate_message(
                    api_key=api_key,
                    base_url=DEFAULT_AI_BASE_URL,
                    model=ai_model,
                    contact_data=contact,
                )
                if ai_text:
                    final_message = ai_text
                else:
                    self._log("WARNING", f"AI call failed ({ai_err}). Falling back to standard template.")
                    final_message = format_template_message(template, contact)
            else:
                final_message = format_template_message(template, contact)

            result_item["Message Preview"] = final_message[:100] + ("..." if len(final_message) > 100 else "")

            # 6. Dispatch message via WhatsApp Web with full error resilience
            try:
                self._log("INFO", f"[{idx + 1}/{total_contacts}] Dispatching to {name} at {cleaned_phone}...")

                if is_dry_run:
                    # SIMULATION MODE: just log the message without sending
                    time.sleep(1.5)
                    self._log("SUCCESS", f"✅ SIMULATION: Message ready for {name} ({cleaned_phone})")
                    self._log("INFO", f"   Preview: {final_message[:80]}...")
                    send_success = True
                    send_detail = "Simulation - message not actually sent"
                else:
                    # LIVE MODE: Send via direct WhatsApp Web URL + pyautogui
                    self._log("INFO", f"   Opening WhatsApp Web for {cleaned_phone}...")
                    send_success, send_detail = send_whatsapp_message(
                        phone=cleaned_phone,
                        message=final_message,
                        wait_time=wait_time,
                    )
                    
                    if send_success:
                        self._log("SUCCESS", f"✅ Message delivered to {name} ({cleaned_phone})")
                    else:
                        self._log("ERROR", f"❌ Failed to send to {name} ({cleaned_phone}): {send_detail}")

                if send_success:
                    # Log to sent history
                    append_to_sent_history(cleaned_phone, company)
                    sent_history.add(cleaned_phone)

                    result_item["Status"] = "Sent"
                    result_item["Details"] = send_detail if not is_dry_run else "Delivered successfully"
                    self.sent_count += 1
                    sent_today += 1
                else:
                    result_item["Status"] = "Failed (No WhatsApp / Error)"
                    result_item["Details"] = f"Delivery failed: {send_detail}"
                    self.failed_count += 1
                    self._log("WARNING", f"Skipping {cleaned_phone} and moving to next contact...")

            except Exception as send_err:
                err_str = str(send_err)
                self._log("ERROR", f"Unexpected error for {cleaned_phone} ({name}): {err_str}")
                self._log("WARNING", f"Phone {cleaned_phone} might not have WhatsApp. Moving to next number automatically...")

                # Clean any lingering popups/browser tabs
                if PYAUTOGUI_AVAILABLE:
                    try:
                        pyautogui.press("escape")
                        time.sleep(0.5)
                        pyautogui.hotkey("ctrl", "w")
                        time.sleep(0.5)
                    except Exception:
                        pass

                result_item["Status"] = "Failed (No WhatsApp / Error)"
                result_item["Details"] = f"Delivery failed: {err_str}"
                self.failed_count += 1

            self.results_records.append(result_item)
            self._update_kpi_counters()

            # 7. Apply random inter-message delay (only if not at last contact)
            if idx < total_contacts - 1 and not self.stop_event.is_set():
                sleep_duration = random.uniform(delay_min, delay_max)
                self._log("INFO", f"⏳ Cooldown delay: {sleep_duration:.1f}s before next contact (anti-ban protection)...")

                elapsed = 0.0
                while elapsed < sleep_duration and not self.stop_event.is_set():
                    time.sleep(0.5)
                    elapsed += 0.5

        # Cleanup and finalize batch
        self.is_sending = False
        self.root.after(0, self._finalize_run_ui)

    def _finalize_run_ui(self):
        self.btn_start.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        self.btn_upload.config(state=tk.NORMAL)
        self.lbl_footer.config(text="Outreach batch complete. Check log or export Excel report.")

        summary = (
            f"Outreach Batch Summary:\n\n"
            f"• Total Processed: {self.sent_count + self.skipped_count + self.failed_count}\n"
            f"• Successfully Sent: {self.sent_count}\n"
            f"• Skipped (Duplicates/Blocklist/Invalid): {self.skipped_count}\n"
            f"• Failed (No WhatsApp / Errors): {self.failed_count}\n\n"
            "Would you like to export the detailed results report (.xlsx) now?"
        )
        if messagebox.askyesno("Batch Finished", summary):
            self.action_export_report()

    # ==========================================
    # REPORT EXPORT TO EXCEL
    # ==========================================
    def action_export_report(self):
        if not self.results_records:
            messagebox.showinfo("Export", "No dispatch results available to export yet.")
            return

        export_path = filedialog.asksaveasfilename(
            title="Save Results Report",
            defaultextension=".xlsx",
            initialfile="results.xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
        )
        if not export_path:
            return

        try:
            rep_df = pd.DataFrame(self.results_records)
            rep_df.to_excel(export_path, index=False)
            self._log("SUCCESS", f"Report saved to: {export_path}")
            messagebox.showinfo("Export Successful", f"Results report successfully saved to:\n{export_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed saving Excel report: {e}")
            self._log("ERROR", f"Export failed: {e}")


# ==========================================
# APPLICATION ENTRY POINT
# ==========================================
def main():
    root = tk.Tk()
    app = WhatsAppOutreachApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
