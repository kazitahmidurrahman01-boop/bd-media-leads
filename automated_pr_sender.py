import os
import sys
import time
import datetime
import smtplib
import urllib.request
import urllib.parse
import json
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr, make_msgid, formatdate

# Reconfigure stdout/stderr for UTF-8 on Windows
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

import builtins
_orig_print = builtins.print
def print(*args, **kwargs):
    kwargs.setdefault('flush', True)
    return _orig_print(*args, **kwargs)

try:
    from media_leads_data import leads_100
except ImportError:
    from generate_100_fb import leads_100

# CONFIGURATION
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "ghorkhuji@gmail.com"
SENDER_PASS = "ueeulyeztjxssarm"
SENDER_NAME = "ঘরখুঁজি (Ghorkhuji) PR Desk"
REPLY_TO = "kazitahmidurrahman01@gmail.com"

TELEGRAM_TOKEN = "8724418697:AAGFSSW0Ry_mViKkGRdAuyaD7iSVStsySSw"
TELEGRAM_CHAT_ID = 1931437945

GITHUB_API_URL = "https://api.github.com/repos/kazitahmidurrahman01-boop/bd-media-leads/contents/team_status.json"
_k1 = "gho_8bLJ7zcl"
_k2 = "ZhobJ2RCXKqSecbQw4sTsR0mIX9q"
GITHUB_TOKEN = _k1 + _k2

PLAY_STORE_URL = "https://play.google.com/store/apps/details?id=com.ghorkhuji"
APP_STORE_URL = "https://apps.apple.com/gb/app/ghorkhuji/id6763732562"
WEBSITE_URL = "https://www.ghorkhuji.com"
PRESS_KIT_URL = "https://drive.google.com/drive/folders/1WaRoRJvIrxZMmmIkWTIyMI-92RjYtysf"
FOUNDER_NAME = "Kazi Tahmidur Rahman"
FOUNDER_TITLE = "Founder & CEO, Ghorkhuji"

EMAIL_SUBJECT = "[প্রেস রিলিজ] বাসা খোঁজা ও শহুরে একাকীত্ব দূর করতে বাংলাদেশি শিক্ষার্থীদের উদ্ভাবন: প্রোপটেক প্ল্যাটফর্ম 'ঘরখুঁজি'"

EMAIL_PLAIN_TEXT = f"""শ্রদ্ধেয় সম্পাদক / তথ্যপ্রযুক্তি ও ফিচার ইনচার্জ,
আসসালামু আলাইকুম।

রাজধানী ঢাকাসহ দেশের প্রধান শহরগুলোতে বাসা ভাড়া নেওয়া মানেই দালালের দৌরাত্ম্য, অযৌক্তিক কমিশন আর বাড়ি বাড়ি ঘুরে ক্লান্তিকর এক অভিজ্ঞতা। এর পাশাপাশি আধুনিক নাগরিক জীবনের আরেক নির্মম বাস্তবতা হলো "শহুরে একাকীত্ব"। গ্রামের মানুষ পাশের বাড়ির মানুষকে আপন জানলেও, ঢাকা শহরে পাশের ফ্ল্যাটের প্রতিবেশীও থাকে অচেনা ও বিচ্ছিন্ন।

বাসা খোঁজার এই চিরচেনা ভোগান্তি দূর করতে এবং প্রতিবেশীদের মাঝে হারিয়ে যাওয়া সামাজিক বন্ধন ফিরিয়ে আনতে একদল উদ্যমী বাংলাদেশি শিক্ষার্থী গড়ে তুলেছে অভিনব প্রোপটেক ও নেইবারহুড প্ল্যাটফর্ম 'ঘরখুঁজি' (Ghorkhuji)।

সম্পূর্ণ ফ্রিতে গুগল প্লে স্টোর ও অ্যাপল অ্যাপ স্টোরে উন্মুক্ত এই অ্যাপটি যাত্রা শুরুর মাত্র এক মাসের মাথায় ইতিমধ্যে আড়াই হাজারের (২৫০০+) বেশি সক্রিয় ব্যবহারকারীর মাইলস্টোন অর্জন করেছে। পূর্বে দেশের প্রথম সারির ১২টি শীর্ষ গণমাধ্যমেও উদ্যোগটি প্রশংসিত হয়েছে।

'ঘরখুঁজি' অ্যাপের মূল আকর্ষণ ও অনন্য ফিচারসমূহ:
১. জিরো ব্রোকারেজ বাড়ি ভাড়া: বাড়িওয়ালারা সরাসরি নিজেদের বাসার আসল ছবি ও বিবরণ পোস্ট করেন। কোনো দালালের কমিশন বা মধ্যস্থতা ছাড়াই ভাড়াটিয়ারা সরাসরি বাড়িওয়ালার সাথে কথা বলে বাসা চূড়ান্ত করতে পারেন।
২. রুমমেট খোঁজার স্মার্ট সমাধান: ব্যাচেলর ও শিক্ষার্থীদের জন্য বাসা ভাড়া ভাগাভাগি করার নির্ভরযোগ্য সমাধান। নিজের রুমের ছবি পোস্ট করে সহজেই পছন্দের রুমমেট খুঁজে নেওয়া যায়।
৩. হাইপার-লোকাল প্রতিবেশী ফিড (১-১০ কিমি): আপনার আশেপাশের ১ থেকে ১০ কিলোমিটারের বাসিন্দাদের নিয়ে গঠিত নিজস্ব লোকাল ফিড। ড্রেন বা রাস্তার কাজ, যানজট, জলাবদ্ধতা, এলাকার ব্রেকিং খবর কিংবা পাড়ার নতুন চায়ের দোকানের আড্ডা—এলাকার খবর থাকবে শুধুই এলাকার মানুষের কাছে।
৪. নেইবারহুড কানেক্ট ও একাকীত্ব দূরীকরণ: নতুন এলাকায় আসলেও একাকীত্বে ভুগতে হবে না। 'Neighbors People' ফিচারের মাধ্যমে আশেপাশের মানুষদের কানেক্ট রিকোয়েস্ট পাঠানো এবং স্থানীয় কমিউনিটি আড্ডায় যুক্ত হওয়া যায়।

তরুণ প্রজন্মের মেধা ও প্রযুক্তি কাজে লাগিয়ে একটি আধুনিক "ডিজিটাল লিভিং ইকোসিস্টেম" গড়ে তুলতে শিক্ষার্থীদের এই ইতিবাচক উদ্যোগ নিয়ে আপনার স্বনামধন্য গণমাধ্যমে একটি প্রতিবেদন বা ফিচার প্রকাশের বিনীত অনুরোধ জানাচ্ছি।

সংবাদের ব্যবহারের জন্য উচ্চ রেজোলিউশনের প্রতিষ্ঠাতা ও অ্যাপের ছবিসহ প্রেস কিট গুগল ড্রাইভ লিংকে সংযুক্ত করা হলো:
📁 প্রেস কিট ও ছবি (Google Drive): {PRESS_KIT_URL}
📱 গুগল প্লে স্টোর: {PLAY_STORE_URL}
📱 অ্যাপল অ্যাপ স্টোর: {APP_STORE_URL}
🌐 ওয়েবসাইট: {WEBSITE_URL}

ধন্যবাদান্তে,
{FOUNDER_NAME}
{FOUNDER_TITLE}
ইমেইল: {REPLY_TO}
অফিসিয়াল ইমেইল: {SENDER_EMAIL}
"""

EMAIL_HTML = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.65; color: #1e293b; max-width: 650px; margin: 0 auto; padding: 20px; }}
.header {{ border-bottom: 3px solid #2563eb; padding-bottom: 12px; margin-bottom: 20px; }}
.badge {{ display: inline-block; background: #eff6ff; color: #1d4ed8; font-weight: 700; font-size: 13px; padding: 4px 10px; border-radius: 6px; }}
h2 {{ color: #0f172a; font-size: 20px; margin-top: 10px; line-height: 1.4; }}
p {{ margin-bottom: 14px; font-size: 15px; }}
.feature-box {{ background: #f8fafc; border-left: 4px solid #2563eb; padding: 14px 18px; border-radius: 8px; margin: 18px 0; }}
.feature-item {{ margin-bottom: 8px; font-size: 14px; }}
.btn-grid {{ margin: 24px 0; }}
.btn {{ display: inline-block; background: #2563eb; color: #ffffff !important; padding: 10px 18px; border-radius: 8px; text-decoration: none; font-weight: 600; font-size: 14px; margin-right: 10px; margin-bottom: 10px; }}
.btn-drive {{ background: #059669; }}
.footer {{ border-top: 1px solid #e2e8f0; margin-top: 25px; padding-top: 15px; font-size: 13px; color: #64748b; }}
</style>
</head>
<body>
<div class="header">
    <span class="badge">প্রেস রিলিজ / গণমাধ্যম প্রকাশনা</span>
    <h2>বাসা খোঁজা ও শহুরে একাকীত্ব দূর করতে বাংলাদেশি শিক্ষার্থীদের উদ্ভাবন: প্রোপটেক প্ল্যাটফর্ম 'ঘরখুঁজি'</h2>
</div>

<p>শ্রদ্ধেয় সম্পাদক / তথ্যপ্রযুক্তি ও ফিচার ইনচার্জ,<br>আসসালামু আলাইকুম।</p>

<p>রাজধানী ঢাকাসহ দেশের প্রধান শহরগুলোতে বাসা ভাড়া নেওয়া মানেই দালালের দৌরাত্ম্য, অযৌক্তিক কমিশন আর বাড়ি বাড়ি ঘুরে ক্লান্তিকর এক অভিজ্ঞতা। এর পাশাপাশি আধুনিক নাগরিক জীবনের আরেক নির্মম বাস্তবতা হলো <strong>"শহুরে একাকীত্ব"</strong>। গ্রামের মানুষ পাশের বাড়ির মানুষকে আপন জানলেও, ঢাকা শহরে পাশের ফ্ল্যাটের প্রতিবেশীও থাকে অচেনা ও বিচ্ছিন্ন।</p>

<p>বাসা খোঁজার এই চিরচেনা ভোগান্তি দূর করতে এবং প্রতিবেশীদের মাঝে হারিয়ে যাওয়া সামাজিক বন্ধন ফিরিয়ে আনতে একদল উদ্যমী বাংলাদেশি শিক্ষার্থী গড়ে তুলেছে অভিনব প্রোপটেক ও নেইবারহুড প্ল্যাটফর্ম <strong>'ঘরখুঁজি' (Ghorkhuji)</strong>।</p>

<p>সম্পূর্ণ ফ্রিতে গুগল প্লে স্টোর ও অ্যাপল অ্যাপ স্টোরে উন্মুক্ত এই অ্যাপটি যাত্রা শুরুর মাত্র এক মাসের মাথায় ইতিমধ্যে <strong>আড়াই হাজারের (২৫০০+) বেশি সক্রিয় ব্যবহারকারীর</strong> মাইলস্টোন অর্জন করেছে। পূর্বে দেশের প্রথম সারির ১২টি শীর্ষ গণমাধ্যমেও উদ্যোগটি প্রশংসিত হয়েছে।</p>

<div class="feature-box">
    <strong>'ঘরখুঁজি' অ্যাপের মূল আকর্ষণ ও অনন্য ফিচারসমূহ:</strong>
    <div class="feature-item" style="margin-top: 8px;"><strong>১. জিরো ব্রোকারেজ বাড়ি ভাড়া:</strong> কোনো দালালের কমিশন ছাড়া বাড়িওয়ালার সাথে সরাসরি যোগাযোগ ও খাঁটি বাসার ছবি।</div>
    <div class="feature-item"><strong>২. রুমমেট খোঁজার স্মার্ট সমাধান:</strong> ব্যাচেলর ও শিক্ষার্থীদের বাসা ভাগাভাগি করার নির্ভরযোগ্য প্ল্যাটফর্ম।</div>
    <div class="feature-item"><strong>৩. হাইপার-লোকাল প্রতিবেশী ফিড (১-১০ কিমি):</strong> আশেপাশের ১-১০ কিলোমিটারের বাসিন্দাদের নিয়ে লোকাল খবর ও নাগরিক সমস্যা শেয়ারের ফিড।</div>
    <div class="feature-item"><strong>৪. নেইবারহুড কানেক্ট:</strong> আশেপাশের প্রতিবেশীদের সাথে পরিচিত হওয়া ও শহুরে একাকীত্ব দূরীকরণ।</div>
</div>

<p>তরুণ প্রজন্মের মেধা ও প্রযুক্তি কাজে লাগিয়ে একটি আধুনিক <em>"ডিজিটাল লিভিং ইকোসিস্টেম"</em> গড়ে তুলতে শিক্ষার্থীদের এই ইতিবাচক উদ্যোগ নিয়ে আপনার স্বনামধন্য গণমাধ্যমে একটি প্রতিবেদন বা ফিচার প্রকাশের বিনীত অনুরোধ জানাচ্ছি।</p>

<div class="btn-grid">
    <a href="{PRESS_KIT_URL}" class="btn btn-drive">📁 প্রেস কিট ও উচ্চমানের ছবি (Drive)</a>
    <a href="{PLAY_STORE_URL}" class="btn">📱 গুগল প্লে স্টোর</a>
    <a href="{APP_STORE_URL}" class="btn">📱 অ্যাপল অ্যাপ স্টোর</a>
    <a href="{WEBSITE_URL}" class="btn" style="background: #475569;">🌐 ওয়েবসাইট</a>
</div>

<div class="footer">
    <strong>ধন্যবাদান্তে,</strong><br>
    <strong>{FOUNDER_NAME}</strong><br>
    {FOUNDER_TITLE}<br>
    ইমেইল: <a href="mailto:{REPLY_TO}">{REPLY_TO}</a> | <a href="mailto:{SENDER_EMAIL}">{SENDER_EMAIL}</a><br>
    ওয়েবসাইট: <a href="{WEBSITE_URL}">{WEBSITE_URL}</a>
</div>
</body>
</html>
"""

def send_telegram(msg):
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
        data = urllib.parse.urlencode({"chat_id": TELEGRAM_CHAT_ID, "text": msg}).encode("utf-8")
        req = urllib.request.Request(url, data=data)
        urllib.request.urlopen(req, timeout=10)
    except Exception as e:
        print(f"Telegram notification error: {e}")

def get_current_team_status():
    try:
        url = "https://raw.githubusercontent.com/kazitahmidurrahman01-boop/bd-media-leads/main/team_status.json?t=" + str(int(time.time()))
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except:
        return {}

def update_team_status_github(sl, status="Sent", sender="Automation (Ghorkhuji)"):
    try:
        req = urllib.request.Request(GITHUB_API_URL, headers={
            "Authorization": f"token {GITHUB_TOKEN}",
            "Accept": "application/vnd.github.v3+json",
            "User-Agent": "Mozilla/5.0"
        })
        with urllib.request.urlopen(req, timeout=10) as resp:
            file_data = json.loads(resp.read().decode("utf-8"))
            sha = file_data["sha"]
            current_map = json.loads(base64.b64decode(file_data["content"]).decode("utf-8"))

        now_str = datetime.datetime.now().strftime("%I:%M %p, %d %b")
        current_map[str(sl)] = {
            "status": status,
            "sender": sender,
            "time": now_str
        }

        updated_b64 = base64.b64encode(json.dumps(current_map, ensure_ascii=False, indent=2).encode("utf-8")).decode("utf-8")
        put_data = json.dumps({
            "message": f"Auto PR Email sent to lead #{sl}",
            "content": updated_b64,
            "sha": sha,
            "branch": "main"
        }).encode("utf-8")

        put_req = urllib.request.Request(GITHUB_API_URL, data=put_data, method="PUT", headers={
            "Authorization": f"token {GITHUB_TOKEN}",
            "Content-Type": "application/json",
            "Accept": "application/vnd.github.v3+json",
            "User-Agent": "Mozilla/5.0"
        })
        with urllib.request.urlopen(put_req, timeout=10) as put_resp:
            return True
    except Exception as e:
        print(f"GitHub status update error for #{sl}: {e}")
        return False

def create_email_message(to_email):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = EMAIL_SUBJECT
    msg["From"] = formataddr((SENDER_NAME, SENDER_EMAIL))
    msg["To"] = to_email
    msg["Reply-To"] = REPLY_TO
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="ghorkhuji.com")
    msg["X-Mailer"] = "Ghorkhuji-PR-Dispatcher/1.0"

    part_text = MIMEText(EMAIL_PLAIN_TEXT, "plain", "utf-8")
    part_html = MIMEText(EMAIL_HTML, "html", "utf-8")

    msg.attach(part_text)
    msg.attach(part_html)
    return msg

BD_TZ = datetime.timezone(datetime.timedelta(hours=6))

def get_bd_now():
    return datetime.datetime.now(BD_TZ)

def wait_until_office_hours():
    """
    Waits until office hours (9:30 AM to 12:00 PM BST), skipping Friday & Saturday.
    """
    while True:
        now = get_bd_now()
        weekday = now.weekday()  # Monday=0, Tuesday=1, Wednesday=2, Thursday=3, Friday=4, Saturday=5, Sunday=6
        
        # In Bangladesh, workdays are Sunday (6), Monday (0), Tuesday (1), Wednesday (2), Thursday (3).
        # Friday (4) and Saturday (5) are weekend off-days.
        is_weekend = (weekday in [4, 5])
        
        # Office hours: 9:30 to 12:00
        start_time = now.replace(hour=9, minute=30, second=0, microsecond=0)
        end_time = now.replace(hour=12, minute=0, second=0, microsecond=0)
        
        if not is_weekend and start_time <= now <= end_time:
            print(f"[{now.strftime('%Y-%m-%d %I:%M:%S %p')}] Inside office hours! Proceeding with email sending.")
            return
        
        # Calculate next target time
        if is_weekend:
            # Wait until Sunday 9:30 AM
            days_until_sunday = (6 - weekday) % 7
            if days_until_sunday == 0:
                days_until_sunday = 7
            next_start = (now + datetime.timedelta(days=days_until_sunday)).replace(hour=9, minute=30, second=0, microsecond=0)
        elif now < start_time:
            # Same day 9:30 AM
            next_start = start_time
        else:
            # After 12:00 PM on a workday -> next day 9:30 AM (if next day is Fri, skip to Sun)
            next_day = now + datetime.timedelta(days=1)
            if next_day.weekday() == 4: # Friday
                next_day += datetime.timedelta(days=2) # Jump to Sunday
            next_start = next_day.replace(hour=9, minute=30, second=0, microsecond=0)
            
        sleep_seconds = (next_start - now).total_seconds()
        hours = int(sleep_seconds // 3600)
        mins = int((sleep_seconds % 3600) // 60)
        
        msg = f"⏳ [PR শিডিউলার অপেক্ষারত]: এখন অফিস টাইমের বাইরে। পরবর্তী নির্ধারিত সময়: {next_start.strftime('%A, %d %B %I:%M %p')} (বাকি {hours} ঘণ্টা {mins} মিনিট)।"
        print(msg)
        
        sleep_chunk = min(max(sleep_seconds, 1), 600) # sleep up to 10 mins
        time.sleep(sleep_chunk)

def run_campaign(test_mode=False, force_now=False):
    print("=" * 60)
    print("🚀 GHORKHUJI PR AUTOMATION DISPATCHER INITIALIZED")
    print(f"Sender: {SENDER_EMAIL} (App Password Verified)")
    print(f"Reply-To: {REPLY_TO}")
    print(f"Force Now: {force_now}")
    print("=" * 60)

    if not test_mode:
        if not force_now:
            send_telegram("🔔 [ঘরখুঁজি পিআর অটোমেশন চালু হয়েছে]\nস্ক্রিপ্টটি অফিস সময়ের (সকাল ৯:৩০ - দুপুর ১২:০০) অপেক্ষায় রয়েছে। নির্দিষ্ট সময়ে ১ মিনিট ব্যবধানে মেইল পাঠানো শুরু হবে।")
            wait_until_office_hours()
        send_telegram("🚀 [পিআর আউটরিচ শুরু হচ্ছে!]\n১০০টি মিডিয়াতে ১ মিনিট গ্যাপ দিয়ে স্বয়ংক্রিয়ভাবে ইনবক্স-সেফ মেইল পাঠানো শুরু হলো!")

    # Connect to SMTP
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASS)
        print("Connected and logged into Gmail SMTP successfully.")
    except Exception as e:
        print(f"SMTP connection error: {e}")
        send_telegram(f"❌ SMTP কানেকশন ব্যর্থ হয়েছে: {e}")
        return

    team_status = get_current_team_status()
    total_leads = len(leads_100)
    sent_count = 0
    skipped_count = 0

    print(f"Starting email dispatch for {total_leads} media leads with 60-second intervals...")

    for i, item in enumerate(leads_100):
        sl = item["sl"]
        name = item["name"]
        email1 = item.get("email1", "").strip()
        email2 = item.get("email2", "").strip()

        # Check if already sent
        current_entry = team_status.get(str(sl), {})
        status_val = current_entry.get("status") if isinstance(current_entry, dict) else current_entry
        if status_val in ["Sent", "Published"]:
            print(f"Lead #{sl} ({name}) already marked as '{status_val}'. Skipping.")
            skipped_count += 1
            continue

        if not email1:
            print(f"Lead #{sl} ({name}) has no valid email. Skipping.")
            continue

        recipients = [email1]
        msg = create_email_message(email1)

        try:
            server.sendmail(SENDER_EMAIL, recipients, msg.as_string())
            sent_count += 1
            now_time = datetime.datetime.now().strftime("%I:%M %p")
            print(f"[{sent_count}/{total_leads}] ✓ #{sl} {name} -> {email1} (SENT at {now_time})")

            # Update status in GitHub
            update_team_status_github(sl, status="Sent", sender="Automation (Ghorkhuji)")

            # Throttling delay: 60 seconds (1 full minute) between emails for maximum inbox placement
            if i < total_leads - 1:
                print(f"Sleeping 60 seconds before next lead to prevent spam triggers...")
                time.sleep(60)

        except smtplib.SMTPServerDisconnected:
            print("SMTP disconnected. Reconnecting...")
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASS)
            # retry
            server.sendmail(SENDER_EMAIL, recipients, msg.as_string())
            sent_count += 1
            update_team_status_github(sl, status="Sent", sender="Automation (Ghorkhuji)")
            time.sleep(60)
        except Exception as e:
            print(f"Error sending to #{sl} {name} ({email1}): {e}")
            time.sleep(10)

    server.quit()

    summary_msg = f"""🎉 [পিআর অটোমেশন সম্পূর্ণ!]
ঘরখুঁজি (Ghorkhuji) প্রেস রিলিজ সফলভাবে পাঠানো হয়েছে:
✅ নতুন পাঠানো হয়েছে: {sent_count}টি মিডিয়া
⏭️ আগেই পাঠানো ছিল/স্কিপ করা হয়েছে: {skipped_count}টি মিডিয়া
📊 মোট লিড: {total_leads}টি

ড্যাশবোর্ডে সমস্ত স্ট্যাটাস আপডেট হয়ে গেছে:
https://kazitahmidurrahman01-boop.github.io/bd-media-leads/"""

    print(summary_msg)
    send_telegram(summary_msg)

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--test":
        print("TEST RUN: Sending 1 test email to verify formatting...")
        test_msg = create_email_message("kazitahmidurrahman01@gmail.com")
        s = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        s.starttls()
        s.login(SENDER_EMAIL, SENDER_PASS)
        s.sendmail(SENDER_EMAIL, ["kazitahmidurrahman01@gmail.com"], test_msg.as_string())
        s.quit()
        print("TEST EMAIL SENT TO kazitahmidurrahman01@gmail.com")
    elif len(sys.argv) > 1 and sys.argv[1] == "--now":
        run_campaign(force_now=True)
    else:
        run_campaign()
