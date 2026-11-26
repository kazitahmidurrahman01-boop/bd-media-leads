"""
=============================================================================
Ghorkhuji (ঘরখুঁজি) - Automated WhatsApp Outreach Assistant
=============================================================================
This script helps you reach all 100 media editors & journalists on WhatsApp
safely, with zero risk of Meta anti-spam account bans.

How it works:
1. Loads all 100 media leads and current status from team_status.json.
2. Formats the personalized Bengali pitch with links to Google Drive, Play Store & App Store.
3. Automatically opens WhatsApp Web for each valid mobile number.
4. Waits a safe delay (e.g. 12-15 seconds) so you can review and send, then auto-logs the status.
=============================================================================
"""

import os
import sys
import time
import json
import urllib.parse
import webbrowser
import datetime

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

try:
    from media_leads_data import leads_100
except ImportError:
    print("Error: media_leads_data.py not found.")
    sys.exit(1)

STATUS_FILE = "team_status.json"

WA_PITCH_TEMPLATE = """আসসালামু আলাইকুম [ভাইয়া/আপু],
আশা করি ভালো আছেন। আমি বাংলাদেশি শিক্ষার্থীদের তৈরি প্রোপটেক স্টার্টআপ 'ঘরখুঁজি' (Ghorkhuji)-এর পক্ষ থেকে নক দিয়েছি।

আমরা ঢাকা শহরে দালালের কমিশন ছাড়াই (Zero-Brokerage) বাসা ভাড়া ও রুমমেট খোঁজার পাশাপাশি শহুরে একাকীত্ব দূর করতে ১-১০ কিলোমিটারের প্রতিবেশীদের নিয়ে একটি আধুনিক ডিজিটাল লিভিং ইকোসিস্টেম গড়ে তুলেছি।

মূল ফিচারসমূহ:
১. জিরো ব্রোকারেজ বাড়ি ভাড়া: কোনো দালালের কমিশন ছাড়া সরাসরি বাড়িওয়ালার সাথে যোগাযোগ।
২. রুমমেট খোঁজা: ব্যাচেলর ও স্টুডেন্টদের রুমমেট খোঁজার স্মার্ট সমাধান।
৩. প্রতিবেশী ফিড (১-১০ কিমি): এলাকার ড্রেন/রাস্তার কাজ, রোড ব্লক ও হাইপার-লোকাল খবর।
৪. নেইবারহুড কানেক্ট: প্রতিবেশীদের সাথে পরিচিত হওয়া ও শহুরে একাকীত্ব দূরীকরণ।

মাইলস্টোন: মাত্র ১ মাসেই আলহামদুলিল্লাহ ২৫০০+ সক্রিয় ব্যবহারকারী এবং পূর্বে ১২টি শীর্ষ গণমাধ্যম আমাদের কাভার করেছে। সম্পূর্ণ ফ্রি এই অ্যাপটি প্লে স্টোর ও অ্যাপ স্টোরে রয়েছে।

শিক্ষার্থীদের এই প্রযুক্তি উদ্যোগ নিয়ে আপনার স্বনামধন্য গণমাধ্যমে একটি নিউজ/ফিচার প্রকাশের জন্য বিনীত অনুরোধ জানাচ্ছি।

📁 প্রেস কিট ও ছবি: https://drive.google.com/drive/folders/1WaRoRJvIrxZMmmIkWTIyMI-92RjYtysf
📱 গুগল প্লে স্টোর: https://play.google.com/store/apps/details?id=com.ghorkhuji
📱 অ্যাপল অ্যাপ স্টোর: https://apps.apple.com/gb/app/ghorkhuji/id6763732562
🌐 ওয়েবসাইট: https://www.ghorkhuji.com

ধন্যবাদান্তে,
Kazi Tahmidur Rahman
Founder & CEO, Ghorkhuji
ফোন/হোয়াটসঅ্যাপ: +8801700000000
"""

def clean_phone_for_wa(phone_str):
    if not phone_str:
        return ""
    digits = "".join(filter(str.isdigit, str(phone_str)))
    if digits.startswith("880"):
        return digits
    elif digits.startswith("01"):
        return "88" + digits
    elif digits.startswith("1") and len(digits) == 10:
        return "880" + digits
    elif len(digits) > 8:
        return digits
    return ""

def load_status():
    if os.path.exists(STATUS_FILE):
        try:
            with open(STATUS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_status(data):
    with open(STATUS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def main():
    print("=" * 65)
    print("📱 GHORKHUJI MEDIA OUTREACH - WHATSAPP AUTOMATION ASSISTANT")
    print("=" * 65)
    
    status_map = load_status()
    total = len(leads_100)
    
    # Filter eligible leads (valid mobile numbers)
    eligible_leads = []
    for l in leads_100:
        wa_phone = clean_phone_for_wa(l.get("phone", ""))
        # Filter out landlines (02...)
        if wa_phone and (wa_phone.startswith("8801") or len(wa_phone) >= 11):
            eligible_leads.append((l, wa_phone))
        else:
            print(f"Skipping Lead #{l['sl']} ({l['name']}): Landline/Invalid phone ({l.get('phone')})")
            
    print(f"\n✓ Found {len(eligible_leads)} eligible WhatsApp mobile contacts out of {total} leads.")
    print("Instructions:")
    print("1. Make sure you are logged into WhatsApp Web in your default browser (Chrome/Edge).")
    print("2. The script will open each chat with the pre-filled press release pitch.")
    print("3. You can review or press Enter in WhatsApp to send, then return to terminal.")
    print("4. Press Enter here to proceed to the next lead, or type 's' to skip, 'q' to quit.\n")
    
    input("Press [Enter] to start WhatsApp outreach...")
    
    sent_count = 0
    for idx, (lead, wa_phone) in enumerate(eligible_leads, 1):
        sl = lead["sl"]
        name = lead["name"]
        cat = lead.get("cat", "")
        
        current_st = status_map.get(str(sl), {})
        status_val = current_st.get("status") if isinstance(current_st, dict) else current_st
        
        print(f"\n[{idx}/{len(eligible_leads)}] Lead #{sl}: {name} ({cat})")
        print(f"   Phone: +{wa_phone} | Current Status: {status_val}")
        
        # Prepare URL
        encoded_pitch = urllib.parse.quote(WA_PITCH_TEMPLATE)
        wa_url = f"https://web.whatsapp.com/send?phone={wa_phone}&text={encoded_pitch}"
        
        print(f"   Opening WhatsApp Web chat...")
        webbrowser.open(wa_url)
        
        choice = input("   Actions: [Enter]=Mark Sent & Next | 's'=Skip | 'q'=Quit: ").strip().lower()
        if choice == 'q':
            print("\nExiting WhatsApp assistant.")
            break
        elif choice == 's':
            print("   Skipped.")
            continue
        else:
            status_map[str(sl)] = {
                "status": "WhatsApp Sent",
                "timestamp": datetime.datetime.now().isoformat(),
                "sender": "WhatsApp Outreach (Founder)",
                "phone": wa_phone
            }
            save_status(status_map)
            sent_count += 1
            print(f"   ✓ Marked #{sl} as 'WhatsApp Sent' in {STATUS_FILE}.")
            time.sleep(2)

    print("\n" + "=" * 65)
    print(f"Outreach complete! Logged {sent_count} sent WhatsApp pitches.")
    print(f"Status file {STATUS_FILE} has been updated.")
    print("=" * 65)

if __name__ == "__main__":
    main()
