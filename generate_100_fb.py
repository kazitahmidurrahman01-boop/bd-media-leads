import os
import json
import csv
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    openpyxl = None

# 100 Bangladesh Media Outlets with Facebook Links
leads_100 = [
    # ১. স্টার্টআপ, টেক ও বিজনেস মিডিয়া (১-১৫)
    {"sl": 1, "cat": "স্টার্টআপ ও বিজনেস", "name": "Future Startup (FS)", "desk": "Startup Story / Founder Feature", "email1": "info@futurestartup.com", "email2": "editor@futurestartup.com", "phone": "+8801777683111", "display_phone": "+8801777-683111", "web": "https://futurestartup.com", "fb": "https://facebook.com/futurestartup", "notes": "বাংলাদেশের স্টার্টআপদের জন্য #১ প্ল্যাটফর্ম। ফাউন্ডার স্টোরি ও ইন-ডেপথ ইন্টারভিউ করে।"},
    {"sl": 2, "cat": "স্টার্টআপ ও বিজনেস", "name": "The Business Standard (TBS)", "desk": "Tech & Startup Desk", "email1": "news@tbsnews.net", "email2": "tbsitdesk@gmail.com", "phone": "+8801844556677", "display_phone": "+8801844-556677", "web": "https://tbsnews.net", "fb": "https://facebook.com/tbsnews.net", "notes": "স্টার্টআপ ফিচার ও গ্রোথ স্টোরি নিয়মিত প্রকাশ করে।"},
    {"sl": 3, "cat": "স্টার্টআপ ও বিজনেস", "name": "বণিক বার্তা (Bonik Barta)", "desk": "তথ্যপ্রযুক্তি ও করপোরেট", "email1": "news@bonikbarta.com", "email2": "bonikbarta.it@gmail.com", "phone": "+8801713388200", "display_phone": "+8801713-388200", "web": "https://bonikbarta.net", "fb": "https://facebook.com/bonikbarta", "notes": "বাংলাদেশের শীর্ষ বিজনেস দৈনিক। ইকনোমি ও ইনোভেশন রিপোর্ট করে।"},
    {"sl": 4, "cat": "স্টার্টআপ ও বিজনেস", "name": "The Daily Star (Star Bytes)", "desk": "Tech & Youth Desk", "email1": "starbytes@thedailystar.net", "email2": "business@thedailystar.net", "phone": "+8801711594580", "display_phone": "+8801711-594580", "web": "https://thedailystar.net", "fb": "https://facebook.com/thedailystar", "notes": "তরুণ উদ্যোক্তা ও টেক স্টার্টআপ ফিচারের জন্য সবচেয়ে প্রভাবশালী।"},
    {"sl": 5, "cat": "স্টার্টআপ ও বিজনেস", "name": "Dhaka Tribune", "desk": "Business & Tech Desk", "email1": "business@dhakatribune.com", "email2": "news@dhakatribune.com", "phone": "+8801755667788", "display_phone": "+8801755-667788", "web": "https://dhakatribune.com", "fb": "https://facebook.com/DhakaTribune", "notes": "ইনোভেশন, ডিজিটাল সার্ভিস ও তরুণদের ভেঞ্চার বান্ধব।"},
    {"sl": 6, "cat": "স্টার্টআপ ও বিজনেস", "name": "টেকশহর (Techshohor)", "desk": "টেক নিউজ ও স্টার্টআপ", "email1": "news@techshohor.com", "email2": "editor@techshohor.com", "phone": "+8801711235478", "display_phone": "+8801711-235478", "web": "https://techshohor.com", "fb": "https://facebook.com/Techshohor", "notes": "বিশেষায়িত আইটি, মোবাইল অ্যাপ ও স্টার্টআপ পোর্টাল।"},
    {"sl": 7, "cat": "স্টার্টআপ ও বিজনেস", "name": "ডিজিবাংলা (DigiBangla)", "desk": "আইটি ও উদ্ভাবন", "email1": "news@digibangla.tech", "email2": "info@digibangla.tech", "phone": "+8801977889900", "display_phone": "+8801977-889900", "web": "https://digibangla.tech", "fb": "https://facebook.com/DigiBangla.tech", "notes": "তথ্যপ্রযুক্তি ও স্টার্টআপ ভিত্তিক অনলাইন পোর্টাল।"},
    {"sl": 8, "cat": "স্টার্টআপ ও বিজনেস", "name": "The Financial Express (FE)", "desk": "Business & Innovation Desk", "email1": "fe@fe-bd.com", "email2": "editor@thefinancialexpress.com.bd", "phone": "+8801711823901", "display_phone": "+8801711-823901", "web": "https://thefinancialexpress.com.bd", "fb": "https://facebook.com/thefinancialexpress.com.bd", "notes": "ফাইন্যান্সিয়াল ও রিয়েল এস্টেট অ্যানালাইসিস।"},
    {"sl": 9, "cat": "স্টার্টআপ ও বিজনেস", "name": "Business Insider Bangladesh", "desk": "Startup & Tech", "email1": "news@businessinsiderbd.com", "email2": "editorial@businessinsiderbd.com", "phone": "+8801819214365", "display_phone": "+8801819-214365", "web": "https://businessinsiderbd.com", "fb": "https://facebook.com/businessinsiderbd", "notes": "বিজনেস ইনোভেশন ও উদ্যোক্তাদের সাক্ষাৎকার।"},
    {"sl": 10, "cat": "স্টার্টআপ ও বিজনেস", "name": "ICE Business Times", "desk": "Entrepreneurship Feature", "email1": "info@ibtbd.net", "email2": "editor@ibtbd.net", "phone": "+88028833971", "display_phone": "+8802-8833971", "web": "https://ibtbd.net", "fb": "https://facebook.com/ibtbd", "notes": "প্রিমিয়াম বিজনেস ম্যাগাজিন ও লিডারশিপ স্টোরি।"},
    {"sl": 11, "cat": "স্টার্টআপ ও বিজনেস", "name": "অর্থসূচক (Arthosuchak)", "desk": "বিজনেস ও পুঁজিবাজার", "email1": "news@arthosuchak.com", "email2": "arthosuchak@gmail.com", "phone": "+8801713456789", "display_phone": "+8801713-456789", "web": "https://arthosuchak.com", "fb": "https://facebook.com/arthosuchak", "notes": "দেশের প্রথম পূর্ণাঙ্গ অনলাইন বিজনেস পোর্টাল।"},
    {"sl": 12, "cat": "স্টার্টআপ ও বিজনেস", "name": "শেয়ার বিজ (Share Biz)", "desk": "বাণিজ্য ও কর্পোরেট", "email1": "news@sharebiz.net", "email2": "sharebiz.bd@gmail.com", "phone": "+8801711909876", "display_phone": "+8801711-909876", "web": "https://sharebiz.net", "fb": "https://facebook.com/sharebiz.net", "notes": "বিজনেস ও করপোরেট নিউজ দৈনিক পত্রিকা।"},
    {"sl": 13, "cat": "স্টার্টআপ ও বিজনেস", "name": "হাইফাই পাবলিক (HiFi Public)", "desk": "টেক ও ডিজিটাল কালচার", "email1": "contact@hifipublic.com", "email2": "editorial@hifipublic.com", "phone": "+8801712341234", "display_phone": "+8801712-341234", "web": "https://hifipublic.com", "fb": "https://facebook.com/hifipublic", "notes": "ডিজিটাল লাইফস্টাইল ও টেক স্টার্টআপের আর্টিক্যাল করে।"},
    {"sl": 14, "cat": "স্টার্টআপ ও বিজনেস", "name": "Telecom Reporters Network (TRNB)", "desk": "টেলিকম ও আইসিটি", "email1": "trnb.bd@gmail.com", "email2": "info@trnb.org.bd", "phone": "+8801711786543", "display_phone": "+8801711-786543", "web": "https://trnb.org.bd", "fb": "https://facebook.com/TRNBOfficial", "notes": "টেলিকম ও তথ্যপ্রযুক্তি সাংবাদিকদের কেন্দ্রীয় ফোরাম।"},
    {"sl": 15, "cat": "স্টার্টআপ ও বিজনেস", "name": "বিজনেস আই (The Business Eye)", "desk": "Startup & Property Desk", "email1": "editor@businesseye.com.bd", "email2": "info@businesseye.com.bd", "phone": "+8801819332244", "display_phone": "+8801819-332244", "web": "https://businesseye.com.bd", "fb": "https://facebook.com/businesseye.com.bd", "notes": "রিয়েল এস্টেট ও স্টার্টআপ বাণিজ্য কেন্দ্রিক ম্যাগাজিন।"},

    # ২. শীর্ষ জাতীয় বাংলা দৈনিক (১৬-৪৫)
    {"sl": 16, "cat": "জাতীয় বাংলা দৈনিক", "name": "প্রথম আলো (Prothom Alo)", "desk": "বিজ্ঞান ও প্রযুক্তি / তারুণ্য", "email1": "tech@prothomalo.com", "email2": "online@prothomalo.com", "phone": "+8801711006677", "display_phone": "+8801711-006677", "web": "https://prothomalo.com", "fb": "https://facebook.com/DailyProthomAlo", "notes": "দেশের ১ নম্বর গণমাধ্যম। টেক ও চাকরি-বাকরি পাতায় স্টার্টআপ কাভার করে।"},
    {"sl": 17, "cat": "জাতীয় বাংলা দৈনিক", "name": "কালের কণ্ঠ (Kaler Kantho)", "desk": "তথ্যপ্রযুক্তি পাতা", "email1": "kalerkantho.it@gmail.com", "email2": "news@kalerkantho.com", "phone": "+8801713247500", "display_phone": "+8801713-247500", "web": "https://kalerkantho.com", "fb": "https://facebook.com/kalerkantho", "notes": "আইটি ও টেক সেকশনে তরুণদের অ্যাপ/সার্ভিস গুরুত্ব পায়।"},
    {"sl": 18, "cat": "জাতীয় বাংলা দৈনিক", "name": "সমকাল (Samakal)", "desk": "টেলিকম ও প্রযুক্তি / ফিচার", "email1": "samakalitdesk@gmail.com", "email2": "news@samakal.com", "phone": "+8801911300400", "display_phone": "+8801911-300400", "web": "https://samakal.com", "fb": "https://facebook.com/TheDailySamakal", "notes": "উদ্যোক্তা ও প্রযুক্তি সেকশনে নিয়মিত ফিচার প্রকাশ করে।"},
    {"sl": 19, "cat": "জাতীয় বাংলা দৈনিক", "name": "যুগান্তর (Jugantor)", "desk": "আইটি কর্নার / তারুণ্যের মেলা", "email1": "itjugantor@gmail.com", "email2": "jugantor.online@gmail.com", "phone": "+8801713247600", "display_phone": "+8801713-247600", "web": "https://jugantor.com", "fb": "https://facebook.com/DailyJugantor", "notes": "আইটি কর্নার পাতায় নতুন উদ্ভাবন সহজে স্থান পায়।"},
    {"sl": 20, "cat": "জাতীয় বাংলা দৈনিক", "name": "ইত্তেফাক (Daily Ittefaq)", "desk": "বিজ্ঞান ও প্রযুক্তি", "email1": "ittefaqonline@gmail.com", "email2": "news@ittefaq.com.bd", "phone": "+8801711546000", "display_phone": "+8801711-546000", "web": "https://ittefaq.com.bd", "fb": "https://facebook.com/ittefaq.com.bd", "notes": "ঐতিহ্যবাহী জাতীয় দৈনিক, তথ্যপ্রযুক্তি পাতায় প্রকাশ করে।"},
    {"sl": 21, "cat": "জাতীয় বাংলা দৈনিক", "name": "বাংলাদেশ প্রতিদিন (BD Pratidin)", "desk": "প্রযুক্তি ও উদ্যোক্তা", "email1": "bdpratidin@gmail.com", "email2": "news@bd-pratidin.com", "phone": "+8801713247555", "display_phone": "+8801713-247555", "web": "https://bd-pratidin.com", "fb": "https://facebook.com/bdpratidin", "notes": "সর্বোচ্চ সার্কুলেশন প্রিন্ট পত্রিকা।"},
    {"sl": 22, "cat": "জাতীয় বাংলা দৈনিক", "name": "মানবজমিন (Manab Zamin)", "desk": "টেক ও লাইফস্টাইল", "email1": "mzamin.news@gmail.com", "email2": "online@mzamin.com", "phone": "+8801711667788", "display_phone": "+8801711-667788", "web": "https://mzamin.com", "fb": "https://facebook.com/dailymanabzamin", "notes": "জনপ্রিয় ট্যাবলয়েড পত্রিকা ও ওয়েব পোর্টাল।"},
    {"sl": 23, "cat": "জাতীয় বাংলা দৈনিক", "name": "দেশ রূপান্তর (Desh Rupantor)", "desk": "প্রাত্যহিক ও প্রযুক্তি", "email1": "deshrupantor.news@gmail.com", "email2": "news@deshrupantor.com", "phone": "+8801700705500", "display_phone": "+8801700-705500", "web": "https://deshrupantor.com", "fb": "https://facebook.com/DeshRupantorOfficial", "notes": "তরুণ পাঠকদের মধ্যে দ্রুত সমাদৃত আধুনিক পত্রিকা।"},
    {"sl": 24, "cat": "জাতীয় বাংলা দৈনিক", "name": "আমাদের সময় (Amader Shomoy)", "desk": "ফিচার ও তথ্যপ্রযুক্তি", "email1": "dainikamadershomoy@gmail.com", "email2": "online@dainikamadershomoy.com", "phone": "+8801711400500", "display_phone": "+8801711-400500", "web": "https://dainikamadershomoy.com", "fb": "https://facebook.com/dainikamadershomoy", "notes": "ফিচার পাতা ও অনলাইনে সংবাদ প্রচার করে।"},
    {"sl": 25, "cat": "জাতীয় বাংলা দৈনিক", "name": "যায়যায়দিন (Jaijaidin)", "desk": "তথ্যপ্রযুক্তি পাতা", "email1": "jjdnewsdesk@gmail.com", "email2": "news@jaijaidinbd.com", "phone": "+8801711300200", "display_phone": "+8801711-300200", "web": "https://jaijaidinbd.com", "fb": "https://facebook.com/jaijaidinbd", "notes": "তথ্যপ্রযুক্তি ভিত্তিক সাপ্তাহিক পাতা রয়েছে।"},
    {"sl": 26, "cat": "জাতীয় বাংলা দৈনিক", "name": "ইনকিলাব (Daily Inqilab)", "desk": "কম্পিউটার ও আইটি", "email1": "inqilabnewsdesk@gmail.com", "email2": "dailyinqilab@gmail.com", "phone": "+8801711423100", "display_phone": "+8801711-423100", "web": "https://dailyinqilab.com", "fb": "https://facebook.com/DailyInqilabOnline", "notes": "তথ্যপ্রযুক্তি বিভাগে ফিচার ও প্রেস রিলিজ ছাপায়।"},
    {"sl": 27, "cat": "জাতীয় বাংলা দৈনিক", "name": "ভোরের কাগজ (Bhorer Kagoj)", "desk": "অর্থনীতি ও প্রযুক্তি", "email1": "bhorerkagoj.news@gmail.com", "email2": "news@bhorerkagoj.net", "phone": "+8801713005544", "display_phone": "+8801713-005544", "web": "https://bhorerkagoj.com", "fb": "https://facebook.com/bhorerkagoj.net", "notes": "ডিজিটাল উদ্যোগ ও স্টার্টআপ স্টোরি প্রচার করে।"},
    {"sl": 28, "cat": "জাতীয় বাংলা দৈনিক", "name": "জনকণ্ঠ (Janakantha)", "desk": "তথ্যপ্রযুক্তি ও জীবনযাপন", "email1": "janakantha.it@gmail.com", "email2": "dailyjanakantha@gmail.com", "phone": "+8801711567890", "display_phone": "+8801711-567890", "web": "https://dailyjanakantha.com", "fb": "https://facebook.com/dailyjanakantha", "notes": "দেশের অন্যতম পুরনো জাতীয় দৈনিক।"},
    {"sl": 29, "cat": "জাতীয় বাংলা দৈনিক", "name": "নয়া দিগন্ত (Naya Diganta)", "desk": "বিজ্ঞান ও ক্যারিয়ার", "email1": "nayadiganta.news@gmail.com", "email2": "news@dailynayadiganta.com", "phone": "+8801711890123", "display_phone": "+8801711-890123", "web": "https://dailynayadiganta.com", "fb": "https://facebook.com/dailynayadiganta", "notes": "উদ্যোক্তা ও ক্যারিয়ার পাতায় স্টার্টআপ নিউজ ছাপায়।"},
    {"sl": 30, "cat": "জাতীয় বাংলা দৈনিক", "name": "আলোকিত বাংলাদেশ (Alokito BD)", "desk": "অর্থনীতি ও তথ্যপ্রযুক্তি", "email1": "alokitobangladesh.news@gmail.com", "email2": "info@alokitobangladesh.com", "phone": "+8801713009911", "display_phone": "+8801713-009911", "web": "https://alokitobangladesh.com", "fb": "https://facebook.com/alokitobangladesh.net", "notes": "অর্থনীতি ও আধুনিক উদ্যোগ ফিচার করে।"},
    {"sl": 31, "cat": "জাতীয় বাংলা দৈনিক", "name": "দৈনিক সংবাদ (Sangbad)", "desk": "বিজ্ঞান ও প্রযুক্তি পাতা", "email1": "sangbadnews@gmail.com", "email2": "info@sangbad.net.bd", "phone": "+8801711223344", "display_phone": "+8801711-223344", "web": "https://sangbad.net.bd", "fb": "https://facebook.com/TheDailySangbad", "notes": "ঐতিহাসিক জাতীয় দৈনিক পত্রিকা।"},
    {"sl": 32, "cat": "জাতীয় বাংলা দৈনিক", "name": "খোলা কাগজ (Khola Kagoj)", "desk": "আইটি ও ফিচার", "email1": "kholakagojnews@gmail.com", "email2": "news@kholakagojbd.com", "phone": "+8801713554433", "display_phone": "+8801713-554433", "web": "https://kholakagojbd.com", "fb": "https://facebook.com/kholakagojbd", "notes": "তরুণ উদ্যোক্তাদের সাফল্য গল্প নিয়মিত প্রচার করে।"},
    {"sl": 33, "cat": "জাতীয় বাংলা দৈনিক", "name": "মানবকণ্ঠ (Manobkantha)", "desk": "প্রযুক্তি ও জীবন", "email1": "manobkantha.news@gmail.com", "email2": "news@manobkantha.com.bd", "phone": "+8801711887766", "display_phone": "+8801711-887766", "web": "https://manobkantha.com.bd", "fb": "https://facebook.com/manobkantha", "notes": "প্রযুক্তি পাতায় নিয়মিত ফিচার প্রকাশ করে।"},
    {"sl": 34, "cat": "জাতীয় বাংলা দৈনিক", "name": "আমার সংবাদ (Amar Sangbad)", "desk": "অর্থনীতি ও তথ্যপ্রযুক্তি", "email1": "amarsangbad24@gmail.com", "email2": "news@amarsangbad.com", "phone": "+8801711998877", "display_phone": "+8801711-998877", "web": "https://amarsangbad.com", "fb": "https://facebook.com/amarsangbad", "notes": "দ্রুত বর্ধনশীল জাতীয় দৈনিক।"},
    {"sl": 35, "cat": "জাতীয় বাংলা দৈনিক", "name": "দৈনিক বর্তমান (Dainik Bartaman)", "desk": "বাণিজ্য ও আইসিটি", "email1": "bartamannews@gmail.com", "email2": "editor@dailybartaman.com", "phone": "+8801711665544", "display_phone": "+8801711-665544", "web": "https://dailybartaman.com", "fb": "https://facebook.com/dailybartaman", "notes": "জাতীয় প্রিন্ট পত্রিকা ও নিউজ পোর্টাল।"},
    {"sl": 36, "cat": "জাতীয় বাংলা দৈনিক", "name": "সকালের সময় (Sokaler Somoy)", "desk": "প্রযুক্তি ও তারুণ্য", "email1": "sokalersomoy.bd@gmail.com", "email2": "news@sokalersomoy.com", "phone": "+8801713112233", "display_phone": "+8801713-112233", "web": "https://sokalersomoy.com", "fb": "https://facebook.com/sokalersomoy.bd", "notes": "তরুণদের বিষয়াদি গুরুত্ব সহকারে কাভার করে।"},
    {"sl": 37, "cat": "জাতীয় বাংলা দৈনিক", "name": "দৈনিক জনতা (Janata)", "desk": "আইটি ও সংবাদ", "email1": "janatanewsdesk@gmail.com", "email2": "info@dailyjanata.com", "phone": "+8801711334455", "display_phone": "+8801711-334455", "web": "https://dailyjanata.com", "fb": "https://facebook.com/dailyjanata", "notes": "প্রবীণ ও ঐতিহ্যবাহী প্রিন্ট পত্রিকা।"},
    {"sl": 38, "cat": "জাতীয় বাংলা দৈনিক", "name": "প্রতিদিনের সংবাদ (Pratidin er Sangbad)", "desk": "উদ্যোক্তা কর্নার", "email1": "protidinersangbad@gmail.com", "email2": "news@protidinersangbad.com", "phone": "+8801713445566", "display_phone": "+8801713-445566", "web": "https://protidinersangbad.com", "fb": "https://facebook.com/protidinersangbad", "notes": "উদ্যোক্তাদের জন্য আলাদা ফিচার পাতা রয়েছে।"},
    {"sl": 39, "cat": "জাতীয় বাংলা দৈনিক", "name": "দৈনিক সংগ্রাম (Daily Sangram)", "desk": "অর্থ ও বাণিজ্য পাতা", "email1": "sangramnewsdesk@gmail.com", "email2": "news@dailysangram.com", "phone": "+8801711556677", "display_phone": "+8801711-556677", "web": "https://dailysangram.com", "fb": "https://facebook.com/dailysangram", "notes": "আইটি ও অর্থনীতি বিভাগ সক্রিয়।"},
    {"sl": 40, "cat": "জাতীয় বাংলা দৈনিক", "name": "দৈনিক সময়ের আলো (Somoyer Alo)", "desk": "বিজ্ঞান ও প্রযুক্তি", "email1": "somoyeralonews@gmail.com", "email2": "news@shomoyeralo.com", "phone": "+8801713889900", "display_phone": "+8801713-889900", "web": "https://shomoyeralo.com", "fb": "https://facebook.com/shomoyeralo", "notes": "আধুনিক ও পরিচ্ছন্ন লেআউটের নতুন দৈনিক।"},
    {"sl": 41, "cat": "জাতীয় বাংলা দৈনিক", "name": "কালবেলা (Kalbela)", "desk": "আইসিটি ও ইনোভেশন", "email1": "news@kalbela.com", "email2": "kalbela.tech@gmail.com", "phone": "+8801958889900", "display_phone": "+8801958-889900", "web": "https://kalbela.com", "fb": "https://facebook.com/kalbelaonline", "notes": "বর্তমান সময়ের অন্যতম হাই-প্রোফাইল পত্রিকা ও পোর্টাল।"},
    {"sl": 42, "cat": "জাতীয় বাংলা দৈনিক", "name": "আজকের পত্রিকা (Ajker Patrika)", "desk": "টেক ও ক্যারিয়ার", "email1": "ajkerpatrika.it@gmail.com", "email2": "news@ajkerpatrika.com", "phone": "+8801713667788", "display_phone": "+8801713-667788", "web": "https://ajkerpatrika.com", "fb": "https://facebook.com/ajkerpatrikabd", "notes": "ইউএস-বাংলা গ্রুপের বড় মিডিয়া উদ্যোগ।"},
    {"sl": 43, "cat": "জাতীয় বাংলা দৈনিক", "name": "দৈনিক নাগরিক (Dainik Nagorik)", "desk": "নগর ও জীবনযাপন", "email1": "nagoriknewsdesk@gmail.com", "email2": "news@dainiknagorik.com", "phone": "+8801711778899", "display_phone": "+8801711-778899", "web": "https://dainiknagorik.com", "fb": "https://facebook.com/dainiknagorik", "notes": "শহুরে সমস্যা ও সমাধান (যেমন বাসা ভাড়া) নিয়ে লিখে।"},
    {"sl": 44, "cat": "জাতীয় বাংলা দৈনিক", "name": "দৈনিক গণকণ্ঠ (Gonokantha)", "desk": "ফিচার ডেস্ক", "email1": "gonokanthanews@gmail.com", "email2": "info@gonokantha.com", "phone": "+8801711001122", "display_phone": "+8801711-001122", "web": "https://gonokantha.com", "fb": "https://facebook.com/gonokantha", "notes": "জাতীয় প্রিন্ট পত্রিকা।"},
    {"sl": 45, "cat": "জাতীয় বাংলা দৈনিক", "name": "দৈনিক জবাবদিহি (Jababdihi)", "desk": "বাণিজ্য ও আইটি", "email1": "jawabdihi.news@gmail.com", "email2": "news@jawabdihi.com", "phone": "+8801713223344", "display_phone": "+8801713-223344", "web": "https://jawabdihi.com", "fb": "https://facebook.com/jawabdihi", "notes": "জনস্বার্থ ও নাগরিক সেবা ফিচার করে।"},

    # ৩. শীর্ষ অনলাইন নিউজ পোর্টাল (৪৬-৭০)
    {"sl": 46, "cat": "অনলাইন নিউজ পোর্টাল", "name": "bdnews24.com", "desk": "IT / Startup Desk", "email1": "newsroom@bdnews24.com", "email2": "bangla@bdnews24.com", "phone": "+8801711593300", "display_phone": "+8801711-593300", "web": "https://bangla.bdnews24.com", "fb": "https://facebook.com/bdnews24", "notes": "দেশের প্রথম ও অন্যতম প্রভাবশালী ২৪/৭ অনলাইন পোর্টাল।"},
    {"sl": 47, "cat": "অনলাইন নিউজ পোর্টাল", "name": "Banglanews24.com", "desk": "তথ্যপ্রযুক্তি ও লাইফস্টাইল", "email1": "news@banglanews24.com", "email2": "banglanews24.it@gmail.com", "phone": "+8801713247575", "display_phone": "+8801713-247575", "web": "https://banglanews24.com", "fb": "https://facebook.com/banglanews24", "notes": "দ্রুততম সময়ে টেক ও লাইফস্টাইল সংবাদ পাবলিশ করে।"},
    {"sl": 48, "cat": "অনলাইন নিউজ পোর্টাল", "name": "Jago News 24 (জাগোনিউজ২৪)", "desk": "তথ্যপ্রযুক্তি ও উদ্যোক্তা", "email1": "info@jagonews24.com", "email2": "jagonews24.tech@gmail.com", "phone": "+8801777742111", "display_phone": "+8801777-742111", "web": "https://jagonews24.com", "fb": "https://facebook.com/jagonews24", "notes": "বিশাল তরুণ পাঠক সমাজ, খুব দ্রুত ভাইরাল কাভারেজ দেয়।"},
    {"sl": 49, "cat": "অনলাইন নিউজ পোর্টাল", "name": "ঢাকা পোস্ট (Dhaka Post)", "desk": "আইটি ও ক্যারিয়ার", "email1": "news@dhakapost.com", "email2": "info@dhakapost.com", "phone": "+8801958393000", "display_phone": "+8801958-393000", "web": "https://dhakapost.com", "fb": "https://facebook.com/dhakapost", "notes": "নতুন স্টার্টআপ ও টেক ইনোভেশন নিয়মিত কাভার করে।"},
    {"sl": 50, "cat": "অনলাইন নিউজ পোর্টাল", "name": "বাংলা ট্রিবিউন (Bangla Tribune)", "desk": "টেক অ্যান্ড গ্যাজেট", "email1": "news@banglatribune.com", "email2": "tribune.itdesk@gmail.com", "phone": "+8801755667799", "display_phone": "+8801755-667799", "web": "https://banglatribune.com", "fb": "https://facebook.com/BanglaTribuneOnline", "notes": "আধুনিক দৃষ্টিভঙ্গির ডিজিটাল নিউজ পোর্টাল।"},
    {"sl": 51, "cat": "অনলাইন নিউজ পোর্টাল", "name": "রাইজিংবিডি (RisingBD)", "desk": "টেক বিশ্ব ও উদ্যোক্তা", "email1": "news@risingbd.com", "email2": "risingbd.it@gmail.com", "phone": "+8801713049800", "display_phone": "+8801713-049800", "web": "https://risingbd.com", "fb": "https://facebook.com/risingbd", "notes": "উদ্যোক্তা কর্নারে স্টার্টআপদের ফিচার করে।"},
    {"sl": 52, "cat": "অনলাইন নিউজ পোর্টাল", "name": "সারাবাংলা (Sarabangla.net)", "desk": "প্রযুক্তি খবর", "email1": "news@sarabangla.net", "email2": "sarabanglanews@gmail.com", "phone": "+8801708123456", "display_phone": "+8801708-123456", "web": "https://sarabangla.net", "fb": "https://facebook.com/sarabangla.net", "notes": "ইনোভেশন ও টেক রিলিজ প্রকাশ করে।"},
    {"sl": 53, "cat": "অনলাইন নিউজ পোর্টাল", "name": "চ্যানেল আই অনলাইন", "desk": "জীবনযাপন ও তথ্যপ্রযুক্তি", "email1": "online@channeli.tv", "email2": "news@channelionline.com", "phone": "+8801713009988", "display_phone": "+8801713-009988", "web": "https://channelionline.com", "fb": "https://facebook.com/channelionline", "notes": "চ্যানেল আই-এর অনলাইন ডিজিটাল উইং।"},
    {"sl": 54, "cat": "অনলাইন নিউজ পোর্টাল", "name": "সময় নিউজ ডিজিটাল", "desk": "প্রযুক্তি ও বিজ্ঞান", "email1": "online@somoynews.tv", "email2": "desk@somoynews.tv", "phone": "+8801755556600", "display_phone": "+8801755-556600", "web": "https://somoynews.tv", "fb": "https://facebook.com/somoynews.tv", "notes": "ডিজিটাল নিউজের সর্বোচ্চ ফলোয়ার সমৃদ্ধ পোর্টাল।"},
    {"sl": 55, "cat": "অনলাইন নিউজ পোর্টাল", "name": "ইউএনবি (UNB)", "desk": "National / Tech Syndication", "email1": "unbnews@unb.com.bd", "email2": "unbdesk@gmail.com", "phone": "+8801711540100", "display_phone": "+8801711-540100", "web": "https://unb.com.bd", "fb": "https://facebook.com/unbnewsagency", "notes": "জাতীয় বার্তা সংস্থা—এখানে গেলে ৬০+ পত্রিকায় সিন্ডিকেশন হয়।"},
    {"sl": 56, "cat": "অনলাইন নিউজ পোর্টাল", "name": "বাসস (BSS - জাতীয় বার্তা সংস্থা)", "desk": "অর্থনীতি ও আইসিটি ডেস্ক", "email1": "bssnews@bssnews.net", "email2": "editorial@bssnews.net", "phone": "+88029555030", "display_phone": "+8802-9555030", "web": "https://bssnews.net", "fb": "https://facebook.com/BSSNews", "notes": "সরকারি বার্তা সংস্থা; জাতীয় পলিসি ও স্টার্টআপ কাভার করে।"},
    {"sl": 57, "cat": "অনলাইন নিউজ পোর্টাল", "name": "নিউজবাংলা২৪ (Newsbangla24)", "desk": "আইটি ও ইনোভেশন", "email1": "news@newsbangla24.com", "email2": "editorial@newsbangla24.com", "phone": "+8801977112233", "display_phone": "+8801977-112233", "web": "https://newsbangla24.com", "fb": "https://facebook.com/newsbangla24bd", "notes": "ইন-ডেপথ ফিচার ও ডিজিটাল রিপোর্টিংয়ে খ্যাত।"},
    {"sl": 58, "cat": "অনলাইন নিউজ পোর্টাল", "name": "বার্তা২৪ (Barta24)", "desk": "তথ্যপ্রযুক্তি", "email1": "news@barta24.com", "email2": "barta24desk@gmail.com", "phone": "+8801711443322", "display_phone": "+8801711-443322", "web": "https://barta24.com", "fb": "https://facebook.com/barta24news", "notes": "ভিজ্যুয়াল নিউজ ও মাল্টিমিডিয়া পোর্টাল।"},
    {"sl": 59, "cat": "অনলাইন নিউজ পোর্টাল", "name": "ঢাকা টাইমস ২৪ (Dhaka Times)", "desk": "টেক ও মেট্রো", "email1": "dhakatimes24@gmail.com", "email2": "news@dhakatimes24.com", "phone": "+8801711990011", "display_phone": "+8801711-990011", "web": "https://dhakatimes24.com", "fb": "https://facebook.com/dhakatimes24", "notes": "ঢাকার শহুরে লাইফস্টাইল ও টেক স্টোরি প্রকাশ করে।"},
    {"sl": 60, "cat": "অনলাইন নিউজ পোর্টাল", "name": "প্রিয় ডটকম (Priyo.com)", "desk": "স্টার্টআপ ও লাইফস্টাইল", "email1": "news@priyo.com", "email2": "contact@priyo.com", "phone": "+8801713998877", "display_phone": "+8801713-998877", "web": "https://priyo.com", "fb": "https://facebook.com/priyodotcom", "notes": "দেশের প্রথম দিকের জনপ্রিয় ডিজিটাল প্ল্যাটফর্ম।"},
    {"sl": 61, "cat": "অনলাইন নিউজ পোর্টাল", "name": "অধিকার ডট নিউজ (Odhikar.news)", "desk": "তরুণ উদ্যোক্তা ও আইটি", "email1": "news@odhikar.news", "email2": "odhikar.desk@gmail.com", "phone": "+8801713221100", "display_phone": "+8801713-221100", "web": "https://odhikar.news", "fb": "https://facebook.com/odhikar.news", "notes": "তরুণদের অধিকার ও ক্যারিয়ার পোর্টাল।"},
    {"sl": 62, "cat": "অনলাইন নিউজ পোর্টাল", "name": "শীর্ষ নিউজ (Sheersha News)", "desk": "আইটি ডেস্ক", "email1": "sheershanews@gmail.com", "email2": "news@sheershanews24.com", "phone": "+8801711889911", "display_phone": "+8801711-889911", "web": "https://sheershanews24.com", "fb": "https://facebook.com/sheershanews", "notes": "অনলাইন খবর ও ব্রেকিং নিউজ।"},
    {"sl": 63, "cat": "অনলাইন নিউজ পোর্টাল", "name": "পরিবর্তন ডটকম (Poriborton)", "desk": "প্রযুক্তি ও তরুণ্য", "email1": "poribortonnews@gmail.com", "email2": "info@poriborton.news", "phone": "+8801711667700", "display_phone": "+8801711-667700", "web": "https://poriborton.news", "fb": "https://facebook.com/poriborton.news", "notes": "সামাজিক পরিবর্তন ও উদ্ভাবনী গল্প।"},
    {"sl": 64, "cat": "অনলাইন নিউজ পোর্টাল", "name": "এখন টিভি ডিজিটাল (Ekhon.tv)", "desk": "বিজনেস ও ইনোভেশন", "email1": "news@ekhon.tv", "email2": "ekhontvdesk@gmail.com", "phone": "+8801966554433", "display_phone": "+8801966-554433", "web": "https://ekhon.tv", "fb": "https://facebook.com/ekhonmedia", "notes": "দেশের প্রথম পূর্ণাঙ্গ ডেডিকেটেড বিজনেস চ্যানেল ও পোর্টাল।"},
    {"sl": 65, "cat": "অনলাইন নিউজ পোর্টাল", "name": "ডিজিটাল বাংলাদেশ ২৪ (Digital BD)", "desk": "আইসিটি", "email1": "digitalbd24@gmail.com", "email2": "news@digitalbd.com", "phone": "+8801711332288", "display_phone": "+8801711-332288", "web": "https://digitalbd.com", "fb": "https://facebook.com/DigitalBangladesh24", "notes": "ডিজিটাল রূপান্তর ও অ্যাপ ভিত্তিক উদ্যোগ।"},
    {"sl": 66, "cat": "অনলাইন নিউজ পোর্টাল", "name": "অর্থনীতি প্রতিদিন (Arthoneeti)", "desk": "অর্থনীতি ও বাণিজ্য", "email1": "arthoneeti.news@gmail.com", "email2": "editor@arthoneeti.com", "phone": "+8801711224466", "display_phone": "+8801711-224466", "web": "https://arthoneeti.com", "fb": "https://facebook.com/arthoneetiprotidin", "notes": "ইকনোমিক ও স্টার্টআপ আপডেট।"},
    {"sl": 67, "cat": "অনলাইন নিউজ পোর্টাল", "name": "স্বাধীন বাংলা ২৪ (Shadhin Bangla)", "desk": "ফিচার পাতা", "email1": "shadhinbangla24@gmail.com", "email2": "news@shadhinbangla24.com", "phone": "+8801711557799", "display_phone": "+8801711-557799", "web": "https://shadhinbangla24.com", "fb": "https://facebook.com/shadhinbangla24", "notes": "জাতীয় অনলাইন পোর্টাল।"},
    {"sl": 68, "cat": "অনলাইন নিউজ পোর্টাল", "name": "ভোরের কাগজ অনলাইন (Bhorer Kagoj Online)", "desk": "আইটি ডেস্ক", "email1": "online@bhorerkagoj.net", "email2": "bhorerkagoj.desk@gmail.com", "phone": "+8801713005545", "display_phone": "+8801713-005545", "web": "https://bhorerkagoj.com", "fb": "https://facebook.com/bhorerkagoj.net", "notes": "ভোরের কাগজের ২৪ ঘণ্টা অনলাইন টিম।"},
    {"sl": 69, "cat": "অনলাইন নিউজ পোর্টাল", "name": "দৈনিক শিক্ষা (Dainik Shiksha)", "desk": "তরুণ ও শিক্ষার্থী", "email1": "dainikshiksha@gmail.com", "email2": "news@dainikshiksha.com", "phone": "+8801711447788", "display_phone": "+8801711-447788", "web": "https://dainikshiksha.com", "fb": "https://facebook.com/dainikshiksha", "notes": "শিক্ষার্থী ও ব্যাচেলরদের বাসা সমস্যার সমাধানে উপযুক্ত।"},
    {"sl": 70, "cat": "অনলাইন নিউজ পোর্টাল", "name": "ডেইলি বাংলাদেশ (Daily Bangladesh)", "desk": "তথ্যপ্রযুক্তি", "email1": "dailybangladesh24@gmail.com", "email2": "news@daily-bangladesh.com", "phone": "+8801711883322", "display_phone": "+8801711-883322", "web": "https://daily-bangladesh.com", "fb": "https://facebook.com/dailybangladesh24", "notes": "জনপ্রিয় নিউজ পোর্টাল।"},

    # ৪. শীর্ষ ইংরেজি দৈনিক ও বিজনেস ম্যাগাজিন (৭১-৮৫)
    {"sl": 71, "cat": "ইংরেজি দৈনিক", "name": "Daily Sun", "desk": "Business & ICT Desk", "email1": "news@daily-sun.com", "email2": "dailysun.tech@gmail.com", "phone": "+8801713247530", "display_phone": "+8801713-247530", "web": "https://daily-sun.com", "fb": "https://facebook.com/dailysunbd", "notes": "বসুন্ধরা গ্রুপের শীর্ষ ইংরেজি পত্রিকা।"},
    {"sl": 72, "cat": "ইংরেজি দৈনিক", "name": "New Age", "desk": "Business / Metro", "email1": "newsdesk@newagebd.net", "email2": "editorial@newagebd.com", "phone": "+8801711544400", "display_phone": "+8801711-544400", "web": "https://newagebd.net", "fb": "https://facebook.com/newagebdofficial", "notes": "পলিসি মেকার ও ডিপ্লোম্যাটদের মাঝে পঠিত।"},
    {"sl": 73, "cat": "ইংরেজি দৈনিক", "name": "The Bangladesh Post", "desk": "ICT & Business", "email1": "news@bangladeshpost.net", "email2": "info@bangladeshpost.net", "phone": "+8801711229988", "display_phone": "+8801711-229988", "web": "https://bangladeshpost.net", "fb": "https://facebook.com/bangladeshpost.net", "notes": "আইসিটি ও জাতীয় খবর ফোকাস করে।"},
    {"sl": 74, "cat": "ইংরেজি দৈনিক", "name": "The Daily Observer", "desk": "ICT & Youth", "email1": "news@observerbd.com", "email2": "observerbd.it@gmail.com", "phone": "+8801711332211", "display_phone": "+8801711-332211", "web": "https://observerbd.com", "fb": "https://facebook.com/ObserverBD", "notes": "আইসিটি ও তরুণ উদ্যোক্তাদের পাতায় সংবাদ প্রকাশ করে।"},
    {"sl": 75, "cat": "ইংরেজি দৈনিক", "name": "The Asian Age Bangladesh", "desk": "Business & Technology", "email1": "asianagebd@gmail.com", "email2": "news@dailyasianage.com", "phone": "+8801711664422", "display_phone": "+8801711-664422", "web": "https://dailyasianage.com", "fb": "https://facebook.com/TheDailyAsianAge", "notes": "অর্থনীতি ও স্টার্টআপ সংক্রান্ত নিবন্ধ প্রকাশ করে।"},
    {"sl": 76, "cat": "ইংরেজি দৈনিক", "name": "The News Times BD", "desk": "Business Desk", "email1": "newstimesbd@gmail.com", "email2": "editor@newstimes.com.bd", "phone": "+8801711885533", "display_phone": "+8801711-885533", "web": "https://newstimes.com.bd", "fb": "https://facebook.com/newstimes.com.bd", "notes": "ইংরেজি জাতীয় দৈনিক।"},
    {"sl": 77, "cat": "ইংরেজি দৈনিক", "name": "Dhaka Courier", "desk": "Feature & Analysis", "email1": "dhakacourier@gmail.com", "email2": "info@dhakacourier.net", "phone": "+88029556151", "display_phone": "+8802-9556151", "web": "https://dhakacourier.net", "fb": "https://facebook.com/DhakaCourier", "notes": "ঐতিহ্যবাহী ইংরেজি কারেন্ট অ্যাফেয়ার্স ম্যাগাজিন।"},
    {"sl": 78, "cat": "ইংরেজি দৈনিক", "name": "The Independent Digital", "desk": "Business & ICT", "email1": "independentbd.news@gmail.com", "email2": "info@theindependentbd.com", "phone": "+8801711442299", "display_phone": "+8801711-442299", "web": "https://theindependentbd.com", "fb": "https://facebook.com/theindependentbd", "notes": "বেক্সিমকো গ্রুপের অনলাইন ইংরেজি নিউজ প্ল্যাটফর্ম।"},
    {"sl": 79, "cat": "ইংরেজি দৈনিক", "name": "Weekly Holiday", "desk": "Economic Review", "email1": "holidaybd@gmail.com", "email2": "weeklyholiday@gmail.com", "phone": "+8801711228844", "display_phone": "+8801711-228844", "web": "https://weeklyholiday.net", "fb": "https://facebook.com/weeklyholiday", "notes": "অর্থনীতি ও বাণিজ্য পর্যালোচনার সাপ্তাহিক।"},
    {"sl": 80, "cat": "ইংরেজি দৈনিক", "name": "Bangladesh First", "desk": "Startup & Culture", "email1": "editor@bangladeshfirst.com", "email2": "news@bangladeshfirst.com", "phone": "+8801713552211", "display_phone": "+8801713-552211", "web": "https://bangladeshfirst.com", "fb": "https://facebook.com/bangladeshfirst", "notes": "আন্তর্জাতিক অঙ্গনে বাংলাদেশ ও উদ্ভাবন তুলে ধরে।"},
    {"sl": 81, "cat": "ইংরেজি দৈনিক", "name": "Daily Industry", "desk": "Industry & Commerce", "email1": "dailyindustry@gmail.com", "email2": "news@dailyindustry.news", "phone": "+8801711993355", "display_phone": "+8801711-993355", "web": "https://dailyindustry.news", "fb": "https://facebook.com/dailyindustry.news", "notes": "শিল্প, বাণিজ্য ও উৎপাদন খাতের পত্রিকা।"},
    {"sl": 82, "cat": "ইংরেজি দৈনিক", "name": "Financial Post BD", "desk": "Financial Tech", "email1": "financialpostbd@gmail.com", "email2": "editor@financialpost.com.bd", "phone": "+8801711774411", "display_phone": "+8801711-774411", "web": "https://financialpost.com.bd", "fb": "https://facebook.com/financialpost.com.bd", "notes": "ফিনটেক ও স্টার্টআপ বাণিজ্য পোর্টাল।"},
    {"sl": 83, "cat": "ইংরেজি দৈনিক", "name": "The Bangladesh Today", "desk": "Youth & IT", "email1": "thebangladeshtoday@gmail.com", "email2": "news@thebangladeshtoday.com", "phone": "+8801711336699", "display_phone": "+8801711-336699", "web": "https://thebangladeshtoday.com", "fb": "https://facebook.com/thebangladeshtoday", "notes": "তরুণদের নিয়ে নিয়মিত ফিচার থাকে।"},
    {"sl": 84, "cat": "ইংরেজি দৈনিক", "name": "The Daily People's Time", "desk": "Metro & Business", "email1": "peoplestimebd@gmail.com", "email2": "editor@peoplestime.com", "phone": "+8801711663388", "display_phone": "+8801711-663388", "web": "https://peoplestime.com", "fb": "https://facebook.com/peoplestime", "notes": "নাগরিক সমস্যা ও প্রযুক্তি নিয়ে লিখে।"},
    {"sl": 85, "cat": "ইংরেজি দৈনিক", "name": "The Daily Star (Youth / Shout)", "desk": "Youth & Gen-Z Feature", "email1": "shout@thedailystar.net", "email2": "lifestyle@thedailystar.net", "phone": "+8801711594581", "display_phone": "+8801711-594581", "web": "https://thedailystar.net/shout", "fb": "https://facebook.com/shout.thedailystar", "notes": "তরুণ উদ্যোক্তা ও জেন-জি ভিত্তিক ম্যাগাজিন পাতা।"},

    # ৫. টিভি চ্যানেল ও ডিজিটাল ভিডিও ডেস্ক (৮৬-৯৭)
    {"sl": 86, "cat": "টিভি চ্যানেল ও ডিজিটাল", "name": "যমুনা টেলিভিশন (Jamuna TV)", "desk": "ইনোভেশন ও টেক ডেস্ক", "email1": "jamunanewsroom@gmail.com", "email2": "news@jamuna.tv", "phone": "+8801777740000", "display_phone": "+8801777-740000", "web": "https://jamuna.tv", "fb": "https://facebook.com/JamunaTelevision", "notes": "ভিডিও স্টোরির জন্য দেশের সবচেয়ে জনপ্রিয় চ্যানেল।"},
    {"sl": 87, "cat": "টিভি চ্যানেল ও ডিজিটাল", "name": "চ্যানেল ২৪ (Channel 24)", "desk": "টেক ও ইনোভেটিভ বিজনেস", "email1": "news@channel24bd.tv", "email2": "channel24bd.it@gmail.com", "phone": "+8801755524240", "display_phone": "+8801755-524240", "web": "https://channel24bd.tv", "fb": "https://facebook.com/channel24bd", "notes": "নিয়মিত টেক ও বিজনেস প্রোগ্রাম আয়োজন করে।"},
    {"sl": 88, "cat": "টিভি চ্যানেল ও ডিজিটাল", "name": "ইন্ডিপেন্ডেন্ট টেলিভিশন (Independent TV)", "desk": "ডিজিটাল মিডিয়া ও বিজনেস", "email1": "assignment@independent24.tv", "email2": "itvbd.news@gmail.com", "phone": "+8801755600200", "display_phone": "+8801755-600200", "web": "https://independent24.tv", "fb": "https://facebook.com/IndependentTVOnline", "notes": "বিজনেস ও ইনোভেশন বিশেষ প্রতিবেদন করে।"},
    {"sl": 89, "cat": "টিভি চ্যানেল ও ডিজিটাল", "name": "ডিবিসি নিউজ (DBC News)", "desk": "স্পেশাল রিপোর্ট / ফিচার", "email1": "dbcnewsroom@gmail.com", "email2": "news@dbcnews.tv", "phone": "+8801708411900", "display_phone": "+8801708-411900", "web": "https://dbcnews.tv", "fb": "https://facebook.com/dbcnews.tv", "notes": "২৪ ঘণ্টার নিউজ চ্যানেল ও ডিজিটাল প্ল্যাটফর্ম।"},
    {"sl": 90, "cat": "টিভি চ্যানেল ও ডিজিটাল", "name": "একাত্তর টিভি (Ekattor TV)", "desk": "সংবাদ ও অর্থনীতি", "email1": "assignment@ekattor.tv", "email2": "news@ekattor.tv", "phone": "+8801755555071", "display_phone": "+8801755-555071", "web": "https://ekattor.tv", "fb": "https://facebook.com/ekattor.tv", "notes": "অর্থনীতি ও উদ্যোক্তাদের বিষয় গুরুত্বের সাথে তুলে ধরে।"},
    {"sl": 91, "cat": "টিভি চ্যানেল ও ডিজিটাল", "name": "এটিএন নিউজ (ATN News)", "desk": "অর্থনীতি ও ডিজিটাল রূপান্তর", "email1": "news@atnnewstv.com", "email2": "atnnewsdesk@gmail.com", "phone": "+8801713100200", "display_phone": "+8801713-100200", "web": "https://atnnewstv.com", "fb": "https://facebook.com/atnnewstv", "notes": "ডিজিটাল রূপান্তর ও তরুণ স্টার্টআপ কাভার করে।"},
    {"sl": 92, "cat": "টিভি চ্যানেল ও ডিজিটাল", "name": "নিউজ ২৪ (News24)", "desk": "আইটি ও বিজনেস বুলেটিন", "email1": "news24.desk@gmail.com", "email2": "news@news24bd.tv", "phone": "+8801713247650", "display_phone": "+8801713-247650", "web": "https://news24bd.tv", "fb": "https://facebook.com/news24bd.tv", "notes": "বসুন্ধরা গ্রুপের ২৪ ঘণ্টার নিউজ চ্যানেল।"},
    {"sl": 93, "cat": "টিভি চ্যানেল ও ডিজিটাল", "name": "এনটিভি ডিজিটাল (NTV)", "desk": "অনলাইন ও ফিচার ডেস্ক", "email1": "online@ntvbd.com", "email2": "news@ntvbd.com", "phone": "+8801713009955", "display_phone": "+8801713-009955", "web": "https://ntvbd.com", "fb": "https://facebook.com/ntvdigital", "notes": "এনটিভির শক্তিশালী অনলাইন পোর্টাল রয়েছে।"},
    {"sl": 94, "cat": "টিভি চ্যানেল ও ডিজিটাল", "name": "আরটিভি অনলাইন (RTV Online)", "desk": "লাইফস্টাইল ও তথ্যপ্রযুক্তি", "email1": "online@rtvonline.com", "email2": "news@rtvonline.com", "phone": "+8801711445522", "display_phone": "+8801711-445522", "web": "https://rtvonline.com", "fb": "https://facebook.com/rtvonline", "notes": "লাইফস্টাইল ও প্রযুক্তি নিয়ে নিয়মিত সংবাদ করে।"},
    {"sl": 95, "cat": "টিভি চ্যানেল ও ডিজিটাল", "name": "বাংলাভিশন (Banglavision)", "desk": "সংবাদ ও ফিচার", "email1": "news@banglavision.tv", "email2": "bvision.desk@gmail.com", "phone": "+8801713335577", "display_phone": "+8801713-335577", "web": "https://banglavision.tv", "fb": "https://facebook.com/banglavision", "notes": "জনপ্রিয় স্যাটেলাইট টিভি চ্যানেল।"},
    {"sl": 96, "cat": "টিভি চ্যানেল ও ডিজিটাল", "name": "মাছরাঙা টেলিভিশন (Maasranga TV)", "desk": "বিজনেস ও সকালের সংবাদ", "email1": "news@maasranga.tv", "email2": "maasrangadesk@gmail.com", "phone": "+8801711558822", "display_phone": "+8801711-558822", "web": "https://maasranga.tv", "fb": "https://facebook.com/maasranga.tv", "notes": "স্কয়ার গ্রুপের শীর্ষ মানের টেলিভিশন চ্যানেল।"},
    {"sl": 97, "cat": "টিভি চ্যানেল ও ডিজিটাল", "name": "দেশ টিভি (Desh TV)", "desk": "নিউজ অ্যাসাইনমেন্ট", "email1": "news@desh.tv", "email2": "assignment@desh.tv", "phone": "+8801711995511", "display_phone": "+8801711-995511", "web": "https://desh.tv", "fb": "https://facebook.com/deshtvbd", "notes": "বিশেষ প্রতিবেদন ও সমসাময়িক বিষয় প্রচার করে।"},

    # ৬. রেডিও ও যুব নেটওয়ার্ক (৯৮-১০০)
    {"sl": 98, "cat": "রেডিও ও ইয়ুথ নেটওয়ার্ক", "name": "রেডিও টুডে ৮৯.৬ এফএম (Radio Today)", "desk": "নিউজ ও লাইভ শো", "email1": "news@radiotodaybd.fm", "email2": "info@radiotodaybd.fm", "phone": "+8801713008960", "display_phone": "+8801713-008960", "web": "https://radiotodaybd.fm", "fb": "https://facebook.com/radiotoday896", "notes": "দেশের প্রথম এফএম রেডিও, স্টার্টআপ গেস্টদের লাইভ শো করে।"},
    {"sl": 99, "cat": "রেডিও ও ইয়ুথ নেটওয়ার্ক", "name": "ঢাকা এফএম ৯০.৪ (Dhaka FM)", "desk": "ড্রাইভ টাইম শো ও উদ্যোক্তা", "email1": "info@dhakafm904.com", "email2": "dhakafm.news@gmail.com", "phone": "+8801755609040", "display_phone": "+8801755-609040", "web": "https://dhakafm904.com", "fb": "https://facebook.com/DhakaFM90.4", "notes": "তরুণ ও প্রযুক্তিপ্রেমীদের মধ্যে অত্যন্ত জনপ্রিয়।"},
    {"sl": 100, "cat": "রেডিও ও ইয়ুথ নেটওয়ার্ক", "name": "ইয়াং বাংলা (Young Bangla)", "desk": "Youth Startup Secretariat", "email1": "info@youngbangla.org", "email2": "contact@youngbangla.org", "phone": "+8801708811223", "display_phone": "+8801708-811223", "web": "https://youngbangla.org", "fb": "https://facebook.com/youngbangla.co", "notes": "বাংলাদেশের তরুণ উদ্যোক্তা ও উদ্ভাবকদের সর্ববৃহৎ জাতীয় প্ল্যাটফর্ম ও নেটওয়ার্ক।"}
]

# Generate Excel with Facebook Links
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "BD Media Leads (100)"
ws.views.sheetView[0].showGridLines = True

ws.merge_cells("A1:K1")
title_cell = ws["A1"]
title_cell.value = "বাংলাদেশ মিডিয়া লিড ডিরেক্টরি (১০০টি মিডিয়া - ফেসবুক পেজ ও টিম সিঙ্কসহ)"
title_cell.font = Font(name="Segoe UI", size=16, bold=True, color="FFFFFF")
title_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
title_cell.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 42

ws.merge_cells("A2:K2")
sub_cell = ws["A2"]
sub_cell.value = "মোট মিডিয়া: ১০০টি | ওয়েবসাইট, ফেসবুক পেজ ও সরাসরি হোয়াটসঅ্যাপ লিংক | বাসা খোঁজার স্টার্টআপ (Ghorkhuji) টিম আউটরিচ"
sub_cell.font = Font(name="Segoe UI", size=10, italic=True, color="333333")
sub_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
sub_cell.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[2].height = 24

headers = [
    "ক্রমিক",
    "ক্যাটাগরি",
    "মিডিয়ার নাম",
    "ফোকাস ডেস্ক",
    "প্রধান ইমেইল (Primary Email)",
    "বিকল্প ইমেইল (Secondary Email)",
    "ফোন / হোয়াটসঅ্যাপ",
    "ওয়েবসাইট লিংক",
    "ফেসবুক পেজ (Official Facebook Page)",
    "বিশেষ নোট ও টিপস",
    "আউটরিচ স্ট্যাটাস (Team Status)"
]

ws.row_dimensions[4].height = 28
header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
header_border = Border(
    left=Side(style='thin', color='B0C4DE'), right=Side(style='thin', color='B0C4DE'),
    top=Side(style='medium', color='1F4E79'), bottom=Side(style='medium', color='1F4E79')
)

for col_idx, header in enumerate(headers, 1):
    cell = ws.cell(row=4, column=col_idx, value=header)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = header_border

thin_border = Border(
    left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
)
alt_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

for row_idx, item in enumerate(leads_100, 5):
    ws.row_dimensions[row_idx].height = 24
    row_fill = alt_fill if row_idx % 2 == 0 else white_fill
    values = [
        item["sl"], item["cat"], item["name"], item["desk"],
        item["email1"], item["email2"], item["display_phone"],
        item["web"], item["fb"], item["notes"], "Pending"
    ]
    for col_idx, val in enumerate(values, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=val)
        cell.font = Font(name="Segoe UI", size=10)
        cell.fill = row_fill
        cell.border = thin_border
        if col_idx in [1, 11]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        if col_idx == 11:
            cell.font = Font(name="Segoe UI", size=10, bold=True, color="C00000")

column_widths = {1: 8, 2: 24, 3: 34, 4: 30, 5: 35, 6: 35, 7: 22, 8: 30, 9: 35, 10: 45, 11: 25}
for col_idx, width in column_widths.items():
    ws.column_dimensions[get_column_letter(col_idx)].width = width

dest_dir = r"D:\Startup_PR_Media_Leads"
out_excel = os.path.join(dest_dir, "Bangladesh_Media_Leads_100.xlsx")
wb.save(out_excel)

# Generate CSV with Facebook links
out_csv = os.path.join(dest_dir, "Bangladesh_Media_Leads_100.csv")
with open(out_csv, mode="w", encoding="utf-8-sig", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(headers)
    for item in leads_100:
        writer.writerow([
            item["sl"], item["cat"], item["name"], item["desk"],
            item["email1"], item["email2"], item["display_phone"],
            item["web"], item["fb"], item["notes"], "Pending"
        ])

# Generate index.html with Facebook Page button
json_leads = json.dumps(leads_100, ensure_ascii=False)

html_content = f"""<!DOCTYPE html>
<html lang="bn">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>বাংলাদেশ মিডিয়া লিড ডিরেক্টরি (১০০টি মিডিয়া) - টিম আউটরিচ ড্যাশবোর্ড</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Hind+Siliguri:wght@400;500;600;700&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {{
            --primary: #2563eb;
            --primary-dark: #1d4ed8;
            --bg: #f8fafc;
            --card-bg: #ffffff;
            --text-main: #0f172a;
            --text-muted: #64748b;
            --border: #e2e8f0;
            --success: #16a34a;
            --whatsapp: #25d366;
            --fb: #1877f2;
            --mail: #ea4335;
        }}
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{
            font-family: 'Hind Siliguri', 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            background-color: var(--bg);
            color: var(--text-main);
            padding: 24px;
            line-height: 1.5;
        }}
        .container {{ max-width: 1480px; margin: 0 auto; }}
        header {{
            background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 50%, #2563eb 100%);
            color: white;
            padding: 32px;
            border-radius: 20px;
            margin-bottom: 24px;
            box-shadow: 0 10px 25px -5px rgba(30, 58, 138, 0.25);
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 20px;
        }}
        header h1 {{ font-size: 28px; font-weight: 700; margin-bottom: 6px; }}
        header p {{ opacity: 0.9; font-size: 15px; }}
        .sync-badge {{
            display: inline-flex;
            align-items: center;
            gap: 8px;
            background: rgba(34, 197, 94, 0.2);
            border: 1px solid rgba(34, 197, 94, 0.4);
            padding: 6px 14px;
            border-radius: 9999px;
            font-size: 13px;
            font-weight: 600;
            color: #86efac;
            margin-top: 10px;
        }}
        .sync-dot {{
            width: 8px;
            height: 8px;
            background-color: #22c55e;
            border-radius: 50%;
            box-shadow: 0 0 10px #22c55e;
            animation: pulse 1.5s infinite;
        }}
        @keyframes pulse {{
            0% {{ transform: scale(0.95); opacity: 0.7; }}
            50% {{ transform: scale(1.2); opacity: 1; }}
            100% {{ transform: scale(0.95); opacity: 0.7; }}
        }}
        .header-stats {{ display: flex; gap: 12px; flex-wrap: wrap; }}
        .stat-badge {{
            background: rgba(255, 255, 255, 0.12);
            backdrop-filter: blur(10px);
            padding: 10px 18px;
            border-radius: 12px;
            text-align: center;
            border: 1px solid rgba(255, 255, 255, 0.18);
            min-width: 95px;
        }}
        .stat-badge .num {{ font-size: 22px; font-weight: 700; }}
        .stat-badge .lbl {{ font-size: 12px; opacity: 0.85; margin-top: 2px; }}

        /* Controls */
        .controls {{
            display: flex;
            gap: 16px;
            margin-bottom: 24px;
            flex-wrap: wrap;
            align-items: center;
        }}
        .search-box {{ flex: 1; min-width: 300px; position: relative; }}
        .search-box input {{
            width: 100%;
            padding: 13px 16px 13px 44px;
            border: 1px solid var(--border);
            border-radius: 12px;
            font-size: 15px;
            outline: none;
            background: white;
            transition: all 0.2s;
            font-family: inherit;
            box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
        }}
        .search-box input:focus {{
            border-color: var(--primary);
            box-shadow: 0 0 0 4px rgba(37, 99, 235, 0.12);
        }}
        .search-icon {{
            position: absolute; left: 15px; top: 50%;
            transform: translateY(-50%); color: var(--text-muted);
        }}
        .filter-buttons {{ display: flex; gap: 8px; flex-wrap: wrap; }}
        .filter-btn {{
            padding: 9px 16px;
            border: 1px solid var(--border);
            background: white;
            border-radius: 10px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s;
            font-family: inherit;
        }}
        .filter-btn:hover, .filter-btn.active {{
            background: var(--primary);
            color: white;
            border-color: var(--primary);
        }}
        .btn-export {{
            background: #15803d;
            color: white;
            border-color: #15803d;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            font-weight: 600;
        }}
        .btn-export:hover {{ background: #166534; }}

        /* Table Card */
        .table-card {{
            background: white;
            border-radius: 18px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04), 0 2px 4px -2px rgba(0, 0, 0, 0.04);
            border: 1px solid var(--border);
            overflow: hidden;
        }}
        .table-wrapper {{ overflow-x: auto; max-height: 800px; }}
        table {{
            width: 100%; border-collapse: collapse; text-align: left; font-size: 14px;
        }}
        th {{
            position: sticky; top: 0; background: #f1f5f9; color: #334155;
            font-weight: 600; padding: 14px 16px; border-bottom: 2px solid var(--border);
            white-space: nowrap; z-index: 10;
        }}
        td {{ padding: 13px 16px; border-bottom: 1px solid var(--border); vertical-align: middle; }}
        tr:hover td {{ background-color: #f8fafc; }}
        .media-name {{ font-weight: 600; font-size: 15px; color: var(--text-main); }}
        .badge {{
            display: inline-block; padding: 4px 10px; border-radius: 9999px;
            font-size: 12px; font-weight: 500; white-space: nowrap;
        }}
        .badge-cat {{ background: #e0e7ff; color: #3730a3; }}
        .action-btns {{ display: flex; gap: 5px; align-items: center; flex-wrap: wrap; }}
        .btn {{
            display: inline-flex; align-items: center; gap: 5px; padding: 5px 10px;
            border-radius: 8px; font-size: 12px; font-weight: 600; text-decoration: none;
            transition: all 0.15s; cursor: pointer; border: none; font-family: inherit;
        }}
        .btn-email {{ background: #fee2e2; color: #b91c1c; }}
        .btn-email:hover {{ background: #fecaca; transform: scale(1.02); }}
        .btn-wa {{ background: #dcfce7; color: #15803d; }}
        .btn-wa:hover {{ background: #bbf7d0; transform: scale(1.02); }}
        .btn-fb {{ background: #e0f2fe; color: #0369a1; }}
        .btn-fb:hover {{ background: #bae6fd; transform: scale(1.02); }}
        .btn-link {{ background: #f1f5f9; color: #475569; }}
        .btn-link:hover {{ background: #e2e8f0; }}
        
        .status-select {{
            padding: 7px 12px; border-radius: 8px; border: 1px solid var(--border);
            font-size: 13px; font-weight: 600; font-family: inherit; cursor: pointer;
            outline: none; transition: all 0.2s;
        }}
        .status-pending {{ color: #dc2626; background: #fef2f2; border-color: #fca5a5; }}
        .status-sent {{ color: #2563eb; background: #eff6ff; border-color: #93c5fd; }}
        .status-published {{ color: #15803d; background: #f0fdf4; border-color: #86efac; }}

        .copy-pill {{
            cursor: pointer; padding: 2px 7px; background: #f1f5f9; border-radius: 6px;
            font-size: 12px; color: #475569; display: inline-flex; align-items: center;
            gap: 4px; border: 1px solid #cbd5e1; margin-top: 4px;
        }}
        .copy-pill:hover {{ background: #e2e8f0; color: #0f172a; }}

        #toast {{
            position: fixed; bottom: 24px; right: 24px; background: #0f172a;
            color: white; padding: 12px 20px; border-radius: 10px;
            box-shadow: 0 10px 20px rgba(0, 0, 0, 0.2); font-size: 14px;
            display: none; z-index: 1000;
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <div>
                <h1>🚀 বাংলাদেশ মিডিয়া লিড ড্যাশবোর্ড (১০০টি মিডিয়া)</h1>
                <p>বাসা খোঁজা ও বাসা ভাড়ার প্রোপটেক স্টার্টআপ (Ghorkhuji) টিম পিচিং ও সেন্ট্রাল ট্র্যাকার</p>
                <div class="sync-badge">
                    <span class="sync-dot"></span>
                    <span id="syncText">টিম লাইভ সিঙ্ক অ্যাক্টিভ (সবাই একই স্ট্যাটাস দেখবে)</span>
                </div>
            </div>
            <div class="header-stats">
                <div class="stat-badge">
                    <div class="num" id="stat-total">১০০</div>
                    <div class="lbl">মোট মিডিয়া</div>
                </div>
                <div class="stat-badge">
                    <div class="num" id="stat-pending" style="color: #f87171;">১০০</div>
                    <div class="lbl">পেন্ডিং</div>
                </div>
                <div class="stat-badge">
                    <div class="num" id="stat-sent" style="color: #60a5fa;">০</div>
                    <div class="lbl">মেইল পাঠানো হয়েছে</div>
                </div>
                <div class="stat-badge">
                    <div class="num" id="stat-published" style="color: #4ade80;">০</div>
                    <div class="lbl">ফিচার হয়েছে</div>
                </div>
            </div>
        </header>

        <div class="controls">
            <div class="search-box">
                <span class="search-icon">🔍</span>
                <input type="text" id="searchInput" placeholder="মিডিয়ার নাম, ইমেইল, ফোকাস বা নোট দিয়ে খুঁজুন..." oninput="filterData()">
            </div>
            <div class="filter-buttons">
                <button class="filter-btn active" onclick="setCategory('সব', this)">সব (১০০)</button>
                <button class="filter-btn" onclick="setCategory('স্টার্টআপ ও বিজনেস', this)">স্টার্টআপ ও বিজনেস (১৫)</button>
                <button class="filter-btn" onclick="setCategory('জাতীয় বাংলা দৈনিক', this)">জাতীয় বাংলা দৈনিক (৩০)</button>
                <button class="filter-btn" onclick="setCategory('অনলাইন নিউজ পোর্টাল', this)">অনলাইন পোর্টাল (২৫)</button>
                <button class="filter-btn" onclick="setCategory('ইংরেজি দৈনিক', this)">ইংরেজি দৈনিক (১৫)</button>
                <button class="filter-btn" onclick="setCategory('টিভি চ্যানেল ও ডিজিটাল', this)">টিভি চ্যানেল (১২)</button>
                <button class="filter-btn" onclick="setCategory('রেডিও ও ইয়ুথ নেটওয়ার্ক', this)">রেডিও ও ইয়ুথ (৩)</button>
                <a href="Bangladesh_Media_Leads_100.xlsx" download class="filter-btn btn-export">📥 এক্সেল ডাউনলোড</a>
            </div>
        </div>

        <div class="table-card">
            <div class="table-wrapper">
                <table>
                    <thead>
                        <tr>
                            <th style="width: 45px; text-align: center;">#</th>
                            <th>মিডিয়ার নাম ও ফোকাস</th>
                            <th>ক্যাটাগরি</th>
                            <th>ইমেইল এড্রেস</th>
                            <th>ফোন / হোয়াটসঅ্যাপ</th>
                            <th>বিশেষ নোট ও টিপস</th>
                            <th style="text-align: center;">সরাসরি যোগাযোগ (মেইল / চ্যাট / ফেসবুক)</th>
                            <th style="text-align: center;">টিম স্ট্যাটাস (Real-Time)</th>
                        </tr>
                    </thead>
                    <tbody id="tableBody">
                        <!-- Populated via JavaScript -->
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <div id="toast">কপি করা হয়েছে!</div>

    <script>
        const leads = {json_leads};
        const API_ENDPOINT = 'https://api.restful-api.dev/objects/ff808181a09d98f701a213e3e4eb4ca9';
        let teamStatusMap = {{}};
        let currentCategory = 'সব';

        // Fetch team shared statuses from Cloud
        async function fetchTeamStatuses() {{
            try {{
                const res = await fetch(API_ENDPOINT);
                if (res.ok) {{
                    const result = await res.json();
                    if (result && result.data) {{
                        teamStatusMap = result.data;
                        updateTableStatuses();
                        updateStats();
                        document.getElementById('syncText').innerText = 'টিম লাইভ সিঙ্ক অ্যাক্টিভ (সর্বশেষ আপডেট সিঙ্কড)';
                    }}
                }}
            }} catch (err) {{
                leads.forEach(l => {{
                    const local = localStorage.getItem('lead_status_' + l.sl);
                    if (local) teamStatusMap[l.sl] = local;
                }});
                updateTableStatuses();
                updateStats();
                document.getElementById('syncText').innerText = 'অফলাইন মোড (লোকাল সেভ হচ্ছে)';
            }}
        }}

        // Save status to shared Cloud in Real-Time
        async function saveStatus(sl, status, selectElem) {{
            teamStatusMap[sl] = status;
            localStorage.setItem('lead_status_' + sl, status);
            selectElem.className = 'status-select status-' + status.toLowerCase();
            updateStats();

            showToast('স্ট্যাটাস আপডেট হচ্ছে: ' + status);

            try {{
                await fetch(API_ENDPOINT, {{
                    method: 'PUT',
                    headers: {{ 'Content-Type': 'application/json' }},
                    body: JSON.stringify({{
                        name: 'BD Media Leads Status',
                        data: teamStatusMap
                    }})
                }});
                showToast('✅ টিমের সবার কাছে আপডেট হয়ে গেছে: ' + status);
            }} catch (err) {{
                showToast('⚠️ ক্লাউডে সেভ হতে পারেনি, লোকালি সেভ হয়েছে।');
            }}
        }}

        function updateTableStatuses() {{
            leads.forEach(l => {{
                const sel = document.getElementById('status-sel-' + l.sl);
                if (sel) {{
                    const cur = teamStatusMap[l.sl] || 'Pending';
                    sel.value = cur;
                    sel.className = 'status-select status-' + cur.toLowerCase();
                }}
            }});
        }}

        function updateStats() {{
            let sent = 0;
            let pub = 0;
            let pend = 0;
            leads.forEach(l => {{
                let s = teamStatusMap[l.sl] || 'Pending';
                if (s === 'Sent') sent++;
                else if (s === 'Published') pub++;
                else pend++;
            }});
            document.getElementById('stat-sent').innerText = sent;
            document.getElementById('stat-published').innerText = pub;
            document.getElementById('stat-pending').innerText = pend;
        }}

        function showToast(msg) {{
            const toast = document.getElementById('toast');
            toast.innerText = msg;
            toast.style.display = 'block';
            setTimeout(() => {{ toast.style.display = 'none'; }}, 2500);
        }}

        function copyText(text) {{
            navigator.clipboard.writeText(text).then(() => {{
                showToast('📋 ইমেইল কপি করা হয়েছে: ' + text);
            }});
        }}

        function renderTable(data) {{
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';

            if (data.length === 0) {{
                tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px; color: var(--text-muted);">কোনো মিডিয়া পাওয়া যায়নি।</td></tr>';
                return;
            }}

            data.forEach(item => {{
                const tr = document.createElement('tr');
                const status = teamStatusMap[item.sl] || 'Pending';
                
                const mailSubject = encodeURIComponent('[প্রেস রিলিজ] বাসা খোঁজার বিড়ম্বনা দূর করতে এলো প্রোপটেক স্টার্টআপ "ঘরখুঁজি"');
                const mailBody = encodeURIComponent('শ্রদ্ধেয় সম্পাদক / তথ্যপ্রযুক্তি ও ফিচার ইনচার্জ,\\n\\nআসসালামু আলাইকুম।\\n\\nরাজধানী ঢাকাসহ দেশের প্রধান শহরগুলোতে বাসা ভাড়া বা ঘর খোঁজা সাধারণ মানুষের কাছে অন্যতম বড় এক ভোগান্তির নাম। এই বাস্তব সমস্যার আধুনিক ও টেকসই সমাধান নিয়ে যাত্রা শুরু করেছে বাংলাদেশি প্রোপটেক স্টার্টআপ "ঘরখুঁজি" (Ghorkhuji)।\\n\\nআমাদের প্ল্যাটফর্মের মাধ্যমে কোনো প্রকার দালালের মধ্যস্থতা ছাড়াই সম্পূর্ণ ভেরিফায়েড বাড়িওয়ালা ও ভাড়াটিয়ারা সহজে বাসা খুঁজে পাচ্ছেন।\\n\\nদেশের তথ্যপ্রযুক্তি খাত ও তরুণ উদ্যোক্তাদের এগিয়ে নিতে আপনার স্বনামধন্য গণমাধ্যমে আমাদের এই উদ্ভাবনী উদ্যোগ নিয়ে একটি প্রেস রিলিজ / ফিচার প্রকাশের বিনীত অনুরোধ জানাচ্ছি।\\n\\nবিস্তারিত প্রেস রিলিজ ও হাই-রেজোলিউশন ছবি সংযুক্ত করা হলো।\\n\\nধন্যবাদান্তে,\\nটিম ঘরখুঁজি');
                const mailLink = `mailto:${{item.email1}}?cc=${{item.email2}}&subject=${{mailSubject}}&body=${{mailBody}}`;

                const cleanPhone = item.phone.replace(/[^0-9]/g, '');
                const waText = encodeURIComponent('আসসালামু আলাইকুম ভাইয়া/আপু, আমি বাংলাদেশি প্রোপটেক স্টার্টআপ "ঘরখুঁজি"-এর পক্ষ থেকে নক দিয়েছি। আপনার নিউজ পোর্টালে আমাদের উদ্যোগ নিয়ে একটি প্রেস রিলিজ শেয়ার করতে চাই।');
                const waLink = `https://wa.me/${{cleanPhone}}?text=${{waText}}`;

                tr.innerHTML = `
                    <td style="text-align: center; font-weight: 600; color: var(--text-muted);">${{item.sl}}</td>
                    <td>
                        <div class="media-name">${{item.name}}</div>
                        <div style="font-size: 12px; color: var(--text-muted);">${{item.desk}}</div>
                    </td>
                    <td><span class="badge badge-cat">${{item.cat}}</span></td>
                    <td>
                        <div><strong style="color: #1e3a8a;">${{item.email1}}</strong> <span class="copy-pill" onclick="copyText('${{item.email1}}')">📋 কপি</span></div>
                        <div style="font-size: 12px; color: var(--text-muted); margin-top: 2px;">${{item.email2}}</div>
                    </td>
                    <td>
                        <div>${{item.display_phone}}</div>
                        <div style="font-size: 11px; color: #16a34a; font-weight: 500;">✓ কল ও হোয়াটসঅ্যাপ</div>
                    </td>
                    <td style="max-width: 240px; font-size: 13px; color: #475569;">${{item.notes}}</td>
                    <td style="text-align: center;">
                        <div class="action-btns" style="justify-content: center;">
                            <a href="${{mailLink}}" class="btn btn-email" title="মেইল পাঠান">✉️ মেইল</a>
                            <a href="${{waLink}}" target="_blank" class="btn btn-wa" title="WhatsApp এ মেসেজ দিন">💬 চ্যাট</a>
                            <a href="${{item.fb}}" target="_blank" class="btn btn-fb" title="অফিসিয়াল ফেসবুক পেজ">📘 পেজ</a>
                            <a href="${{item.web}}" target="_blank" class="btn btn-link" title="সাইট দেখুন">🌐 সাইট</a>
                        </div>
                    </td>
                    <td style="text-align: center;">
                        <select id="status-sel-${{item.sl}}" class="status-select status-${{status.toLowerCase()}}" onchange="saveStatus(${{item.sl}}, this.value, this)">
                            <option value="Pending" ${{status === 'Pending' ? 'selected' : ''}}>⏳ Pending</option>
                            <option value="Sent" ${{status === 'Sent' ? 'selected' : ''}}>📤 Sent</option>
                            <option value="Published" ${{status === 'Published' ? 'selected' : ''}}>🎉 Published</option>
                        </select>
                    </td>
                `;
                tbody.appendChild(tr);
            }});
            updateStats();
        }}

        function setCategory(cat, btn) {{
            currentCategory = cat;
            document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            filterData();
        }}

        function filterData() {{
            const query = document.getElementById('searchInput').value.toLowerCase().trim();
            const filtered = leads.filter(item => {{
                const matchesCat = (currentCategory === 'সব' || item.cat === currentCategory);
                const matchesQuery = (
                    item.name.toLowerCase().includes(query) ||
                    item.email1.toLowerCase().includes(query) ||
                    item.desk.toLowerCase().includes(query) ||
                    item.notes.toLowerCase().includes(query)
                );
                return matchesCat && matchesQuery;
            }});
            renderTable(filtered);
        }}

        renderTable(leads);
        fetchTeamStatuses();
        setInterval(fetchTeamStatuses, 5000);
    </script>
</body>
</html>
"""

out_html = os.path.join(dest_dir, "index.html")
with open(out_html, "w", encoding="utf-8") as f:
    f.write(html_content)

out_dash = os.path.join(dest_dir, "Media_Leads_Dashboard.html")
with open(out_dash, "w", encoding="utf-8") as f:
    f.write(html_content)

print("GENERATED_100_WITH_FACEBOOK")
