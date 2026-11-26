@echo off
chcp 65001 >nul
title ঘরখুঁজি (Ghorkhuji) - WhatsApp PR Automation Assistant
color 0A
cd /d D:\Startup_PR_Media_Leads
echo =====================================================================
echo    ঘরখুঁজি (Ghorkhuji) - WhatsApp মিডিয়া আউটরিচ অটোমেশন
echo =====================================================================
echo.
echo ১. ম্যানুয়াল / স্টেপ-বাই-স্টেপ মোড (প্রতি চ্যাটে Enter প্রেস করে এগোবেন)
echo ২. অটোমেটিক মোড (প্রতি চ্যাটে ১৪ সেকেন্ড বিরতি দিয়ে নিজে নিজেই এগোবে)
echo.
set /p mode="মোড নির্বাচন করুন (1 বা 2 চাপুন): "

if "%mode%"=="2" (
    echo.
    echo অটো-মোড চালু হচ্ছে... (প্রতি লিডে ১৪ সেকেন্ড বিরতি)
    python whatsapp_auto_outreach.py --auto --delay=14
) else (
    echo.
    echo ম্যানুয়াল মোড চালু হচ্ছে...
    python whatsapp_auto_outreach.py
)

pause
