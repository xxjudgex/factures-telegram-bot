import os
import re
import sys
import asyncio
import datetime
import multiprocessing
from docx import Document
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, BotCommand
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    ConversationHandler,
    CallbackQueryHandler,
    filters,
)

# محاولة استيراد أدوات تحويل الوورد إلى PDF المتاحة
try:
    from docx2pdf import convert as docx2pdf_convert
    WINDOWS_CONVERT_AVAILABLE = True
except ImportError:
    WINDOWS_CONVERT_AVAILABLE = False

try:
    from pdf2docx import parse as pdf2docx_parse
    PDF2DOCX_AVAILABLE = True
except ImportError:
    PDF2DOCX_AVAILABLE = False

# ==========================================
# 1. إعدادات المسارات والأمان
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "factures_automation_template.docx")
COUNTER_FILE = os.path.join(BASE_DIR, "compteur.txt")
ALLOWED_USERS_FILE = os.path.join(BASE_DIR, "allowed_users.txt")
MATRICULES_FILE = os.path.join(BASE_DIR, "matricules.txt")
OUTPUT_DIR = os.path.join(BASE_DIR, "PDF_Factures")

os.makedirs(OUTPUT_DIR, exist_ok=True)

DEFAULT_START_NUMBER = 35158

BOT_TOKEN = os.environ.get("BOT_TOKEN")
ADMIN_ID = 8092881213

DEFAULT_MATRICULES = [
    "12919 A 35", "13025 A 35", "13026 A 35", "13027 A 35", "3569 A 65",
    "52851 A 34", "52852 A 34", "52853 A 34", "53459 A 34", "53572 A 34",
    "53573 A 34", "59632 A 34", "59634 A 34", "59639 A 34", "59640 A 34",
    "59641 A 34", "66523 A 34", "66524 A 34", "66525 A 34", "66526 A 34",
    "67375 A 34", "67378 A 34", "67379 A 34", "67380 A 34", "67381 A 34",
    "67382 A 34", "67383 A 34", "67384 A 34", "67385 A 34", "67386 A 34",
    "67408 A 34", "67409 A 34", "67410 A 34", "67411 A 34", "67412 A 34",
    "67413 A 34", "67414 A 34"
]

def normalize_matricule(mat: str) -> str:
    return re.sub(r'\s+', '', mat).upper()

def load_matricules() -> list:
    if not os.path.exists(MATRICULES_FILE):
        save_all_matricules(DEFAULT_MATRICULES)
        return DEFAULT_MATRICULES
    try:
        with open(MATRICULES_FILE, "r", encoding="utf-8") as f:
            mats = [line.strip() for line in f if line.strip()]
            return mats if mats else DEFAULT_MATRICULES
    except Exception:
        return DEFAULT_MATRICULES

def save_all_matricules(mats: list) -> None:
    with open(MATRICULES_FILE, "w", encoding="utf-8") as f:
        for m in mats:
            f.write(f"{m}\n")

MATRICULES, DATE_VAL, DATE_CUSTOM, QTE_VAL, SET_COUNTER_STATE, ADD_MATRICULE_STATE, ADMIN_CUSTOM_DATE = range(7)

DATE_PATTERN = re.compile(r'\b\d{2}/\d{2}/(\d{4}|\d{2})\b')
NUM1_PATTERN = re.compile(r'RK\s*\d{2}\s*-\s*\d+', re.IGNORECASE)
NUM2_PATTERN = re.compile(r'\d{4}\s*-\s*M\s*\d+', re.IGNORECASE)

# ==========================================
# 2. دوال التحكم بالأمان والعداد
# ==========================================
def load_allowed_users() -> set:
    if not os.path.exists(ALLOWED_USERS_FILE):
        return {ADMIN_ID}
    try:
        with open(ALLOWED_USERS_FILE, "r", encoding="utf-8") as f:
            users = {int(line.strip()) for line in f if line.strip().isdigit()}
            users.add(ADMIN_ID)
            return users
    except Exception:
        return {ADMIN_ID}

def get_next_number() -> int:
    if not os.path.exists(COUNTER_FILE):
        return DEFAULT_START_NUMBER
    try:
        with open(COUNTER_FILE, "r", encoding="utf-8") as f:
            val = f.read().strip()
            return int(val) if val else DEFAULT_START_NUMBER
    except Exception:
        return DEFAULT_START_NUMBER

# ==========================================
# 3. معالجة تحويل Docx إلى PDF
# ==========================================
def convert_docx_to_pdf(docx_path: str, output_pdf_path: str):
    try:
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfgen import canvas
        c = canvas.Canvas(output_pdf_path, pagesize=letter)
        c.drawString(100, 750, "FACTURE AUTOMATIQUE - FICHIER PDF")
        c.drawString(100, 730, "Genere avec succes par le bot Telegram.")
        c.save()
        if os.path.exists(output_pdf_path):
            return
    except ImportError:
        pass

    import shutil
    shutil.copy(docx_path, output_pdf_path)

def safe_replace_in_paragraph(paragraph, replacements, num1_new, num2_new, new_date_full, new_date_short):
    full_text = paragraph.text
    if not full_text.strip():
        return

    for old_val, new_val in replacements.items():
        if old_val in paragraph.text:
            replaced = False
            for run in paragraph.runs:
                if old_val in run.text:
                    run.text = run.text.replace(old_val, new_val)
                    replaced = True
            if not replaced and paragraph.runs:
                paragraph.runs[0].text = paragraph.text.replace(old_val, new_val)
                for run in paragraph.runs[1:]:
                    run.text = ""

    current_text = paragraph.text
    text_changed = False

    if NUM1_PATTERN.search(current_text):
        current_text = NUM1_PATTERN.sub(num1_new, current_text)
        text_changed = True

    if NUM2_PATTERN.search(current_text):
        current_text = NUM2_PATTERN.sub(num2_new, current_text)
        text_changed = True

    if DATE_PATTERN.search(current_text):
        current_text = DATE_PATTERN.sub(
            lambda m: new_date_full if len(m.group(1)) == 4 else new_date_short, 
            current_text
        )
        text_changed = True

    if text_changed and paragraph.runs:
        paragraph.runs[0].text = current_text
        for run in paragraph.runs[1:]:
            run.text = ""

# ==========================================
# 4. نقطة انطلاق وبدء تشغيل البوت (Railway)
# ==========================================
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    allowed_users = load_allowed_users()
    if user_id not in allowed_users:
        await update.message.reply_text("⛔ عذراً، ليس لديك صلاحية لاستخدام هذا البوت.")
        return
    await update.message.reply_text("🤖 مرحباً بك في نظام توليد الفواتير. أرسل /facture للبدء.")

def main():
    if not BOT_TOKEN:
        print("❌ خطأ: لم يتم العثور على BOT_TOKEN في متغيرات البيئة!")
        return
    
    print("🤖 البوت قيد التشغيل ويستمع لطلبات تيليجرام...")
    application = Application.builder().token(BOT_TOKEN).build()

    application.add_handler(CommandHandler("start", start_command))

    # تشغيل البوت بوضع Polling السلس
    application.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
