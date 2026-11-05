import os
import re
import sys
import asyncio
import datetime
import multiprocessing

print("=== STARTING BOT SCRIPT ===")

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

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
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

def add_matricule(mat: str) -> bool:
    mats = load_matricules()
    clean_new = normalize_matricule(mat)
    existing_clean = [normalize_matricule(m) for m in mats]
    if clean_new not in existing_clean:
        mats.append(mat)
        save_all_matricules(mats)
        return True
    return False

def remove_matricule(mat_to_remove: str) -> bool:
    mats = load_matricules()
    clean_target = normalize_matricule(mat_to_remove)
    new_mats = [m for m in mats if normalize_matricule(m) != clean_target]
    if len(new_mats) < len(mats):
        save_all_matricules(new_mats)
        return True
    return False

MATRICULES, DATE_VAL, DATE_CUSTOM, QTE_VAL, SET_COUNTER_STATE, ADD_MATRICULE_STATE, ADMIN_CUSTOM_DATE = range(7)

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

def save_allowed_user(user_id: int) -> None:
    users = load_allowed_users()
    users.add(user_id)
    with open(ALLOWED_USERS_FILE, "w", encoding="utf-8") as f:
        for uid in users:
            f.write(f"{uid}\n")

def remove_allowed_user(user_id: int) -> None:
    users = load_allowed_users()
    if user_id in users and user_id != ADMIN_ID:
        users.remove(user_id)
        with open(ALLOWED_USERS_FILE, "w", encoding="utf-8") as f:
            for uid in users:
                f.write(f"{uid}\n")

def get_next_number() -> int:
    if not os.path.exists(COUNTER_FILE):
        return DEFAULT_START_NUMBER
    try:
        with open(COUNTER_FILE, "r", encoding="utf-8") as f:
            val = f.read().strip()
            return int(val) if val else DEFAULT_START_NUMBER
    except Exception:
        return DEFAULT_START_NUMBER

def save_next_number(num: int) -> None:
    with open(COUNTER_FILE, "w", encoding="utf-8") as f:
        f.write(str(num))

# توليد ملف PDF بتصميم مطابق 100% للنموذج الأصلي (حسب الصورة المرفقة)
def generate_pdf_worker(matricule, date_full, date_short, qte_choice, current_num, queue):
    try:
        now = datetime.datetime.now()
        year_short = now.strftime("%y")

        num_facture_str = f"RK {year_short}-{current_num}"
        
        qte_float = float(qte_choice)
        pu_ht = 1316.66
        montant_ht = qte_float * pu_ht
        tva = montant_ht * 0.20
        total_ttc = montant_ht + tva

        safe_matricule = "".join([c for c in matricule if c.isalnum() or c in (' ', '_', '-')]).strip()
        if not safe_matricule:
            safe_matricule = f"Facture_{current_num}"

        final_pdf = os.path.join(OUTPUT_DIR, f"{safe_matricule}.pdf")

        c = canvas.Canvas(final_pdf, pagesize=letter)
        width, height = letter

        # 1. رأس الفاتورة (Logo الشركة والعنوان بخط عريض مطابق للصورة)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(40, height - 40, "RACHKAM NEGOCE SARL")
        
        c.setFont("Helvetica", 9)
        c.drawString(40, height - 55, "Transport et Travaux Divers - Argana - Ait Melloul")

        # 2. إطار رقم الفاتورة والتاريخ
        box_y = height - 90
        c.setLineWidth(1)
        c.rect(40, box_y, width - 80, 25)
        
        c.setFont("Helvetica-Bold", 9)
        c.drawString(48, box_y + 8, f"FACTURE PROFORMA : {num_facture_str}")
        c.drawRightString(width - 48, box_y + 8, f"DATE LE {date_full}")

        # 3. إطار اسم العميل (CLIENT : DIVERS)
        client_y = box_y - 35
        c.rect(40, client_y, width - 80, 22)
        c.drawString(48, client_y + 7, "CLIENT : DIVERS")

        # 4. جدول التفاصيل (Table Headers)
        y_table_top = client_y - 30
        c.rect(40, y_table_top - 5, width - 80, 20)
        
        c.setFont("Helvetica-Bold", 8)
        c.drawString(50, y_table_top + 2, "Description")
        c.drawString(210, y_table_top + 2, "QTE/T")
        c.drawString(285, y_table_top + 2, "MATRICULE")
        c.drawString(370, y_table_top + 2, "P, U / HT")
        c.drawString(470, y_table_top + 2, "Montant HT")

        # صف البيانات (Row Data)
        y_row = y_table_top - 25
        c.rect(40, y_row - 5, width - 80, 20)
        c.setFont("Helvetica", 8)
        c.drawString(50, y_row + 2, "CIMENT CPJ 45 SAC")
        c.drawString(210, y_row + 2, f"{qte_float:.2f}")
        c.drawString(285, y_row + 2, matricule)
        c.drawString(370, y_row + 2, f"{pu_ht:,.2f}")
        c.drawString(465, y_row + 2, f"{montant_ht:,.2f}".replace(",", " "))

        # 5. المجاميع (Totals داخل جدول صغير على اليمين تحت الجدول الرئيسي)
        y_tot = y_row - 25
        
        c.rect(340, y_tot - 55, width - 380, 55)
        
        c.setFont("Helvetica-Bold", 8)
        c.drawString(350, y_tot - 12, "TOTAL HT")
        c.drawRightString(width - 48, y_tot - 12, f"{montant_ht:,.2f}".replace(",", " "))

        c.drawString(350, y_tot - 30, "TVA 20%")
        c.drawRightString(width - 48, y_tot - 30, f"{tva:,.2f}".replace(",", " "))

        c.drawString(350, y_tot - 48, "TOTAL TTC")
        c.drawRightString(width - 48, y_tot - 48, f"{total_ttc:,.2f}".replace(",", " "))

        # 6. خانة التوقيع والكاشيه (CACHET ET SIGNATURE)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(40, y_tot - 25, "CACHET ET SIGNATURE")

        # 7. تذييل الصفحة (Footer الرسمي للشركة مطابق للصورة تماماً)
        footer_y = 45
        c.setLineWidth(1.5)
        c.line(40, footer_y + 35, width - 40, footer_y + 35)

        c.setFont("Helvetica-Bold", 7)
        c.drawString(40, footer_y + 22, "SIEGE SOCIAL: N°49 BLOC 4 LOT ARGANA AIT MELLOUL - INEZGANE")
        c.drawString(40, footer_y + 12, "ICE: 001518734000059   IF/76147197   TP/49855558   CNSS 9152638   RC 9149 TRIBUNAL INEZGANE")
        c.drawString(40, footer_y + 2, "TELE/FAX: 05 28 24 06 93   GSM: 06 61 28 26 08   rachkam01@yahoo.fr")

        c.save()

        queue.put(("SUCCESS", final_pdf))
    except Exception as e:
        queue.put(("ERROR", str(e)))

async def admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id != ADMIN_ID:
        await update.message.reply_text("❌ هاد الأمر خاص بالأدمين فقط.")
        return

    current_num = get_next_number()
    total_mats = len(load_matricules())
    keyboard = [
        [InlineKeyboardButton("🚀 إنشاء فواتير سريعة (تحديد الماتريكولات)", callback_data="admin_quick_select")],
        [InlineKeyboardButton("🚛 إدارة أسطول الماتريكولات", callback_data="admin_manage_matricules")],
        [InlineKeyboardButton("👥 قائمة المستخدمين", callback_data="admin_list_users")],
        [InlineKeyboardButton("🚫 إزالة / حظر مستخدم", callback_data="admin_remove_user_prompt")],
        [InlineKeyboardButton("🔢 تعديل رقم العداد الحالي", callback_data="admin_set_counter")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    await update.message.reply_text(
        f"⚙️ **لوحة تحكم الأدمين:**\n\n"
        f"📌 رقم الفاتورة الموالي: `{current_num}`\n"
        f"🚛 عدد الماتريكولات: `{total_mats}`\n\n"
        f"اختر العملية المطلوبة:",
        reply_markup=reply_markup,
        parse_mode="Markdown"
    )

def build_matricule_selection_keyboard(selected_set: set):
    mats = load_matricules()
    keyboard = []
    row = []
    for i, m in enumerate(mats):
        is_selected = m in selected_set
        prefix = "✅ " if is_selected else "⏹ "
        clean_id = normalize_matricule(m)
        row.append(InlineKeyboardButton(f"{prefix}{m}", callback_data=f"selmat_{clean_id}"))
        if len(row) == 2 or i == len(mats) - 1:
            keyboard.append(row)
            row = []
    
    control_row = []
    if len(selected_set) == len(mats):
        control_row.append(InlineKeyboardButton("❌ إلغاء تحديد الكل", callback_data="selmat_none"))
    else:
        control_row.append(InlineKeyboardButton("☑️ تحديد الكل", callback_data="selmat_all"))
    keyboard.append(control_row)

    if selected_set:
        keyboard.append([InlineKeyboardButton(f"➡️ متابعة الإنشاء ({len(selected_set)} محددة)", callback_data="selmat_confirm")])
    
    keyboard.append([InlineKeyboardButton("🔙 العودة للوحة التحكم", callback_data="admin_back")])
    return InlineKeyboardMarkup(keyboard)

async def admin_callback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    if query.from_user.id != ADMIN_ID and not (data.startswith("date_") or data.startswith("qte_")):
        await query.answer("❌ غير مسموح لك بإجراء هذه العملية.", show_alert=True)
        return

    if data == "admin_quick_select":
        context.user_data['admin_selected_mats'] = set()
        await query.edit_message_text(
            "📋 **حدد الماتريكولات التي تريد إنشاء فواتير لها:**\n"
            "اضغط على الماتريكول لتحديده/إلغائه:",
            reply_markup=build_matricule_selection_keyboard(context.user_data['admin_selected_mats']),
            parse_mode="Markdown"
        )

    elif data.startswith("selmat_"):
        action
