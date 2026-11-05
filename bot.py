import os
import re
import sys
import asyncio
import subprocess
import datetime
import multiprocessing

print("=== STARTING BOT SCRIPT ===")

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

DATE_PATTERN = re.compile(r'\b\d{2}/\d{2}/(\d{4}|\d{2})\b')
NUM1_PATTERN = re.compile(r'RK\s*\d{2}\s*-\s*\d+', re.IGNORECASE)
NUM2_PATTERN = re.compile(r'\d{4}\s*-\s*M\s*\d+', re.IGNORECASE)

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

def convert_docx_to_pdf(docx_path: str, output_pdf_path: str):
    if sys.platform.startswith("win"):
        try:
            from docx2pdf import convert
            convert(docx_path, output_pdf_path)
            return
        except Exception:
            pass

    try:
        cmd = [
            "libreoffice", "--headless", "--convert-to", "pdf",
            docx_path, "--outdir", os.path.dirname(output_pdf_path)
        ]
        subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        generated_pdf = os.path.splitext(docx_path)[0] + ".pdf"
        if os.path.exists(generated_pdf) and generated_pdf != output_pdf_path:
            if os.path.exists(output_pdf_path):
                os.remove(output_pdf_path)
            os.rename(generated_pdf, output_pdf_path)
            return
    except Exception:
        pass

    # إذا لم يتوفر LibreOffice، نقوم بتحويل آمن أو توليد تقرير نظيف يعتمد على القالب
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

def generate_pdf_worker(matricule, date_full, date_short, qte_choice, current_num, queue):
    try:
        if not os.path.exists(TEMPLATE_PATH):
            queue.put(("ERROR", f"ملف القالب غير موجود: {TEMPLATE_PATH}"))
            return

        now = datetime.datetime.now()
        year_full = now.strftime("%Y")
        year_short = now.strftime("%y")

        num_facture_1 = f"RK {year_short}-{current_num}"
        num_facture_2 = f"{year_full}-M{current_num}"

        doc = Document(TEMPLATE_PATH)

        replacements = {
            "{MATRICULE}": matricule,
            "{DATE}": date_full,
            "{NUM1}": num_facture_1,
            "{NUM2}": num_facture_2,
            "{QTE}": f"{qte_choice}.00",
            "53459 A 34": matricule,
            "53459 a 34": matricule,
            "53459 A34": matricule,
            "53459 a34": matricule,
            "53459A34": matricule,
            "53459a34": matricule,
        }

        if qte_choice == "26":
            replacements["27.00"] = "26.00"
            replacements["27,00"] = "26,00"
            replacements[" 27 "] = " 26 "
        else:
            replacements["26.00"] = "27.00"
            replacements["26,00"] = "27,00"
            replacements[" 26 "] = " 27 "

        containers = [doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    containers.append(cell.paragraphs)

        for section in doc.sections:
            for header_footer in [section.header, section.first_page_header, section.footer, section.first_page_footer]:
                if header_footer:
                    containers.append(header_footer.paragraphs)

        for paragraph_list in containers:
            for p in paragraph_list:
                safe_replace_in_paragraph(p, replacements, num_facture_1, num_facture_2, date_full, date_short)

        safe_matricule = "".join([c for c in matricule if c.isalnum() or c in (' ', '_', '-')]).strip()
        if not safe_matricule:
            safe_matricule = f"Facture_{current_num}"

        temp_docx = os.path.join(OUTPUT_DIR, f"temp_{safe_matricule}.docx")
        final_pdf = os.path.join(OUTPUT_DIR, f"{safe_matricule}.pdf")

        doc.save(temp_docx)

        try:
            convert_docx_to_pdf(temp_docx, final_pdf)
        finally:
            if os.path.exists(temp_docx):
                os.remove(temp_docx)

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
        action = data.replace("selmat_", "")
        selected_set = context.user_data.get('admin_selected_mats', set())
        mats = load_matricules()

        if action == "all":
            selected_set = set(mats)
        elif action == "none":
            selected_set = set()
        elif action == "confirm":
            if not selected_set:
                await query.answer("⚠️ المرجو اختيار ماتريكول واحد على الأقل!", show_alert=True)
                return
            context.user_data['matricules'] = list(selected_set)
            
            now = datetime.datetime.now()
            today_str = now.strftime("%d/%m/%Y")
            keyboard = [
                [InlineKeyboardButton(f"📅 تاريخ اليوم ({today_str})", callback_data="date_today")],
                [InlineKeyboardButton("✏️ إدخال تاريخ آخر", callback_data="admin_date_custom")],
                [InlineKeyboardButton("❌ إلغاء", callback_data="cancel_action")]
            ]
            await query.edit_message_text(
                f"✅ تم تحديد {len(selected_set)} ماتريكول.\n\nاختر التاريخ المطلوب:",
                reply_markup=InlineKeyboardMarkup(keyboard)
            )
            return
        else:
            for m in mats:
                if normalize_matricule(m) == action:
                    if m in selected_set:
                        selected_set.remove(m)
                    else:
                        selected_set.add(m)
                    break

        context.user_data['admin_selected_mats'] = selected_set
        await query.edit_message_text(
            f"📋 **حدد الماتريكولات التي تريد إنشاء فواتير لها ({len(selected_set)} محددة):**",
            reply_markup=build_matricule_selection_keyboard(selected_set),
            parse_mode="Markdown"
        )

    elif data.startswith("date_"):
        await date_button_handler(update, context)

    elif data.startswith("qte_"):
        await qte_button_handler(update, context)

    elif data == "admin_manage_matricules":
        mats = load_matricules()
        keyboard = [
            [InlineKeyboardButton("➕ إضافة ماتريكول جديد", callback_data="admin_add_mat_prompt")],
            [InlineKeyboardButton("❌ حذف ماتريكول", callback_data="admin_remove_mat_prompt")],
            [InlineKeyboardButton("📋 عرض كل الماتريكولات", callback_data="admin_view_all_mats")],
            [InlineKeyboardButton("🔙 العودة للوحة التحكم", callback_data="admin_back")]
        ]
        await query.edit_message_text(
            f"🚛 **إدارة أسطول الماتريكولات:**\n\n"
            f"إجمالي الماتريكولات المقبولة حالياً: `{len(mats)}`",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="Markdown"
        )

    elif data == "admin_view_all_mats":
        mats = load_matricules()
        mats_list_str = "\n".join([f"• `{m}`" for m in mats])
        keyboard = [[InlineKeyboardButton("🔙 العودة لـ إدارة الماتريكولات", callback_data="admin_manage_matricules")]]
        await query.edit_message_text(
            f"📋 **قائمة أسطول الماتريكولات المسجلة ({len(mats)}):**\n\n"
            f"{mats_list_str}",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="Markdown"
        )

    elif data == "admin_remove_mat_prompt":
        mats = load_matricules()
        keyboard = []
        row = []
        for i, m in enumerate(mats):
            clean_id = normalize_matricule(m)
            row.append(InlineKeyboardButton(f"❌ {m}", callback_data=f"delmat_{clean_id}"))
            if len(row) == 2 or i == len(mats) - 1:
                keyboard.append(row)
                row = []
        keyboard.append([InlineKeyboardButton("🔙 العودة", callback_data="admin_manage_matricules")])
        await query.edit_message_text(
            "اضغط على الماتريكول الذي تريد حذفه نهائياً من الأسطول:",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="Markdown"
        )

    elif data.startswith("delmat_"):
        target_clean = data.replace("delmat_", "")
        mats = load_matricules()
        removed_name = ""
        for m in mats:
            if normalize_matricule(m) == target_clean:
                removed_name = m
                break
        
        if removed_name and remove_matricule(removed_name):
            msg = f"✅ تم حذف الماتريكول `{removed_name}` بنجاح!"
        else:
            msg = "❌ تعذر حذف الماتريكول."

        keyboard = [[InlineKeyboardButton("🔙 العودة لـ إدارة الماتريكولات", callback_data="admin_manage_matricules")]]
        await query.edit_message_text(msg, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown")

    elif data == "admin_list_users":
        users = load_allowed_users()
        msg = "📋 **قائمة المستخدمين المسموح لهم:**\n\n"
        for uid in users:
            role = "👑 Admin" if uid == ADMIN_ID else "👤 User"
            msg += f"• `{uid}` - {role}\n"
        
        keyboard = [[InlineKeyboardButton("🔙 العودة للوحة التحكم", callback_data="admin_back")]]
        await query.edit_message_text(msg, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data == "admin_remove_user_prompt":
        users = load_allowed_users()
        keyboard = []
        for uid in users:
            if uid != ADMIN_ID:
                keyboard.append([InlineKeyboardButton(f"❌ إزالة `{uid}`", callback_data=f"remove_uid_{uid}")])
        
        keyboard.append([InlineKeyboardButton("🔙 العودة", callback_data="admin_back")])
        await query.edit_message_text(
            "اختر المستخدم الذي تريد إزالته وحظره من البوت:",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="Markdown"
        )

    elif data.startswith("remove_uid_"):
        target_uid = int(data.split("_")[2])
        remove_allowed_user(target_uid)
        keyboard = [[InlineKeyboardButton("🔙 العودة للوحة التحكم", callback_data="admin_back")]]
        await query.edit_message_text(
            f"✅ تم مسح وحظر المستخدم (`{target_uid}`) بنجاح!", 
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="Markdown"
        )

    elif data == "admin_back":
        current_num = get_next_number()
        total_mats = len(load_matricules())
        keyboard = [
            [InlineKeyboardButton("🚀 إنشاء فواتير سريعة (تحديد الماتريكولات)", callback_data="admin_quick_select")],
            [InlineKeyboardButton("🚛 إدارة أسطول الماتريكولات", callback_data="admin_manage_matricules")],
            [InlineKeyboardButton("👥 قائمة المستخدمين", callback_data="admin_list_users")],
            [InlineKeyboardButton("🚫 إزالة / حظر مستخدم", callback_data="admin_remove_user_prompt")],
            [InlineKeyboardButton("🔢 تعديل رقم العداد الحالي", callback_data="admin_set_counter")]
        ]
        await query.edit_message_text(
            f"⚙️ **لوحة تحكم الأدمين:**\n\n📌 رقم الفاتورة الموالي: `{current_num}`\n🚛 عدد الماتريكولات: `{total_mats}`",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="Markdown"
        )

    elif data.startswith("approve_") or data.startswith("reject_"):
        action, user_id_str = data.split("_")
        target_user_id = int(user_id_str)

        if action == "approve":
            save_allowed_user(target_user_id)
            await query.edit_message_text(f"✅ تم قبول المستخدم (`{target_user_id}`)", parse_mode="Markdown")
            try:
                await context.bot.send_message(
                    chat_id=target_user_id,
                    text="🎉 تم قبول طلبك! يمكنك الآن استخدام البوت بإرسال /start"
                )
            except Exception:
                pass

        elif action == "reject":
            await query.edit_message_text(f"❌ تم رفض المستخدم (`{target_user_id}`)", parse_mode="Markdown")
            try:
                await context.bot.send_message(
                    chat_id=target_user_id,
                    text="❌ للأسف، تم رفض طلب انضمامك لاستخدام البوت."
                )
            except Exception:
                pass

async def prompt_admin_custom_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    keyboard = [[InlineKeyboardButton("❌ إلغاء", callback_data="cancel_action")]]
    await query.edit_message_text(
        "أرسل التاريخ بالشكل (JJMM) مثلاً `1509`:", 
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
    return ADMIN_CUSTOM_DATE

async def save_admin_custom_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    date_input = update.message.text.strip()
    now = datetime.datetime.now()
    year_full = now.strftime("%Y")
    year_short = now.strftime("%y")

    digits = "".join([c for c in date_input if c.isdigit()])[:4]
    if len(digits) == 4:
        day = f"{min(max(int(digits[:2]), 1), 31):02d}"
        month = f"{min(max(int(digits[2:]), 1), 12):02d}"
        context.user_data['date_full'] = f"{day}/{month}/{year_full}"
        context.user_data['date_short'] = f"{day}/{month}/{year_short}"
    else:
        keyboard = [[InlineKeyboardButton("❌ إلغاء", callback_data="cancel_action")]]
        await update.message.reply_text(
            "❌ التاريخ غير صحيح. اكتب 4 أرقام (مثلاً 1509):",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
        return ADMIN_CUSTOM_DATE

    keyboard = [
        [InlineKeyboardButton("26.00", callback_data="qte_26"), InlineKeyboardButton("27.00", callback_data="qte_27")],
        [InlineKeyboardButton("❌ إلغاء", callback_data="cancel_action")]
    ]
    await update.message.reply_text(
        f"✅ تم تسجيل التاريخ: `{context.user_data['date_full']}`\n\nاختر الكمية (QTE):",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="Markdown"
    )
    return QTE_VAL

async def prompt_add_matricule(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    keyboard = [[InlineKeyboardButton("❌ إلغاء", callback_data="cancel_action")]]
    await query.edit_message_text(
        "أرسل الماتريكول الجديد بنفس الصيغة التي تريدها أن تظهر في الفاتورة:\n\nمثال:\n`99999 A 34`", 
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
    return ADD_MATRICULE_STATE

async def save_new_matricule_value(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    if text:
        added = add_matricule(text)
        if added:
            await update.message.reply_text(f"✅ تم إضافة الماتريكول الجديد بنجاح: `{text}`\n\nافتح /admin لإدارة الأسطول.", parse_mode="Markdown")
        else:
            await update.message.reply_text(f"⚠️ الماتريكول `{text}` موجود بالفعل في القائمة!", parse_mode="Markdown")
    return ConversationHandler.END

async def prompt_set_counter(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    keyboard = [[InlineKeyboardButton("❌ إلغاء", callback_data="cancel_action")]]
    await query.edit_message_text(
        "أرسل الرقم الجديد الذي تريد أن يبدأ منه العداد (مثلاً: `35158`):", 
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
    return SET_COUNTER_STATE

async def set_counter_value(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    if text.isdigit():
        new_val = int(text)
        save_next_number(new_val)
        await update.message.reply_text(f"✅ تم تحديث العداد بنجاح! الرقم القادم هو: `{new_val}`", parse_mode="Markdown")
    else:
        await update.message.reply_text("❌ المرجو إدخال رقم صحيح فقط.")
    return ConversationHandler.END

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    allowed_users = load_allowed_users()

    context.user_data.clear()

    if user.id not in allowed_users:
        await update.message.reply_text(
            "🔒 هذا البوت خاص واستخدامه يتطلب إذن المالك.\n"
            "تم إرسال طلب الانضمام إلى الأدمين للمراجعة..."
        )
        
        keyboard = [
            [
                InlineKeyboardButton("قبول ✅", callback_data=f"approve_{user.id}"),
                InlineKeyboardButton("رفض ❌", callback_data=f"reject_{user.id}"),
            ]
        ]
        await context.bot.send_message(
            chat_id=ADMIN_ID,
            text=f"🔔 **طلب دخول جديد للبوت:**\n\n"
                 f"👤 الاسم: {user.full_name}\n"
                 f"🏷️ اليوزر: @{user.username if user.username else 'بدون يوزر'}\n"
                 f"🆔 الـ ID: `{user.id}`",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
        return ConversationHandler.END

    await update.message.reply_text(
        "مرحباً بك فـ بوت إنشاء الفواتير! 📄\n\n"
        "أرسل قائمة الماتريكولات (كل ماتريكول فسطر بوحده):\n\n"
        "مثال:\n"
