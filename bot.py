async def admin_callback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    if query.from_user.id != ADMIN_ID and not (data.startswith("date_") or data.startswith("qte_")):
        await query.answer("غير مسموح لك بإجراء هذه العملية.", show_alert=True)
        return

    if data == "admin_quick_select":
        context.user_data['admin_selected_mats'] = set()
        await query.edit_message_text(
            "حدد الماتريكولات التي تريد إنشاء فواتير لها:\n"
            "اضغط على الماتريكول لتحديده/إلغائه:",
            reply_markup=build_matricule_selection_keyboard(context.user_data['admin_selected_mats'])
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
                await query.answer("المرجو اختيار ماتريكول واحد على الأقل!", show_alert=True)
                return
            context.user_data['matricules'] = list(selected_set)
            
            now = datetime.datetime.now()
            today_str = now.strftime("%d/%m/%Y")
            keyboard = [
                [InlineKeyboardButton(f"تاريخ اليوم ({today_str})", callback_data="date_today")],
                [InlineKeyboardButton("إدخال تاريخ آخر", callback_data="admin_date_custom")],
                [InlineKeyboardButton("إلغاء", callback_data="cancel_action")]
            ]
            await query.edit_message_text(
                f"تم تحديد {len(selected_set)} ماتريكول.\n\nاختر التاريخ المطلوب:",
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
            f"حدد الماتريكولات التي تريد إنشاء فواتير لها ({len(selected_set)} محددة):",
            reply_markup=build_matricule_selection_keyboard(selected_set)
        )

    elif data.startswith("date_"):
        await date_button_handler(update, context)

    elif data.startswith("qte_"):
        await qte_button_handler(update, context)

    elif data == "admin_manage_matricules":
        mats = load_matricules()
        keyboard = [
            [InlineKeyboardButton("إضافة ماتريكول جديد", callback_data="admin_add_mat_prompt")],
            [InlineKeyboardButton("حذف ماتريكول", callback_data="admin_remove_mat_prompt")],
            [InlineKeyboardButton("عرض كل الماتريكولات", callback_data="admin_view_all_mats")],
            [InlineKeyboardButton("العودة للوحة التحكم", callback_data="admin_back")]
        ]
        await query.edit_message_text(
            f"إدارة أسطول الماتريكولات:\n\n"
            f"إجمالي الماتريكولات المقبولة حالياً: {len(mats)}",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif data == "admin_view_all_mats":
        mats = load_matricules()
        mats_list_str = "\n".join([f"• {m}" for m in mats])
        keyboard = [[InlineKeyboardButton("العودة لـ إدارة الماتريكولات", callback_data="admin_manage_matricules")]]
        await query.edit_message_text(
            f"قائمة أسطول الماتريكولات المسجلة ({len(mats)}):\n\n"
            f"{mats_list_str}",
            reply_markup=InlineKeyboardMarkup(keyboard)
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
        keyboard.append([InlineKeyboardButton("العودة", callback_data="admin_manage_matricules")])
        await query.edit_message_text(
            "اضغط على الماتريكول الذي تريد حذفه نهائياً من الأسطول:",
            reply_markup=InlineKeyboardMarkup(keyboard)
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
            msg = f"تم حذف الماتريكول {removed_name} بنجاح!"
        else:
            msg = "تعذر حذف الماتريكول."

        keyboard = [[InlineKeyboardButton("العودة لـ إدارة الماتريكولات", callback_data="admin_manage_matricules")]]
        await query.edit_message_text(msg, reply_markup=InlineKeyboardMarkup(keyboard))

    elif data == "admin_list_users":
        users = load_allowed_users()
        msg = "قائمة المستخدمين المسموح لهم:\n\n"
        for uid in users:
            role = "Admin" if uid == ADMIN_ID else "User"
            msg += f"• {uid} - {role}\n"
        
        keyboard = [[InlineKeyboardButton("العودة للوحة التحكم", callback_data="admin_back")]]
        await query.edit_message_text(msg, reply_markup=InlineKeyboardMarkup(keyboard))

    elif data == "admin_remove_user_prompt":
        users = load_allowed_users()
        keyboard = []
        for uid in users:
            if uid != ADMIN_ID:
                keyboard.append([InlineKeyboardButton(f"إزالة {uid}", callback_data=f"remove_uid_{uid}")])
        
        keyboard.append([InlineKeyboardButton("العودة", callback_data="admin_back")])
        await query.edit_message_text(
            "اختر المستخدم الذي تريد إزالته وحظره من البوت:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif data.startswith("remove_uid_"):
        target_uid = int(data.split("_")[2])
        remove_allowed_user(target_uid)
        keyboard = [[InlineKeyboardButton("العودة للوحة التحكم", callback_data="admin_back")]]
        await query.edit_message_text(
            f"تم مسح وحظر المستخدم ({target_uid}) بنجاح!", 
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif data == "admin_back":
        current_num = get_next_number()
        total_mats = len(load_matricules())
        keyboard = [
            [InlineKeyboardButton("إنشاء فواتير سريعة", callback_data="admin_quick_select")],
            [InlineKeyboardButton("إدارة أسطول الماتريكولات", callback_data="admin_manage_matricules")],
            [InlineKeyboardButton("قائمة المستخدمين", callback_data="admin_list_users")],
            [InlineKeyboardButton("إزالة / حظر مستخدم", callback_data="admin_remove_user_prompt")],
            [InlineKeyboardButton("تعديل رقم العداد الحالي", callback_data="admin_set_counter")]
        ]
        await query.edit_message_text(
            f"لوحة تحكم الأدمين:\n\nرقم الفاتورة الموالي: {current_num}\nعدد الماتريكولات: {total_mats}",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif data.startswith("approve_") or data.startswith("reject_"):
        action, user_id_str = data.split("_")
        target_user_id = int(user_id_str)

        if action == "approve":
            save_allowed_user(target_user_id)
            await query.edit_message_text(f"تم قبول المستخدم ({target_user_id})")
            try:
                await context.bot.send_message(
                    chat_id=target_user_id,
                    text="تم قبول طلبك! يمكنك الآن استخدام البوت بإرسال /start"
                )
            except Exception:
                pass

        elif action == "reject":
            await query.edit_message_text(f"تم رفض المستخدم ({target_user_id})")
            try:
                await context.bot.send_message(
                    chat_id=target_user_id,
                    text="للأسف، تم رفض طلب انضمامك لاستخدام البوت."
                )
            except Exception:
                pass
