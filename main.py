from typing import Final
from telegram import Update, InlineKeyboardMarkup, InlineKeyboardButton
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    filters,
    ContextTypes,
)
from datetime import datetime
from openpyxl import Workbook

# === Налаштовунання ===
TOKEN: Final = "8660996297:AAGmveqCnaQxDkegzMQ1BAMq4Ic1dF_kEkg"

# === Дані ===
user_stats = {}   # chat_id -> month -> user_id -> {"name": username, "count": int}
chat_names = {}   # chat_id -> назва групи


# === /start ===
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привіт! Я вважаю повідомлення у групах📊\n"
     
    )


# === Рахунок повідомлень у групі ===
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message:
        return
    
    print("Повідомлення надійшло:", update.message.text)  


    chat = update.effective_chat

    # ігноруємо особисті повідомлення
    if chat.type == "private":
        return

    # лог для перевірки, що повідомлення надходять
    print(f"[LOG] Повідомлення від {update.message.from_user.full_name} в {chat.title}: {update.message.text}")

    user = update.message.from_user
    chat_id = chat.id
    chat_names[chat_id] = chat.title

    user_id = user.id
    username = user.full_name

    month_key = datetime.now().strftime("%Y-%m")

    if chat_id not in user_stats:
        user_stats[chat_id] = {}
    if month_key not in user_stats[chat_id]:
        user_stats[chat_id][month_key] = {}
    if user_id not in user_stats[chat_id][month_key]:
        user_stats[chat_id][month_key][user_id] = {"name": username, "count": 0}

    user_stats[chat_id][month_key][user_id]["count"] += 1


# === Перевірка, чи користувач адміністратор групи ===
async def is_group_admin(user_id: int, chat_id: int, context: ContextTypes.DEFAULT_TYPE) -> bool:
    try:
        member = await context.bot.get_chat_member(chat_id, user_id)
        # Користувач адміністратор, якщо він Creator або Administrator
        return member.status in ["creator", "administrator"]
    except Exception as e:
        print(f"[ERROR] Помилка при перевірці статусу адміністратора: {e}")
        return False

# === /groups в личке ===
async def groups_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.type != "private":
        return

    if not chat_names:
        await update.message.reply_text("Я ще не бачив повідомлень у групах.")
        return

    user_id = update.effective_user.id
    
    # Фільтруємо групи, де користувач є адміністратором
    admin_groups = []
    for chat_id, name in chat_names.items():
        if await is_group_admin(user_id, chat_id, context):
            admin_groups.append((chat_id, name))
    
    if not admin_groups:
        await update.message.reply_text("Ви не є адміністратором жодної групи, де я записую статистику.")
        return

    # Створюємо inline кнопки для груп, де користувач адміністратор
    keyboard = []
    for chat_id, name in admin_groups:
        button = InlineKeyboardButton(text=name, callback_data=f"group_{chat_id}")
        keyboard.append([button])
    
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("Вибери групу (твої групи):"  , reply_markup=reply_markup)

# === Обробка нажаття на кнопку групи ===
async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    if not query.data.startswith("group_"):
        return
    
    user_id = query.from_user.id
    
    try:
        chat_id = int(query.data.replace("group_", ""))
        print(f"[DEBUG] Користувач {user_id} вибрав групу: {chat_id}")
    except ValueError as e:
        print(f"[ERROR] Помилка парсингу callback_data: {e}")
        await query.edit_message_text("Помилка команди")
        return
    
    # Перевіряємо, чи користувач адміністратор цієї групи
    is_admin = await is_group_admin(user_id, chat_id, context)
    if not is_admin:
        print(f"[ERROR] Користувач {user_id} не є адміністратором групи {chat_id}")
        await query.edit_message_text("❌ У вас немає доступу до статистики цієї групи. Ви повинні бути адміністратором.")
        return
    
    month_key = datetime.now().strftime("%Y-%m")
    print(f"[DEBUG] Поточний місяць: {month_key}, Наявні групи: {list(user_stats.keys())}")

    if chat_id not in user_stats:
        print(f"[ERROR] chat_id {chat_id} не знайдено в user_stats")
        await query.edit_message_text("Немає даних щодо цієї групи.")
        return
    
    if month_key not in user_stats[chat_id]:
        print(f"[ERROR] Місяць {month_key} не знайдено для {chat_id}. Наявні місяці: {list(user_stats[chat_id].keys())}")
        await query.edit_message_text("Немає даних щодо цієї групи.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Stats"

    ws.append(["User", "Messages"])
    for user in user_stats[chat_id][month_key].values():
        ws.append([user["name"], user["count"]])

    file_name = f"stats_{chat_id}_{month_key}.xlsx"
    wb.save(file_name)

    with open(file_name, "rb") as file:
        await query.message.reply_document(file)

# === Надсилання Excel файлу за групою ===
async def get_group_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.type != "private":
        return

    user_id = update.effective_user.id
    command = update.message.text
    print(f"[DEBUG] Команда отримана: {command}")
    
    try:
        # Парсимо ID з команди /get_{chat_id}
        chat_id_str = command.replace("/get_", "").strip()
        chat_id = int(chat_id_str)
        print(f"[DEBUG] Розпарсений chat_id: {chat_id}")
    except ValueError as e:
        print(f"[ERROR] Помилка парсингу: {e}")
        await update.message.reply_text("Помилка команди")
        return
    
    # Перевіряємо, чи користувач адміністратор цієї групи
    is_admin = await is_group_admin(user_id, chat_id, context)
    if not is_admin:
        print(f"[ERROR] Користувач {user_id} не є адміністратором групи {chat_id}")
        await update.message.reply_text("❌ У вас немає доступу до статистики цієї групи. Ви повинні бути адміністратором.")
        return

    month_key = datetime.now().strftime("%Y-%m")
    print(f"[DEBUG] Поточний місяць: {month_key}, Наявні групи: {list(user_stats.keys())}")

    if chat_id not in user_stats:
        print(f"[ERROR] chat_id {chat_id} не знайдено в user_stats")
        await update.message.reply_text("Немає даних щодо цієї групи.")
        return
    
    if month_key not in user_stats[chat_id]:
        print(f"[ERROR] Місяць {month_key} не знайдено для {chat_id}. Наявні місяці: {list(user_stats[chat_id].keys())}")
        await update.message.reply_text("Немає даних щодо цієї групи.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Stats"

    ws.append(["User", "Messages"])
    for user in user_stats[chat_id][month_key].values():
        ws.append([user["name"], user["count"]])

    file_name = f"stats_{chat_id}_{month_key}.xlsx"
    wb.save(file_name)

    with open(file_name, "rb") as file:
        await update.message.reply_document(file)


# === Запуск бота ===
def main():
    app = Application.builder().token(TOKEN).build()

    # команди
    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("groups", groups_command))
    app.add_handler(CallbackQueryHandler(button_callback))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/get_-?\d+$"), get_group_stats))

    # обробка всіх повідомлень у групах
    app.add_handler(MessageHandler(filters.TEXT, handle_message))

    print("Бот запущен...")
    app.run_polling()


if __name__ == "__main__":
    main()