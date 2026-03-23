from typing import Final
from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
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


# === /groups в личке ===
async def groups_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.type != "private":
        return

    if not chat_names:
        await update.message.reply_text("Я ще не бачив повідомлень у групах.")
        return

    text = "Вибери групу:\n\n"
    for chat_id, name in chat_names.items():
        text += f"{name} → /get_{chat_id}\n"

    await update.message.reply_text(text)

print("DEBUG user_stats:", user_stats)
# === Надсилання Excel файлу за групою ===
async def get_group_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.type != "private":
        return

    command = update.message.text
    try:
        chat_id = int(command.replace("/get_", ""))
    except ValueError:
        await update.message.reply_text("Помилка команди")
        return

    month_key = datetime.now().strftime("%Y-%m")

    if chat_id not in user_stats or month_key not in user_stats[chat_id]:
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
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/get_"), get_group_stats))

    # обробка всіх повідомлень у групах
    app.add_handler(MessageHandler(filters.TEXT, handle_message))

    print("Бот запущен...")
    app.run_polling()


if __name__ == "__main__":
    main()