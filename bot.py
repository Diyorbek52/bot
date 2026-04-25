import os

from dotenv import load_dotenv
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from schedule_core import extract_group_schedule, render_schedule_png

load_dotenv()

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
EXCEL_FILE = os.getenv("EXCEL_FILE", "Расписание.xlsx").strip()
SHEET_NAME = os.getenv("SHEET_NAME", "").strip() or None


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Пришли название группы так, как оно в шапке Excel "
        "(можно часть текста — найду по вхождению).\n\n"
        "Пример: `И-9-2025 (О)" \
        "`",
        parse_mode="Markdown",
    )


async def on_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.message.text:
        return

    q = update.message.text.strip()
    await update.message.reply_text("Собираю расписание...")

    try:
        rows = extract_group_schedule(EXCEL_FILE, q, sheet_name=SHEET_NAME)
        title = f"Расписание: {q}"
        path = render_schedule_png(rows, title)
        with open(path, "rb") as f:
            await context.bot.send_photo(
                chat_id=update.effective_chat.id,
                photo=f,
                caption=title,
            )
    except Exception as e:
        await update.message.reply_text(str(e))


def main() -> None:
    if not TELEGRAM_BOT_TOKEN:
        raise RuntimeError(
            "Не задан TELEGRAM_BOT_TOKEN.\n"
            "Создай файл `.env` и добавь: TELEGRAM_BOT_TOKEN=...\n"
            "Для работы без Telegram запусти: python local_bot.py"
        )

    app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))

    print("Бот в Telegram запущен. Ctrl+C — остановить.")
    app.run_polling()


if __name__ == "__main__":
    main()
