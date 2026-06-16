import os
import base64
import logging
import requests
import anthropic
import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")

# ID админа и токен форекс-бота для уведомлений
ADMIN_CHAT_ID = 1818355998
FOREX_BOT_TOKEN = "8728286693:AAHkUcWqPbyXtnuMp_e_7FV7IdNzUT2pcCE"

def load_products():
    df = pd.read_excel("products.xlsx")
    df.columns = ["kod", "name", "weight"]
    products = {}
    for _, row in df.iterrows():
        kod = str(row["kod"]).strip().lower()
        try:
            weight = float(str(row["weight"]).replace(",", "."))
        except:
            weight = 0
        products[kod] = {"name": str(row["name"]).strip(), "weight": weight}
    return products

PRODUCTS = load_products()

async def log_to_admin(context, user, photo_file_id, result_text):
    """Пересылает уведомление через Shah_forex_bot"""
    try:
        user_name = user.full_name or "Без имени"
        user_username = f"@{user.username}" if user.username else "нет username"
        user_id = user.id

        caption = (
            f"👤 {user_name} ({user_username})\n"
            f"🆔 ID: {user_id}\n"
            f"📸 Фото накладной\n"
            f"─────────────────\n"
            f"{result_text}"
        )

        # Скачиваем фото
        file = await context.bot.get_file(photo_file_id)
        photo_bytes = await file.download_as_bytearray()

        # Отправляем через форекс-бота
        requests.post(
            f"https://api.telegram.org/bot{FOREX_BOT_TOKEN}/sendPhoto",
            data={"chat_id": ADMIN_CHAT_ID, "caption": caption[:1024]},
            files={"photo": ("photo.jpg", bytes(photo_bytes), "image/jpeg")}
        )
    except Exception as e:
        logger.error(f"Ошибка логирования: {e}")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! 👋\n\n"
        "Я считаю вес товаров по накладной.\n\n"
        "📸 Просто отправь мне фото накладной — и я посчитаю общий вес!"
    )

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📷 Получил фото, анализирую накладную...")
    try:
        photo = update.message.photo[-1]
        file = await context.bot.get_file(photo.file_id)
        photo_bytes = await file.download_as_bytearray()
        photo_b64 = base64.standard_b64encode(photo_bytes).decode("utf-8")

        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

        prompt = """На этом фото — накладная с товарами. 
Найди все строки с товарами и извлеки:
- Код товара (например p-121, р-79, p-231 и т.д.)
- Количество (колонка "Кол-во")

Верни ТОЛЬКО в таком формате, по одной строке на товар:
КОД|КОЛИЧЕСТВО

Пример:
p-121|1
p-122|2
p-79|1

Если количество дробное (например 0,99) — оставь как есть.
Никакого другого текста не добавляй."""

        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1000,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/jpeg",
                                "data": photo_b64,
                            },
                        },
                        {"type": "text", "text": prompt}
                    ],
                }
            ],
        )

        raw_text = response.content[0].text.strip()
        lines = raw_text.strip().split("\n")
        results = []
        total_weight = 0.0
        not_found = []

        for line in lines:
            line = line.strip()
            if "|" not in line:
                continue
            parts = line.split("|")
            if len(parts) != 2:
                continue

            kod_raw = parts[0].strip().lower().replace("р-", "p-")
            try:
                qty = float(parts[1].strip().replace(",", "."))
            except:
                continue

            if kod_raw in PRODUCTS:
                product = PRODUCTS[kod_raw]
                weight = product["weight"] * qty
                total_weight += weight
                name_short = product["name"].replace("/", "").replace("не выбивать", "").replace("НЕ ВЫБИВАТЬ", "").strip()
                if len(name_short) > 35:
                    name_short = name_short[:35] + "..."
                results.append(f"✅ {kod_raw.upper()} × {qty} шт = {weight:.3f} кг\n   {name_short}")
            else:
                not_found.append(f"❓ {kod_raw.upper()} × {qty} — не найден в таблице")

        if not results and not not_found:
            await update.message.reply_text("⚠️ Не удалось распознать товары. Попробуй сфотографировать чётче.")
            return

        reply = "📦 *Результат расчёта веса:*\n\n"
        if results:
            reply += "\n".join(results)
            reply += f"\n\n{'─'*30}\n"
            reply += f"⚖️ *ОБЩИЙ ВЕС: {total_weight:.3f} кг*"
        if not_found:
            reply += "\n\n" + "\n".join(not_found)

        await update.message.reply_text(reply, parse_mode="Markdown")

        # Логируем через форекс-бота
        log_text = f"⚖️ Общий вес: {total_weight:.3f} кг"
        if not_found:
            log_text += f"\n❓ Не найдено: {len(not_found)} позиций"
        await log_to_admin(context, update.message.from_user, photo.file_id, log_text)

    except Exception as e:
        logger.error(f"Error: {e}")
        await update.message.reply_text("❌ Произошла ошибка. Попробуй ещё раз.")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📸 Отправь мне *фото накладной* — я посчитаю вес!", parse_mode="Markdown")

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    logger.info("Бот запущен!")
    app.run_polling()

if __name__ == "__main__":
    main()
