import os
import base64
import logging
import hashlib
import anthropic
import openpyxl
from datetime import datetime, time
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, JobQueue

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
ADMIN_ID = 1818355998
THRESHOLD_KG = 0.300

# Хранилище за день
daily_stats = {
    "date": datetime.now().strftime("%d.%m.%Y"),
    "unique_hashes": set(),
    "total_unique": 0,
    "with_discrepancy": [],
    "duplicates": 0
}

def reset_daily_stats():
    daily_stats["date"] = datetime.now().strftime("%d.%m.%Y")
    daily_stats["unique_hashes"] = set()
    daily_stats["total_unique"] = 0
    daily_stats["with_discrepancy"] = []
    daily_stats["duplicates"] = 0

def make_hash(items_dict):
    sorted_items = sorted(items_dict.items())
    hash_str = str(sorted_items)
    return hashlib.md5(hash_str.encode()).hexdigest()

def load_products():
    wb = openpyxl.load_workbook("products.xlsx")
    ws = wb.active
    products = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        kod = str(row[0]).strip().lower()
        name = str(row[1]).strip() if row[1] else ""
        try:
            weight = float(str(row[2]).replace(",", "."))
        except:
            weight = 0
        products[kod] = {"name": name, "weight": weight}
    return products

PRODUCTS = load_products()

async def send_daily_report(context):
    today = datetime.now().strftime("%d.%m.%Y")
    if daily_stats["date"] != today:
        reset_daily_stats()
        return

    total = daily_stats["total_unique"]
    disc_count = len(daily_stats["with_discrepancy"])
    dupes = daily_stats["duplicates"]

    report = f"📊 *Итог за {today}:*\n\n"
    report += f"✅ Уникальных накладных: *{total}*\n"
    report += f"🔄 Дубликатов отклонено: *{dupes}*\n"
    report += f"🚨 С расхождением: *{disc_count}*\n"

    if daily_stats["with_discrepancy"]:
        report += "\n*Накладные с расхождением:*\n"
        for item in daily_stats["with_discrepancy"]:
            report += f"   • {item}\n"

    if total == 0:
        report = f"📊 *Итог за {today}:*\nНакладных сегодня не было."

    await context.bot.send_message(chat_id=ADMIN_ID, text=report, parse_mode="Markdown")
    reset_daily_stats()

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! 👋\n\n"
        "Я считаю вес товаров по накладной.\n\n"
        "📸 Отправь фото накладной — посчитаю вес и проверю расхождение!\n"
        "🕙 Каждый день в 23:00 отправляю сводку администратору."
    )

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Сброс статистики если новый день
    today = datetime.now().strftime("%d.%m.%Y")
    if daily_stats["date"] != today:
        reset_daily_stats()

    await update.message.reply_text("📷 Получил фото, анализирую накладную...")
    try:
        photo = update.message.photo[-1]
        file = await context.bot.get_file(photo.file_id)
        photo_bytes = await file.download_as_bytearray()
        photo_b64 = base64.standard_b64encode(photo_bytes).decode("utf-8")

        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

        prompt = """На этом фото — накладная с товарами. 

1. Найди все строки с товарами и извлеки код товара и количество.
2. Посмотри — есть ли на фото рукописное число обозначающее фактический вес (обычно написано от руки внизу накладной).

Верни ТОЛЬКО в таком формате:

ТОВАРЫ:
КОД|КОЛИЧЕСТВО

ФАКТ_ВЕС:
число (если есть рукописный вес, иначе напиши НЕТ)

Пример:
ТОВАРЫ:
p-121|1
p-122|2

ФАКТ_ВЕС:
13.18

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

        # Парсим ответ
        товары_секция = ""
        факт_вес_секция = "НЕТ"

        if "ТОВАРЫ:" in raw_text and "ФАКТ_ВЕС:" in raw_text:
            parts = raw_text.split("ФАКТ_ВЕС:")
            товары_секция = parts[0].replace("ТОВАРЫ:", "").strip()
            факт_вес_секция = parts[1].strip()
        else:
            товары_секция = raw_text

        # Считаем расчётный вес
        lines = товары_секция.strip().split("\n")
        results = []
        total_weight = 0.0
        not_found = []
        items_dict = {}

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
                items_dict[kod_raw] = qty
                name_short = product["name"].replace("/", "").strip()
                if len(name_short) > 35:
                    name_short = name_short[:35] + "..."
                results.append(f"✅ {kod_raw.upper()} × {qty} шт = {weight:.3f} кг\n   {name_short}")
            else:
                not_found.append(f"❓ {kod_raw.upper()} × {qty} — не найден")

        if not results and not not_found:
            await update.message.reply_text("⚠️ Не удалось распознать товары. Попробуй сфотографировать чётче.")
            return

        # Проверка дубликата
        if items_dict:
            nav_hash = make_hash(items_dict)
            if nav_hash in daily_stats["unique_hashes"]:
                daily_stats["duplicates"] += 1
                await update.message.reply_text(
                    "🔄 *Эта накладная уже была обработана сегодня — дубликат!*\n"
                    f"Расчётный вес: {total_weight:.3f} кг",
                    parse_mode="Markdown"
                )
                return
            else:
                daily_stats["unique_hashes"].add(nav_hash)
                daily_stats["total_unique"] += 1

        # Формируем ответ
        reply = "📦 *Результат расчёта веса:*\n\n"
        if results:
            reply += "\n".join(results)
            reply += f"\n\n{'─'*30}\n"
            reply += f"⚖️ *РАСЧЁТНЫЙ ВЕС: {total_weight:.3f} кг*"

        # Сравниваем с фактическим весом
        has_discrepancy = False
        if факт_вес_секция and факт_вес_секция.upper() != "НЕТ":
            try:
                fact_weight = float(факт_вес_секция.replace(",", ".").strip())
                diff = abs(fact_weight - total_weight)
                reply += f"\n📝 *ФАКТИЧЕСКИЙ ВЕС: {fact_weight:.3f} кг*"
                reply += f"\n📊 *РАЗНИЦА: {diff:.3f} кг*"

                if diff > THRESHOLD_KG:
                    has_discrepancy = True
                    reply += f"\n\n🚨 *РАСХОЖДЕНИЕ {diff:.3f} кг — ПРОВЕРИТЬ!*"
                    if fact_weight > total_weight:
                        reply += f"\n   Фактический БОЛЬШЕ расчётного на {diff:.3f} кг"
                    else:
                        reply += f"\n   Фактический МЕНЬШЕ расчётного на {diff:.3f} кг"
                    # Добавляем в статистику
                    user = update.message.from_user
                    name = user.first_name or "Неизвестный"
                    daily_stats["with_discrepancy"].append(
                        f"{name}: расч. {total_weight:.3f} кг / факт. {fact_weight:.3f} кг / разница {diff:.3f} кг"
                    )
                else:
                    reply += f"\n\n✅ *Расхождение в норме*"
            except:
                pass

        if not_found:
            reply += "\n\n" + "\n".join(not_found)

        await update.message.reply_text(reply, parse_mode="Markdown")

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

    # Ежедневная сводка в 23:00
    job_queue = app.job_queue
    job_queue.run_daily(send_daily_report, time=time(hour=23, minute=0))

    logger.info("Бот запущен!")
    app.run_polling()

if __name__ == "__main__":
    main()
