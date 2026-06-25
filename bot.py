import os
import io
import base64
import logging
import hashlib
import sqlite3
import anthropic
import openpyxl
from datetime import datetime, time
from telegram import Update, Bot, InputFile
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
FOREX_BOT_TOKEN = "8728286693:AAHkUcWqPbyXtnuMp_e_7FV7IdNzUT2pcCE"
FOREX_CHAT_ID = 1818355998
THRESHOLD_KG = 0.300
DB_PATH = "/app/nakladnye.db"

daily_stats = {
    "date": datetime.now().strftime("%d.%m.%Y"),
    "unique_hashes": set(),
    "total_unique": 0,
    "with_discrepancy": [],
    "duplicates": 0
}

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS nakladnye (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT,
        time TEXT,
        sender TEXT,
        marshut TEXT,
        calc_weight REAL,
        fact_weight REAL,
        diff REAL,
        has_discrepancy INTEGER,
        file_id TEXT
    )''')
    conn.commit()
    conn.close()

def save_nakladnaya(sender, marshut, calc_weight, fact_weight, diff, has_discrepancy, file_id):
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute('''INSERT INTO nakladnye 
            (date, time, sender, marshut, calc_weight, fact_weight, diff, has_discrepancy, file_id)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''',
            (
                datetime.now().strftime("%d.%m.%Y"),
                datetime.now().strftime("%H:%M"),
                sender, marshut or "",
                calc_weight, fact_weight, diff,
                1 if has_discrepancy else 0,
                file_id or ""
            )
        )
        conn.commit()
        conn.close()
    except Exception as e:
        logger.error(f"Ошибка сохранения в БД: {e}")

def search_nakladnye(marshut=None, date=None):
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        query = "SELECT date, time, sender, marshut, calc_weight, fact_weight, diff, has_discrepancy, file_id FROM nakladnye WHERE 1=1"
        params = []
        if marshut:
            query += " AND UPPER(marshut) LIKE UPPER(?)"
            params.append(f"%{marshut}%")
        if date:
            query += " AND date = ?"
            params.append(date)
        query += " ORDER BY id DESC LIMIT 20"
        c.execute(query, params)
        rows = c.fetchall()
        conn.close()
        return rows
    except Exception as e:
        logger.error(f"Ошибка поиска в БД: {e}")
        return []

def get_daily_stats_from_db(date):
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        # Общее количество
        c.execute("SELECT COUNT(*) FROM nakladnye WHERE date=?", (date,))
        total = c.fetchone()[0]
        # С расхождением
        c.execute("SELECT COUNT(*) FROM nakladnye WHERE date=? AND has_discrepancy=1", (date,))
        disc = c.fetchone()[0]
        # По водителям
        c.execute("SELECT sender, COUNT(*) as cnt FROM nakladnye WHERE date=? GROUP BY sender ORDER BY cnt DESC", (date,))
        by_sender = c.fetchall()
        # Расхождения детально
        c.execute("SELECT sender, marshut, calc_weight, fact_weight, diff FROM nakladnye WHERE date=? AND has_discrepancy=1", (date,))
        disc_details = c.fetchall()
        conn.close()
        return total, disc, by_sender, disc_details
    except Exception as e:
        logger.error(f"Ошибка получения статистики: {e}")
        return 0, 0, [], []

def reset_daily_stats():
    daily_stats["date"] = datetime.now().strftime("%d.%m.%Y")
    daily_stats["unique_hashes"] = set()
    daily_stats["total_unique"] = 0
    daily_stats["with_discrepancy"] = []
    daily_stats["duplicates"] = 0

def make_hash(items_dict):
    sorted_items = sorted(items_dict.items())
    return hashlib.md5(str(sorted_items).encode()).hexdigest()

def normalize_marshut(text):
    # Латиница → кириллица для похожих букв + убираем тире + всё в верхний регистр
    lat_to_cyr = {
        'A': 'А', 'B': 'В', 'C': 'С', 'E': 'Е', 'H': 'Н',
        'K': 'К', 'M': 'М', 'O': 'О', 'P': 'Р', 'T': 'Т',
        'X': 'Х', 'D': 'Д', 'V': 'В', 'G': 'Г', 'N': 'Н'
    }
    result = ""
    for ch in text.upper():
        if ch == '-':
            continue
        result += lat_to_cyr.get(ch, ch)
    return result

def make_hashtag(text):
    import re
    return re.sub(r'[^a-zA-ZА-Яа-я0-9]', '', text)

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

async def send_to_forex(text, photo_bytes=None):
    try:
        forex_bot = Bot(token=FOREX_BOT_TOKEN)
        if photo_bytes:
            await forex_bot.send_photo(
                chat_id=FOREX_CHAT_ID,
                photo=InputFile(photo_bytes, filename="nakladnaya.jpg"),
            )
        else:
            await forex_bot.send_message(
                chat_id=FOREX_CHAT_ID,
                text=text,
                parse_mode="Markdown"
            )
    except Exception as e:
        logger.error(f"Ошибка отправки в forex бот: {e}")

async def send_daily_report(context):
    today = datetime.now().strftime("%d.%m.%Y")
    total, disc, by_sender, disc_details = get_daily_stats_from_db(today)
    dupes = daily_stats["duplicates"]

    report = f"📊 *Итог за {today}:*\n\n"
    report += f"✅ Всего накладных: *{total}*\n"
    report += f"🔄 Дубликатов отклонено: *{dupes}*\n"
    report += f"🚨 С расхождением: *{disc}*\n"

    if by_sender:
        report += f"\n👥 *По водителям:*\n"
        for sender, cnt in by_sender:
            report += f"   • {sender} — {cnt} накладных\n"

    if disc_details:
        report += f"\n🚨 *Расхождения:*\n"
        for sender, marshut, calc, fact, diff in disc_details:
            report += f"   • {sender} [{marshut}]: расч. {calc:.3f} / факт. {fact:.3f} / разница {diff:.3f} кг\n"

    if total == 0:
        report = f"📊 *Итог за {today}:*\nНакладных сегодня не было."

    await send_to_forex(report)
    reset_daily_stats()

async def cmd_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    today = datetime.now().strftime("%d.%m.%Y")
    total, disc, by_sender, disc_details = get_daily_stats_from_db(today)
    dupes = daily_stats["duplicates"]

    report = f"📊 *Текущий итог за {today}:*\n\n"
    report += f"✅ Всего накладных: *{total}*\n"
    report += f"🔄 Дубликатов отклонено: *{dupes}*\n"
    report += f"🚨 С расхождением: *{disc}*\n"

    if by_sender:
        report += f"\n👥 *По водителям:*\n"
        for sender, cnt in by_sender:
            report += f"   • {sender} — {cnt} накладных\n"

    if total == 0:
        report = f"📊 *Итог за {today}:*\nНакладных пока не было."

    await update.message.reply_text(report, parse_mode="Markdown")

async def cmd_find(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if not args:
        await update.message.reply_text(
            "Использование:\n"
            "/find B14 — все накладные маршрута B14\n"
            "/find B14 23.06.2026 — за конкретную дату"
        )
        return

    marshut = normalize_marshut(args[0])
    date = args[1] if len(args) > 1 else None
    rows = search_nakladnye(marshut=marshut, date=date)

    if not rows:
        await update.message.reply_text(
            f"📭 Накладных по маршруту *{marshut}* не найдено.",
            parse_mode="Markdown"
        )
        return

    await update.message.reply_text(
        f"🔍 *Найдено по маршруту {marshut}: {len(rows)} шт*",
        parse_mode="Markdown"
    )

    for row in rows:
        date_r, time_r, sender, marshut_r, calc, fact, diff, disc, file_id = row
        caption = f"📅 {date_r} {time_r} | 👤 {sender}\n"
        caption += f"📍 Маршрут: {marshut_r}\n"
        caption += f"⚖️ Расч: {calc:.3f} кг"
        if fact:
            caption += f" | 📝 Факт: {fact:.3f} кг"
        if diff:
            caption += f" | 📊 Разница: {diff:.3f} кг"
        if disc:
            caption += f"\n🚨 РАСХОЖДЕНИЕ!"

        if file_id:
            try:
                await update.message.reply_photo(photo=file_id, caption=caption)
            except Exception as e:
                logger.error(f"Ошибка отправки фото: {e}")
                await update.message.reply_text(caption)
        else:
            await update.message.reply_text(caption)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! 👋\n\n"
        "Я считаю вес товаров по накладной.\n\n"
        "📸 Отправь фото накладной — посчитаю вес!\n"
        "📊 /report — статистика за день\n"
        "🔍 /find B14 — поиск по маршруту\n"
        "🔍 /find B14 23.06.2026 — поиск за дату"
    )

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    today = datetime.now().strftime("%d.%m.%Y")
    if daily_stats["date"] != today:
        reset_daily_stats()

    await update.message.reply_text("📷 Получил фото, анализирую накладную...")
    try:
        photo = update.message.photo[-1]
        file_id = photo.file_id
        file = await context.bot.get_file(file_id)
        photo_bytes = await file.download_as_bytearray()
        photo_b64 = base64.standard_b64encode(photo_bytes).decode("utf-8")

        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

        prompt = """На этом фото — накладная с товарами. Водитель пишет от руки два значения внизу накладной:
1. Фактический вес — число с запятой или точкой, например: 4,31 или 12.5
2. Код маршрута — буква + цифры, например: А21, D15, B14

Найди все строки с товарами и извлеки код товара и количество.
Также найди рукописный вес и код маршрута.

Верни ТОЛЬКО в таком формате:

ТОВАРЫ:
КОД|КОЛИЧЕСТВО

ФАКТ_ВЕС:
число (если есть, иначе НЕТ)

МАРШРУТ:
код маршрута (если есть, иначе НЕТ)

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

        товары_секция = ""
        факт_вес_секция = "НЕТ"
        маршрут_секция = "НЕТ"

        if "ТОВАРЫ:" in raw_text:
            parts = raw_text.split("ФАКТ_ВЕС:")
            товары_секция = parts[0].replace("ТОВАРЫ:", "").strip()
            if len(parts) > 1:
                rest = parts[1]
                if "МАРШРУТ:" in rest:
                    wparts = rest.split("МАРШРУТ:")
                    факт_вес_секция = wparts[0].strip()
                    маршрут_секция = wparts[1].strip()
                else:
                    факт_вес_секция = rest.strip()

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

        user = update.message.from_user
        sender_name = f"{user.first_name or ''} {user.last_name or ''}".strip() or "Неизвестный"
        now_time = datetime.now().strftime("%H:%M")
        marshut = normalize_marshut(маршрут_секция.strip()) if маршрут_секция.upper() != "НЕТ" else None

        reply = "📦 *Результат расчёта веса:*\n\n"
        if results:
            reply += "\n".join(results)
            reply += f"\n\n{'─'*30}\n"
            reply += f"⚖️ *РАСЧЁТНЫЙ ВЕС: {total_weight:.3f} кг*"
        if marshut:
            reply += f"\n📍 *Маршрут: {marshut}*"

        has_discrepancy = False
        fact_weight = None
        diff = None

        if факт_вес_секция and факт_вес_секция.upper() != "НЕТ":
            try:
                fact_weight = float(факт_вес_секция.replace(",", ".").strip())
                diff = abs(fact_weight - total_weight)
                reply += f"\n📝 *ФАКТИЧЕСКИЙ ВЕС: {fact_weight:.3f} кг*"
                reply += f"\n📊 *РАЗНИЦА: {diff:.3f} кг*"
                if diff > THRESHOLD_KG:
                    has_discrepancy = True
                    reply += f"\n\n🚨🔦 *РАСХОЖДЕНИЕ {diff:.3f} кг — ПРОВЕРИТЬ!*"
                    if fact_weight > total_weight:
                        reply += f"\n   Фактический БОЛЬШЕ расчётного на {diff:.3f} кг"
                    else:
                        reply += f"\n   Фактический МЕНЬШЕ расчётного на {diff:.3f} кг"
                    daily_stats["with_discrepancy"].append(
                        f"{sender_name} [{marshut or '?'}]: расч. {total_weight:.3f} / факт. {fact_weight:.3f} / разница {diff:.3f} кг"
                    )
                else:
                    reply += f"\n\n✅ *Расхождение в норме*"
            except:
                pass

        if not_found:
            reply += "\n\n" + "\n".join(not_found)

        save_nakladnaya(sender_name, marshut, total_weight, fact_weight, diff, has_discrepancy, file_id)
        await update.message.reply_text(reply, parse_mode="Markdown")

        name_tag = make_hashtag(sender_name)
        marshut_tag = make_hashtag(marshut) if marshut else ""
        hashtags = f"#{marshut_tag} #{name_tag}" if marshut_tag else f"#{name_tag}"

        admin_msg = f"{hashtags}\n"
        admin_msg += f"🕐 *{now_time}* | 👤 *{sender_name}*\n"
        if marshut:
            admin_msg += f"📍 *Маршрут: {marshut}*\n"
        admin_msg += f"⚖️ Расчётный вес: *{total_weight:.3f} кг*\n"
        if fact_weight is not None:
            admin_msg += f"📝 Фактический: *{fact_weight:.3f} кг*\n"
            admin_msg += f"📊 Разница: *{diff:.3f} кг*\n"
            if has_discrepancy:
                admin_msg += f"\n🚨🔦 *РАСХОЖДЕНИЕ! Проверить накладную!*"
            else:
                admin_msg += f"✅ В норме"
        else:
            admin_msg += "📝 Фактический вес не указан"
        admin_msg += f"\n\n📊 Накладных сегодня: *{daily_stats['total_unique']}*"

        await send_to_forex(None, io.BytesIO(bytes(photo_bytes)))
        await send_to_forex(admin_msg)

    except Exception as e:
        logger.error(f"Error: {e}")
        await update.message.reply_text("❌ Произошла ошибка. Попробуй ещё раз.")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📸 Отправь мне *фото накладной* — я посчитаю вес!", parse_mode="Markdown")

def main():
    init_db()
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("report", cmd_report))
    app.add_handler(CommandHandler("find", cmd_find))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    job_queue = app.job_queue
    job_queue.run_daily(send_daily_report, time=time(hour=23, minute=0))
    logger.info("Бот запущен!")
    app.run_polling()

if __name__ == "__main__":
    main()
