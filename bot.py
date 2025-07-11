import os
import sqlite3
from datetime import datetime
import pytz
from aiogram import Bot, Dispatcher, executor, types
from dotenv import load_dotenv
import logging
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# Настройка логов
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    filename="data/bot.log"
)
logger = logging.getLogger(__name__)

# Загружаем токен
load_dotenv()
TOKEN = os.getenv("BOT_TOKEN")
if not TOKEN:
    logger.error("Токен не найден! Проверьте файл .env")
    exit(1)

# Настройки
ADMINS = [376996574]  # Замените на ваш ID (узнать через @userinfobot)
MOSCOW_TZ = pytz.timezone('Europe/Moscow')

# Инициализация бота
bot = Bot(token=TOKEN)
dp = Dispatcher(bot)

# База данных
def get_db():
    conn = sqlite3.connect('data/shifts.db', check_same_thread=False)
    cursor = conn.cursor()
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS shifts
                    (id INTEGER PRIMARY KEY AUTOINCREMENT,
                     user_id INTEGER NOT NULL,
                     user_name TEXT NOT NULL,
                     start_time TEXT NOT NULL,
                     end_time TEXT,
                     comment TEXT)''')
    conn.commit()
    return conn, cursor

conn, cursor = get_db()

# Утилиты
def current_time():
    return datetime.now(MOSCOW_TZ).strftime("%Y-%m-%d %H:%M:%S")

def format_time(db_time):
    if not db_time:
        return None
    dt = datetime.strptime(db_time, "%Y-%m-%d %H:%M:%S")
    return dt.astimezone(MOSCOW_TZ)

# ================= КОМАНДЫ ДЛЯ ВСЕХ =================
@dp.message_handler(commands=['start'])
async def start(message: types.Message):
    text = """
👋 Привет! Я бот для учета рабочих смен (Московское время).

📌 Основные команды:
/start_shift - Начать смену
/end_shift - Закончить смену
/my_shifts - Мои смены
    """
    await message.answer(text)

@dp.message_handler(commands=['start_shift'])
async def start_shift(message: types.Message):
    user_id = message.from_user.id
    user_name = message.from_user.full_name
    
    cursor.execute("SELECT * FROM shifts WHERE user_id=? AND end_time IS NULL", (user_id,))
    if cursor.fetchone():
        await message.answer("⚠️ У вас уже есть активная смена!")
        return
    
    cursor.execute("INSERT INTO shifts (user_id, user_name, start_time) VALUES (?, ?, ?)", 
                  (user_id, user_name, current_time()))
    conn.commit()
    await message.answer(f"✅ Смена начата в {format_time(current_time()).strftime('%H:%M %d.%m.%Y')}")

@dp.message_handler(commands=['end_shift'])
async def end_shift(message: types.Message):
    user_id = message.from_user.id
    
    cursor.execute("SELECT id, start_time FROM shifts WHERE user_id=? AND end_time IS NULL", (user_id,))
    shift = cursor.fetchone()
    
    if not shift:
        await message.answer("ℹ️ У вас нет активной смены!")
        return
    
    cursor.execute("UPDATE shifts SET end_time=? WHERE id=?", (current_time(), shift[0]))
    conn.commit()
    
    start_dt = format_time(shift[1])
    end_dt = format_time(current_time())
    duration = end_dt - start_dt
    
    await message.answer(
        f"⏹ Смена завершена в {end_dt.strftime('%H:%M %d.%m.%Y')}\n"
        f"⏱ Длительность: {duration}"
    )

@dp.message_handler(commands=['my_shifts'])
async def my_shifts(message: types.Message):
    user_id = message.from_user.id
    cursor.execute("""
        SELECT start_time, end_time 
        FROM shifts 
        WHERE user_id=?
        ORDER BY start_time DESC
        LIMIT 10
    """, (user_id,))
    
    shifts = cursor.fetchall()
    
    if not shifts:
        await message.answer("📭 У вас еще нет смен")
        return
    
    text = "📅 Ваши последние смены (МСК):\n\n"
    for i, (start, end) in enumerate(shifts, 1):
        start_dt = format_time(start)
        text += f"{i}. 🟢 {start_dt.strftime('%H:%M %d.%m.%Y')}\n"
        
        if end:
            end_dt = format_time(end)
            duration = end_dt - start_dt
            text += f"   🔴 {end_dt.strftime('%H:%M %d.%m.%Y')}\n"
            text += f"   ⏱ {duration}\n\n"
        else:
            text += "   🟠 В процессе\n\n"
    
    await message.answer(text)

# ================= АДМИН-КОМАНДЫ (EXCEL) =================
@dp.message_handler(commands=['export'])
async def export_shifts(message: types.Message):
    if message.from_user.id not in ADMINS:
        await message.answer("⛔ Доступ запрещен")
        return

    try:
        # Получаем данные
        cursor.execute("""
            SELECT 
                user_id,
                user_name,
                strftime('%d.%m.%Y %H:%M', start_time, 'localtime') as start,
                CASE WHEN end_time IS NULL 
                     THEN 'В процессе' 
                     ELSE strftime('%d.%m.%Y %H:%M', end_time, 'localtime') 
                END as end,
                CASE WHEN end_time IS NULL 
                     THEN '-' 
                     ELSE round((julianday(end_time) - julianday(start_time)) * 24, 2) 
                END as hours
            FROM shifts
            ORDER BY start_time DESC
        """)
        
        shifts = cursor.fetchall()
        
        if not shifts:
            await message.answer("📭 Нет данных о сменах")
            return

        # Создаем Excel-файл
        filename = f"data/shifts_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Смены"
        
        # Заголовки
        headers = ["ID", "ФИО", "Начало смены", "Конец смены", "Часы"]
        ws.append(headers)
        
        # Стили для заголовков
        bold_font = Font(bold=True)
        center_alignment = Alignment(horizontal='center')
        for cell in ws[1]:
            cell.font = bold_font
            cell.alignment = center_alignment
        
        # Данные
        for shift in shifts:
            ws.append(shift)
        
        # Авто-ширина колонок
        for col in ws.columns:
            column = col[0].column_letter
            max_length = max(len(str(cell.value)) for cell in col)
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column].width = adjusted_width
        
        # Сохраняем файл
        wb.save(filename)
        
        # Отправляем файл
        with open(filename, 'rb') as file:
            await message.answer_document(
                document=file,
                caption=f"📊 Отчет по сменам ({len(shifts)} записей)"
            )
            
    except Exception as e:
        await message.answer(f"⚠️ Ошибка: {str(e)}")
        logger.exception("Ошибка экспорта:")

@dp.message_handler(commands=['stats'])
async def show_stats(message: types.Message):
    if message.from_user.id not in ADMINS:
        await message.answer("⛔ Доступ запрещен")
        return

    cursor.execute("""
        SELECT 
            user_name,
            COUNT(*) as total_shifts,
            SUM(CASE WHEN end_time IS NULL THEN 1 ELSE 0 END) as active_shifts,
            ROUND(SUM((julianday(end_time) - julianday(start_time)) * 24, 1) as total_hours
        FROM shifts
        GROUP BY user_id
        ORDER BY total_shifts DESC
    """)
    
    stats = cursor.fetchall()
    
    text = "📈 <b>Статистика по сотрудникам:</b>\n\n"
    for user in stats:
        text += f"👤 <b>{user[0]}</b>\n"
        text += f"▪ Всего смен: {user[1]}\n"
        text += f"▪ Активных: {user[2]}\n"
        text += f"▪ Всего часов: {user[3] or 0}\n\n"
    
    await message.answer(text, parse_mode="HTML")

# ================= ЗАПУСК =================
if __name__ == '__main__':
    # Создаем папки
    os.makedirs('data', exist_ok=True)
    
    logger.info("🔄 Бот запускается...")
    executor.start_polling(dp, skip_updates=True)