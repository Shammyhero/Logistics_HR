import asyncio
import logging
import os
from dotenv import load_dotenv

from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart, Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    CallbackQuery,
    Message,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    ReplyKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardRemove,
    FSInputFile
)
import sqlite3
import pandas as pd
from datetime import datetime

# Load environment variables
load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_CHAT_ID = os.getenv("ADMIN_CHAT_ID")

DB_FILE = os.getenv("DB_FILE_PATH", "applications.db")

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS applications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            phone TEXT,
            position TEXT,
            username TEXT,
            telegram_id INTEGER,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

init_db()

class ApplicationForm(StatesGroup):
    name = State()
    phone = State()
    position = State()

dp = Dispatcher()

@dp.message(CommandStart())
async def command_start_handler(message: Message, state: FSMContext) -> None:
    """
    This handler receives messages with `/start` command
    """
    await state.clear() # Clear any previous state
    
    welcome_message = (
        "Hi there! 👋\n"
        "Looking for better loads and pay?\n\n"
        "I’ll guide you through a quick application for Deload Logistics 🚛\n\n"
        "Tap below to begin."
    )
    
    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="Begin Application", callback_data="start_app")]
        ]
    )
    
    # Path to the image file
    image_path = os.path.join(os.path.dirname(__file__), "images", "image.png")
    
    # Check if image exists before sending, fallback to text if not
    if os.path.exists(image_path):
        photo = FSInputFile(image_path)
        await message.answer_photo(photo=photo, caption=welcome_message, reply_markup=keyboard)
    else:
        await message.answer(text=welcome_message, reply_markup=keyboard)


@dp.callback_query(F.data == "start_app")
async def start_application(callback: CallbackQuery, state: FSMContext):
    await state.set_state(ApplicationForm.name)
    await callback.message.answer("1. Please enter your Full Name:")
    await callback.answer()

@dp.message(ApplicationForm.name)
async def process_name(message: Message, state: FSMContext):
    # Save name to state
    await state.update_data(name=message.text)
    
    # Transition to phone number step
    await state.set_state(ApplicationForm.phone)
    
    # Create contact share button
    kb = [
        [KeyboardButton(text="Share Contact", request_contact=True)]
    ]
    keyboard = ReplyKeyboardMarkup(keyboard=kb, resize_keyboard=True, one_time_keyboard=True)
    
    await message.answer("2. Please share your phone number using the button below:", reply_markup=keyboard)

@dp.message(ApplicationForm.phone)
async def process_phone(message: Message, state: FSMContext):
    # Check if user shared via contact button or typed it
    phone_number = message.contact.phone_number if message.contact else message.text
    await state.update_data(phone=phone_number)
    
    # Transition to position step
    await state.set_state(ApplicationForm.position)
    
    # Present position choices
    kb = [
        [KeyboardButton(text="Company driver")],
        [KeyboardButton(text="Owner Operator")],
        [KeyboardButton(text="Lease Operator")]
    ]
    keyboard = ReplyKeyboardMarkup(keyboard=kb, resize_keyboard=True, one_time_keyboard=True)
    
    await message.answer("3. What is your position?", reply_markup=keyboard)

@dp.message(ApplicationForm.position)
async def process_position(message: Message, state: FSMContext):
    # Save the final position
    await state.update_data(position=message.text)
    
    # Retrieve all collected data
    user_data = await state.get_data()
    name = user_data.get('name')
    phone = user_data.get('phone')
    position = user_data.get('position')
    username = f"@{message.from_user.username}" if message.from_user.username else "No username"
    
    # Notify User it was successful
    success_text = "Thank you! Your application has been submitted successfully. We will get in touch with you shortly."
    await message.answer(success_text, reply_markup=ReplyKeyboardRemove())
    
    # Save to database
    try:
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute('''
            INSERT INTO applications (name, phone, position, username, telegram_id)
            VALUES (?, ?, ?, ?, ?)
        ''', (name, phone, position, username, message.from_user.id))
        conn.commit()
        conn.close()
    except Exception as e:
        logging.error(f"Failed to save to database: {e}")

    # Notify Admin if ADMIN_CHAT_ID is set
    admin_notification = (
        "🚛 **New Applicant!**\n\n"
        f"**Name:** {name}\n"
        f"**Phone:** {phone}\n"
        f"**Position:** {position}\n"
        f"**Telegram User:** {username} (ID: {message.from_user.id})"
    )
    
    if ADMIN_CHAT_ID and ADMIN_CHAT_ID != "your_admin_chat_id_here":
        try:
            bot = message.bot
            await bot.send_message(chat_id=ADMIN_CHAT_ID, text=admin_notification, parse_mode="Markdown")
        except Exception as e:
            logging.error(f"Failed to send application to admin: {e}")
    else:
        logging.warning("No ADMIN_CHAT_ID set. Application was not forwarded anywhere.")
        print("--- APPLICATION RECEIVED ---")
        print(admin_notification)
        print("----------------------------")
        
    # Clear state fully
    await state.clear()


@dp.message(Command("exporttogr"))
async def command_export_handler(message: Message) -> None:
    # Restrict to Admin Chat
    if str(message.chat.id) != str(ADMIN_CHAT_ID):
        return

    # Check if user is admin
    try:
        chat_member = await message.bot.get_chat_member(message.chat.id, message.from_user.id)
        if chat_member.status not in ['administrator', 'creator']:
            await message.reply("You must be an admin to use this command.")
            return
    except Exception as e:
         logging.error(f"Failed to verify admin status: {e}")
         return

    # Extract all from DB to Excel
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM applications ORDER BY timestamp DESC", conn)
    conn.close()
    
    if df.empty:
        await message.reply("No applications found yet.")
        return

    # Generate Excel file
    excel_file = "deload_applications.xlsx"
    df.to_excel(excel_file, index=False)
    
    # Send file to group
    document = FSInputFile(excel_file)
    await message.reply_document(document=document, caption="Here are the latest applications:")
    
    # Clean up file
    if os.path.exists(excel_file):
        os.remove(excel_file)


async def main() -> None:
    if not BOT_TOKEN or BOT_TOKEN == "your_telegram_bot_token_here":
        logging.error("BOT_TOKEN is missing or invalid in your .env file!")
        return

    bot = Bot(token=BOT_TOKEN)
    
    logging.info("Starting bot polling...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    asyncio.run(main())
