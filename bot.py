import asyncio
import os
import nest_asyncio
import io
import time
import pandas as pd
from telegram import Update, InputFile
from telegram.ext import CallbackContext, Application, CommandHandler, MessageHandler, filters
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
import logging

# Load environment variables
load_dotenv()
TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

nest_asyncio.apply()
running_tasks = {}
stop_event = asyncio.Event()

async def start(update: Update, context: CallbackContext) -> None:
    user = update.message.from_user
    username = user.username if user.username else user.first_name
    user_id = user.id
    
    running_tasks[user_id] = None
    await update.message.reply_text(f"Halo! {username} Kirimkan file Excel untuk memulai pengecekan.")

async def stop(update: Update, context: CallbackContext) -> None:
    user_id = update.message.from_user.id
    if user_id in running_tasks and running_tasks[user_id] is not None:
        stop_event.set()
        running_tasks[user_id] = None
        await update.message.reply_text('Proses dihentikan. Mengirimkan hasil yang tersedia...')
        
        results_file_path = 'results_file.txt'
        if os.path.exists(results_file_path):
            with open(results_file_path, 'rb') as result_file:
                await update.message.reply_document(InputFile(result_file, filename='results_file.txt'))
        else:
            await update.message.reply_text('Tidak ada hasil yang tersedia untuk dikirim.')
    else:
        await update.message.reply_text('Tidak ada proses yang sedang berjalan untuk dihentikan.')

async def process_excel(update: Update, context: CallbackContext) -> None:
    user_id = update.message.from_user.id
    if update.message.document:
        logger.info("process_excel function called")
        document = update.message.document
        file_name = document.file_name
        file = await document.get_file()
        file_path = f'user_files/{file_name}'
        
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        await file.download_to_drive(file_path)
        await update.message.reply_text('Memulai proses pemrosesan file Excel...')

        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

        def reset_page():
            driver.get("https://myim3.indosatooredoo.com/ceknomor")
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "NIK")))

        reset_page()

        try:
            df = pd.read_excel(file_path, dtype={'NIK': str, 'NO KK': str})
            logger.info("Excel file loaded successfully")
        except Exception as e:
            await update.message.reply_text(f'Gagal memuat file Excel: {str(e)}')
            driver.quit()
            return

        results_file_path = f'results_{file_name}.txt'
        with open(results_file_path, 'a') as result_file:
            for index, row in df.iterrows():
                if stop_event.is_set():
                    await update.message.reply_text('Proses dihentikan. Mengirimkan hasil yang tersedia...')
                    driver.quit()
                    if os.path.exists(results_file_path):
                        with open(results_file_path, 'rb') as processed_file:
                            await update.message.reply_document(InputFile(processed_file, filename=results_file_path))
                    else:
                        await update.message.reply_text('Tidak ada hasil yang tersedia untuk dikirim.')
                    return

                try:
                    nik = row['NIK']
                    kk = row['NO KK']
                    
                    nik_field = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.ID, "NIK")))
                    kk_field = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.ID, "KK")))

                    nik_field.clear()
                    nik_field.send_keys(nik)
                    
                    kk_field.clear()
                    kk_field.send_keys(kk)
                    
                    captcha_box = WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.CLASS_NAME, "captchaBox")))
                    captcha_box.click()
                    
                    driver.execute_script("document.getElementById('checkSubmitButton').disabled = false;")
                    submit_button = WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.ID, "checkSubmitButton")))
                    submit_button.click()
                    
                    time.sleep(2)
                    
                    try:
                        modal = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, "myModal")))
                        if modal.is_displayed():
                            result_text = f"Submission failed for NIK: {nik}, KK: {kk}. The modal appeared."
                            result_file.write(result_text + '\n')
                            df.at[index, 'Result'] = "Invalid data"
                            close_button = driver.find_element(By.XPATH, "//button[@data-dismiss='modal']")
                            close_button.click()
                            reset_page()
                            continue
                    except TimeoutException:
                        pass
                    
                    phone_numbers = []
                    phone_elements = driver.find_elements(By.XPATH, "//ul[@class='list-unstyled margin-5-top']//li")
                    for element in phone_elements:
                        phone_numbers.append(element.text.strip())
                    
                    result_text = ", ".join(phone_numbers) if phone_numbers else "No numbers found"
                    result_file.write(f"Result for NIK: {nik}, KK: {kk} -> {result_text}\n")
                    df.at[index, 'Result'] = result_text
                    reset_page()
                
                except (NoSuchElementException, TimeoutException, WebDriverException) as e:
                    error_message = f"An error occurred with NIK: {nik}, KK: {kk}. Error: {str(e)}"
                    result_file.write(error_message + '\n')
                    df.at[index, 'Result'] = f"Error: {str(e)}"
                    reset_page()
                    continue

        driver.quit()

        if os.path.exists(results_file_path):
            with open(results_file_path, 'rb') as processed_file:
                await update.message.reply_document(InputFile(processed_file, filename=f'{file_name}'))
        else:
            await update.message.reply_text('Tidak ada hasil yang tersedia untuk dikirim.')

    else:
        await update.message.reply_text('Harap kirimkan file Excel dengan format yang benar.')

async def main() -> None:
    application = Application.builder().token(TOKEN).build()
    application.add_handler(CommandHandler('start', start))
    application.add_handler(CommandHandler('stop', stop))
    application.add_handler(MessageHandler(filters.Document.ALL, process_excel))
    await application.run_polling()

if __name__ == '__main__':
    try:
        asyncio.run(main())
    except RuntimeError as e:
        if str(e) == "This event loop is already running":
            loop = asyncio.get_event_loop()
            loop.run_until_complete(main())
        else:
            raise
