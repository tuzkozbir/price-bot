import os
import logging
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from copy import copy

BOT_TOKEN = os.environ.get("BOT_TOKEN", "")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def transform_price(input_path, output_path):
    """Преобразует прайс-лист: удаляет лишнее, объединяет цены"""
    
    wb = load_workbook(input_path)
    ws = wb.active
    
    max_row = ws.max_row
    max_col = ws.max_column
    
    # Собираем данные о ценах ДО удаления столбцов
    # Цены в столбцах: M=13, O=15, Q=17, S=19
    price_data = {}
    for row in range(1, max_row + 1):
        prices = []
        m_val = ws.cell(row=row, column=13).value
        o_val = ws.cell(row=row, column=15).value
        q_val = ws.cell(row=row, column=17).value
        s_val = ws.cell(row=row, column=19).value
        
        if m_val and m_val != 0 and not isinstance(m_val, str):
            prices.append(f"250тр: {int(m_val)}₽")
        elif isinstance(m_val, str) and "250" in str(m_val):
            prices.append(str(m_val))
            
        if o_val and o_val != 0 and not isinstance(o_val, str):
            prices.append(f"100т: {int(o_val)}₽")
        if q_val and q_val != 0 and not isinstance(q_val, str):
            prices.append(f"50т: {int(q_val)}₽")
        if s_val and s_val != 0 and not isinstance(s_val, str):
            prices.append(f"25тр: {int(s_val)}₽")
        
        if prices:
            price_data[row] = "\n".join(prices)
    
    # Удаляем столбцы СПРАВА НАЛЕВО (чтобы индексы не сбивались)
    # U=21 (Итого), T=20, S=19, R=18, Q=17, P=16, O=15, N=14 - лишние столбцы цен и пустые
    # I=9 (Ваш заказ), H=8 (Наличие)
    cols_to_delete = [21, 20, 19, 18, 17, 16, 15, 14, 9, 8]
    
    for col in sorted(cols_to_delete, reverse=True):
        ws.delete_cols(col)
    
    # Теперь столбец M (13) стал столбцом с ценами, записываем объединённые цены
    # После удаления H(8) и I(9), столбец M сдвинулся на 2 влево = столбец 11
    price_col = 11
    
    for row, combined_price in price_data.items():
        cell = ws.cell(row=row, column=price_col)
        cell.value = combined_price
        cell.alignment = Alignment(wrap_text=True, vertical='top')
    
    # Устанавливаем ширину столбца с ценами
    ws.column_dimensions['K'].width = 20
    
    # Удаляем первые 2 строки (контакты)
    ws.delete_rows(1, 2)
    
    wb.save(output_path)
    
    return {
        "success": True,
        "rows": ws.max_row
    }


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Привет! Отправь мне Excel файл с прайсом.\n\n"
        "Я удалю лишние столбцы и объединю цены в один столбец."
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    
    if not document.file_name.endswith(('.xlsx', '.xls')):
        await update.message.reply_text("❌ Отправь Excel файл (.xlsx)")
        return
    
    await update.message.reply_text("⏳ Обрабатываю...")
    
    try:
        file = await context.bot.get_file(document.file_id)
        input_path = f"/tmp/input_{document.file_name}"
        output_path = f"/tmp/telegram_{document.file_name}"
        
        await file.download_to_drive(input_path)
        
        result = transform_price(input_path, output_path)
        
        if result["success"]:
            await update.message.reply_document(
                document=open(output_path, 'rb'),
                filename=f"telegram_{document.file_name}",
                caption=f"✅ Готово! Строк: {result['rows']}"
            )
        
        # Удаляем временные файлы
        if os.path.exists(input_path):
            os.remove(input_path)
        if os.path.exists(output_path):
            os.remove(output_path)
            
    except Exception as e:
        logger.error(f"Error: {e}")
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")


def main():
    if not BOT_TOKEN:
        print("❌ BOT_TOKEN не установлен!")
        return
    
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    
    print("🤖 Бот запущен!")
    app.run_polling()


if __name__ == '__main__':
    main()
