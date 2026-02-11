def transform_price(input_path, output_path):
    wb = load_workbook(input_path)
    ws = wb.active
    
    max_row = ws.max_row
    
    # Сохраняем высоты строк
    original_row_heights = {}
    for row in range(1, max_row + 1):
        if ws.row_dimensions[row].height:
            original_row_heights[row] = ws.row_dimensions[row].height
    
    # Сохраняем цены
    price_data = {}
    for row in range(1, max_row + 1):
        prices = []
        m_val = ws.cell(row=row, column=13).value
        o_val = ws.cell(row=row, column=15).value
        q_val = ws.cell(row=row, column=17).value
        s_val = ws.cell(row=row, column=19).value
        
        if m_val and m_val != 0 and not isinstance(m_val, str):
            try:
                prices.append(f"250тр: {int(m_val)}₽")
            except:
                pass
        if o_val and o_val != 0 and not isinstance(o_val, str):
            try:
                prices.append(f"100т: {int(o_val)}₽")
            except:
                pass
        if q_val and q_val != 0 and not isinstance(q_val, str):
            try:
                prices.append(f"50т: {int(q_val)}₽")
            except:
                pass
        if s_val and s_val != 0 and not isinstance(s_val, str):
            try:
                prices.append(f"25тр: {int(s_val)}₽")
            except:
                pass
        
        if prices:
            price_data[row] = "\n".join(prices)
    
    # Очищаем путь к фото
    for row in range(1, max_row + 1):
        cell = ws.cell(row=row, column=3)
        if cell.value and "D:\\" in str(cell.value):
            cell.value = None
    
    # Разъединяем ячейки
    merged_ranges = list(ws.merged_cells.ranges)
    for merged_range in merged_ranges:
        ws.unmerge_cells(str(merged_range))
    
    # Извлекаем и уменьшаем изображения
    new_images = []
    for img in ws._images:
        if hasattr(img.anchor, '_from'):
            anchor = img.anchor._from
            if anchor.col == 2:
                row_num = anchor.row
                try:
                    img_data = img._data()
                    pil_img = Image.open(io.BytesIO(img_data))
                    pil_img.thumbnail((70, 70), Image.LANCZOS)
                    
                    buf = io.BytesIO()
                    pil_img.save(buf, format='PNG')
                    buf.seek(0)
                    
                    new_img = XLImage(buf)
                    new_img.anchor = f'A{row_num + 1}'
                    new_images.append((row_num, new_img))
                except:
                    pass
    
    ws._images = []
    
    # Удаляем столбцы
    cols_to_delete = [29, 28, 27, 26, 25, 24, 21, 20, 19, 18, 17, 16, 15, 14, 13, 10, 9, 8, 4, 3, 2]
    for col in sorted(cols_to_delete, reverse=True):
        ws.delete_cols(col)
    
    # Добавляем цены
    for row, combined_price in price_data.items():
        cell = ws.cell(row=row, column=9)
        if not isinstance(cell, MergedCell):
            cell.value = combined_price
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            cell.font = Font(size=12)
    
    # Ширины столбцов
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 8
    ws.column_dimensions['F'].width = 11
    ws.column_dimensions['G'].width = 8
    ws.column_dimensions['H'].width = 12
    ws.column_dimensions['I'].width = 18
    
    # Удаляем первые 6 строк
    ws.delete_rows(1, 6)
    
    # Добавляем изображения (со сдвигом +1 для заголовка)
    for row_num, img in new_images:
        new_row = row_num - 6 + 2  # +2: один для сдвига, один для заголовка
        if new_row > 1:
            img.anchor = f'A{new_row}'
            ws.add_image(img)
    
    # Применяем высоты строк (со сдвигом +1 для заголовка)
    for old_row, height in original_row_heights.items():
        new_row = old_row - 6 + 1
        if new_row > 1:
            ws.row_dimensions[new_row].height = height
    
    # Вставляем строку заголовков
    ws.insert_rows(1)
    headers = ['Фото', 'Наименование', 'Код', 'Страна', 'Штук в блоке', 'Годен до', 'Вес', 'Объём', 'Цены']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col)
        cell.value = header
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    
    ws.row_dimensions[1].height = 20
    
    wb.save(output_path)
    return {"success": True, "rows": ws.max_row, "images": len(ws._images)}
