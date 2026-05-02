import os
import win32com.client
from tkinter import messagebox, filedialog
from datetime import datetime

# Константы Excel для читаемости
xlDown = -4121
xlUp = -4162
xlPasteValues = -4163
xlPasteFormats = -4122

def run_reestr_process(p1, p2, version):
    if not p1:
        messagebox.showerror("Ошибка", "Выберите Документ №1 (Реестр из МОС)")
        return

    # Проверка на одинаковые файлы
    if p2 and os.path.abspath(p1) == os.path.abspath(p2):
        messagebox.showerror("Ошибка", "Документ №1 и Документ №2 не могут быть одним и тем же файлом!")
        return

    only_one = not bool(p2)
    if only_one:
        if not messagebox.askyesno("Вопрос", "Второй файл не выбран. Обработать только Документ №1?"): 
            return

    try:
        save_path = filedialog.asksaveasfilename(
            initialfile=f"Единый_реестр_ТК_{datetime.now().strftime('%d.%m.%Y')}.xlsx",
            defaultextension=".xlsx"
        )
        if not save_path: 
            return

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb1 = excel.Workbooks.Open(os.path.abspath(p1))
        ws1 = wb1.Worksheets(1)

        try:
            if not only_one:
                wb2 = excel.Workbooks.Open(os.path.abspath(p2))
                ws2 = wb2.Worksheets(1)
                
                # Находим последнюю строку в файле 1
                lrow1 = ws1.Cells(ws1.Rows.Count, 2).End(xlUp).Row
                # Находим последнюю строку в файле 2
                lrow2 = ws2.Cells(ws2.Rows.Count, 2).End(xlUp).Row
                
                if lrow2 >= 5:
                    # Копируем данные со второго листа (предполагая структуру с 5 строки по 45 колонку)
                    ws2.Range(ws2.Cells(5, 2), ws2.Cells(lrow2, 45)).Copy()
                    ws1.Cells(lrow1 + 1, 2).PasteSpecial(xlPasteValues)
                    
                    # Копируем форматирование заголовка (строка 5) на новые данные
                    last_row_new = ws1.Cells(ws1.Rows.Count, 2).End(xlUp).Row
                    if last_row_new > lrow1:
                         ws1.Rows(5).Copy()
                         ws1.Range(ws1.Cells(lrow1 + 1, 1), ws1.Cells(last_row_new, 45)).PasteSpecial(xlPasteFormats)
                
                wb2.Close(False)

            last_row = ws1.Cells(ws1.Rows.Count, 2).End(xlUp).Row
            num = 0
            all_addresses = []

            for i in range(5, last_row + 1):
                val_col2 = ws1.Cells(i, 2).Value
                if not val_col2: 
                    continue
                
                num += 1
                ws1.Cells(i, 1).Value = num
                
                # Обработка адреса (кол 8) и комментария (кол 26)
                h_raw = ws1.Cells(i, 8).Value
                z_raw = ws1.Cells(i, 26).Value
                
                h = str(h_raw or "").strip()
                z = str(z_raw or "")
                
                if "Россия" in h: 
                    h = h.split("Россия")[-1].strip(", ")
                
                if "домофон" in h.lower():
                    idx = h.lower().find("домофон")
                    z = f"{z} {h[idx:]}".strip()
                    h = h[:idx].strip(", ")
                
                ws1.Cells(i, 8).Value = h
                ws1.Cells(i, 26).Value = z
                all_addresses.append(h.lower())

                # Расчет типа доставки
                try:
                    w_val = ws1.Cells(i, 12).Value
                    m1_val = ws1.Cells(i, 28).Value  # AB
                    m2_val = ws1.Cells(i, 29).Value  # AC
                    m3_val = ws1.Cells(i, 30).Value  # AD
                    
                    w = float(w_val) if w_val is not None else 0.0
                    m_vals = [float(x) if x is not None else 0.0 for x in [m1_val, m2_val, m3_val]]
                    m = max(m_vals)
                    
                    t_cell = ws1.Cells(i, 9)  # Столбец I
                    
                    # НОВОЕ ПРАВИЛО: если любой габарит > 1.801, то всегда "Стандартная с выгрузкой"
                    if m > 1.801:
                        t_cell.Value = "Стандартная с выгрузкой"
                    elif w > 30: 
                        t_cell.Value = "Стандартная с выгрузкой"
                    elif 10.01 <= w <= 30: 
                        t_cell.Value = "Малогабаритная" if m <= 1.8 else "Стандартная с выгрузкой"
                    else: 
                        t_cell.Value = "Курьерская" if m <= 1.8 else "Малогабаритная"
                except ValueError:
                    pass

                # Обработка времени
                ao_val = ws1.Cells(i, 41).Value
                ao = str(ao_val or "").lower()
                
                if "маркетплейс" in ao or "маркетплэйс" in ao: 
                    ws1.Cells(i, 3).Value = "08:00-23:00"
                if "заказ в пвз" in z.lower(): 
                    ws1.Cells(i, 3).Value = "10:00 - 21:00"

            # Подсветка дублей адресов
            for i in range(5, last_row + 1):
                addr = str(ws1.Cells(i, 8).Value or "").strip().lower()
                if addr and all_addresses.count(addr) > 1:
                    ws1.Cells(i, 8).Interior.ColorIndex = 6

            wb1.SaveAs(os.path.abspath(save_path))
            messagebox.showinfo("Готово", "Реестр успешно сформирован!")
            
        finally:
            wb1.Close(False)
            excel.Quit()

    except Exception as e:
        messagebox.showerror("Критическая ошибка", f"Произошла ошибка при обработке:\n{str(e)}")
        if 'excel' in locals():
            try: excel.Quit()
            except: pass
