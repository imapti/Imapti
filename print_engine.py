import win32com.client
import os
import time

def run_mass_print(folder_path, copies=1, is_duplex=False, is_fit=False):
    """Движок для массовой печати файлов Excel (.xlsx, .xls, .xlsm)"""
    if not folder_path or not os.path.exists(folder_path):
        return

    folder_path = os.path.abspath(folder_path)
    files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.xlsx', '.xls', '.xlsm'))]
    
    if not files:
        raise Exception("В выбранной папке не найдено файлов Excel!")

    excel = None
    try:
        # Запускаем процесс Excel
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        for file_name in files:
            file_path = os.path.join(folder_path, file_name)
            wb = excel.Workbooks.Open(file_path, ReadOnly=True)
            
            try:
                for ws in wb.Sheets:
                    # Настройка "Вписать на одну страницу"
                    if is_fit:
                        ws.PageSetup.Zoom = False
                        ws.PageSetup.FitToPagesWide = 1
                        ws.PageSetup.FitToPagesTall = 1
                    
                    # Печать конкретного листа
                    # Примечание: Excel берет Duplex из настроек принтера по умолчанию
                    ws.PrintOut(Copies=copies)
                
                time.sleep(0.3) # Короткая пауза для очереди
            finally:
                wb.Close(False)
                
    except Exception as e:
        raise Exception(f"Ошибка печати Excel: {str(e)}")
    finally:
        if excel:
            excel.Quit()
