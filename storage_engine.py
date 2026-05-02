import json, os, webbrowser
from tkinter import filedialog

CONFIG_FILE = "config.txt"
DB_NAME = "storage_data_v5.json"

def get_save_directory():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            p = f.read().strip()
            if os.path.exists(p): return p
    
    folder = filedialog.askdirectory(title="Выберите папку для базы ячеек")
    if folder:
        path = os.path.join(folder, "данные ячеек")
        if not os.path.exists(path): os.makedirs(path)
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f: f.write(path)
        return path
    return None

def load_data(folder):
    path = os.path.join(folder, DB_NAME)
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f: 
                return json.load(f)
        except json.JSONDecodeError:
            pass
            
    return {
        "sections": [
            {"name": "Основное хранение", "cells": [{"cell_id": "1-1", "order_id": "", "note": ""} for _ in range(12)]},
            {"name": "Уличное хранение", "cells": [{"cell_id": "У-1", "order_id": "", "note": ""} for _ in range(8)]},
            {"name": "Прочее", "cells": [{"cell_id": "Р-1", "order_id": "", "note": ""} for _ in range(4)]}
        ]
    }

def save_data(data, folder):
    if not folder: return
    try:
        with open(os.path.join(folder, DB_NAME), 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Ошибка сохранения: {e}")

def open_order_link(oid):
    if oid:
        # ИСПРАВЛЕНО: убран лишний пробел в URL
        url = f"https://mos-terminal.api.lenta.com/order/{oid}"
        webbrowser.open(url)
