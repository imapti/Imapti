import os
import sys
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

# Настройка путей
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path: 
    sys.path.append(current_dir)

try:
    import print_engine
    import reestr_engine
    import storage_engine
except ImportError as e:
    print(f"КРИТИЧЕСКАЯ ОШИБКА ИМПОРТА: {e}")
    print("Убедитесь, что файлы print_engine.py, reestr_engine.py, storage_engine.py находятся в той же папке.")
    input("Нажмите Enter для выхода...")
    sys.exit(1)

APP_VERSION = "7.11 Stable"
BG_COLOR = "#ffe5b4" 

class LogisticsApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Логистический Ассистент v{APP_VERSION}")
        
        # Исправлено: безопасное развертывание окна
        try:
            self.root.state('zoomed')
        except:
            self.root.geometry("1920x1080")
            self.root.attributes('-zoomed', True)
        
        self.db_folder = None
        self.data = None
        self.current_module = "none"
        self.last_width = 0
        self.cell_widgets = {}

        self.duplex_var = tk.IntVar(value=0)
        self.fit_var = tk.IntVar(value=0)

        # Sidebar
        self.sidebar = tk.Frame(self.root, bg="#2c3e50", width=220)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        tk.Label(self.sidebar, text="МЕНЮ", font=("Arial", 18, "bold"), bg="#2c3e50", fg="#ecf0f1", pady=40).pack()
        self.btn_print = self.create_nav_btn("МАССОВАЯ ПЕЧАТЬ", self.show_print_module)
        self.btn_reestr = self.create_nav_btn("РЕЕСТР ДЛЯ ТК", self.show_reestr_module)
        self.btn_storage = self.create_nav_btn("ЯЧЕЕЧНОЕ ХРАНЕНИЕ", self.show_storage_module)

        self.main_area = tk.Frame(self.root, bg=BG_COLOR)
        self.main_area.pack(side="right", expand=True, fill="both")

        # Универсальный бинд вставки
        self.root.bind_all("<Control-KeyPress>", self.universal_paste_handler)
        self.root.bind_all("<Button-1>", self.global_click_handler)
        self.root.bind("<Configure>", self.on_resize)
        
        # Загрузка начального модуля с обработкой ошибок
        try:
            self.show_print_module()
        except Exception as e:
            messagebox.showerror("Ошибка запуска", f"Не удалось загрузить модуль печати:\n{str(e)}")

    def universal_paste_handler(self, event):
        # Проверяем, что нажат именно Ctrl+V (или русская М)
        # event.state & 0x0004 проверяет флаг Ctrl
        if event.state & 0x0004:
            if event.keysym.lower() in ('v', 'м', 'cyrillic_em'):
                widget = self.root.focus_get()
                if isinstance(widget, (tk.Entry, tk.Text)):
                    try:
                        # Выполняем вставку через буфер обмена Windows
                        clipboard = self.root.clipboard_get()
                        if clipboard:
                            widget.insert(tk.INSERT, clipboard)
                        return "break"  # Прерываем стандартную обработку, чтобы не было задвоения
                    except: 
                        pass
        return None  # Возвращаем None, чтобы другие обработчики сработали для остальных клавиш

    def create_nav_btn(self, text, command):
        btn = tk.Button(self.sidebar, text=text, command=command, bg="#34495e", fg="white", 
                        font=("Arial", 11, "bold"), relief="flat", pady=15, cursor="hand2")
        btn.pack(fill="x", padx=15, pady=5)
        return btn

    def global_click_handler(self, event):
        if self.current_module == "storage": self.reset_highlights()

    def add_magic_footer(self, parent, description=""):
        if description:
            desc_lbl = tk.Label(parent, text=description, font=("Arial", 12), bg=BG_COLOR, 
                                fg="#2c3e50", justify="left", wraplength=1100, anchor="w")
            desc_lbl.pack(pady=(25, 0), padx=60, fill="x")
        tk.Label(parent, text="Создано магией команды логистики", 
                 font=("Arial", 10, "italic"), bg=BG_COLOR, fg="#7f8c8d").pack(side="bottom", pady=15)

    def on_resize(self, e):
        if self.current_module == "storage":
            w = self.root.winfo_width()
            if abs(w - self.last_width) > 30:
                self.last_width = w
                self.render_storage()

    def clear_main(self):
        self.current_module = "transition"
        for w in self.main_area.winfo_children(): w.destroy()

    def update_colors(self, active):
        for b in [self.btn_print, self.btn_reestr, self.btn_storage]: b.config(bg="#34495e")
        active.config(bg="#1abc9c")

    def quick_paste(self, widget):
        try:
            widget.event_generate("<<Paste>>")
        except: pass

    # --- МОДУЛЬ 1: ПЕЧАТЬ ---
    def show_print_module(self):
        self.clear_main(); self.current_module = "print"; self.update_colors(self.btn_print)
        tk.Label(self.main_area, text="Массовая печать документов", font=("Arial", 26, "bold"), bg=BG_COLOR).pack(pady=40)
        p_frame = tk.Frame(self.main_area, bg=BG_COLOR); p_frame.pack(pady=10, fill="x")
        sub_row = tk.Frame(p_frame, bg=BG_COLOR); sub_row.pack(anchor="center")
        self.folder_ent = tk.Entry(sub_row, font=("Arial", 16), width=45, bd=1, relief="solid"); self.folder_ent.pack(side="left", ipady=8)
        tk.Button(sub_row, text="ОБЗОР ПАПКИ", bg="#34495e", fg="white", font=("Arial", 10, "bold"), padx=15,
                  command=lambda: self.folder_ent.insert(0, filedialog.askdirectory())).pack(side="left", padx=10, ipady=5)
        opt_frame = tk.Frame(self.main_area, bg=BG_COLOR); opt_frame.pack(pady=20)
        tk.Checkbutton(opt_frame, text="Двусторонняя печать", variable=self.duplex_var, font=("Arial", 12, "bold"), bg=BG_COLOR, fg="#c0392b").pack(side="left", padx=20)
        tk.Checkbutton(opt_frame, text="Вписать на одну страницу", variable=self.fit_var, font=("Arial", 12, "bold"), bg=BG_COLOR, fg="#2980b9").pack(side="left", padx=20)
        tk.Label(self.main_area, text="Количество копий:", font=("Arial", 16, "bold"), bg=BG_COLOR).pack(pady=(20, 5))
        c_frame = tk.Frame(self.main_area, bg=BG_COLOR); c_frame.pack()
        tk.Button(c_frame, text="-", font=("Arial", 20, "bold"), width=4, bg="#bdc3c7", command=self.dec_copies).pack(side="left")
        self.copies_ent = tk.Entry(c_frame, font=("Arial", 22, "bold"), width=6, justify="center"); self.copies_ent.insert(0, "1"); self.copies_ent.pack(side="left", padx=15)
        tk.Button(c_frame, text="+", font=("Arial", 20, "bold"), width=4, bg="#bdc3c7", command=self.inc_copies).pack(side="left")
        tk.Button(self.main_area, text="ЗАПУСТИТЬ ПЕЧАТЬ", bg="#27ae60", fg="white", font=("Arial", 20, "bold"), padx=80, pady=25, 
                  command=self.start_mass_print).pack(pady=40)
        self.add_magic_footer(self.main_area, "Принцип работы: программа печатает все Excel-файлы из выбранной папки.")

    def inc_copies(self):
        try:
            v = int(self.copies_ent.get() or 0); self.copies_ent.delete(0, tk.END); self.copies_ent.insert(0, str(v+1))
        except: pass
    def dec_copies(self):
        try:
            v = int(self.copies_ent.get() or 0); 
            if v > 1: self.copies_ent.delete(0, tk.END); self.copies_ent.insert(0, str(v-1))
        except: pass

    def start_mass_print(self):
        try:
            print_engine.run_mass_print(self.folder_ent.get(), int(self.copies_ent.get() or 1), bool(self.duplex_var.get()), bool(self.fit_var.get()))
            messagebox.showinfo("Готово", "Печать завершена!")
        except Exception as e: messagebox.showerror("Ошибка", str(e))

    # --- МОДУЛЬ 2: РЕЕСТР ---
    def show_reestr_module(self):
        self.clear_main(); self.current_module = "reestr"; self.update_colors(self.btn_reestr)
        tk.Label(self.main_area, text="Реестр для ТК", font=("Arial", 26, "bold"), bg=BG_COLOR).pack(pady=30)
        r_frame = tk.Frame(self.main_area, bg=BG_COLOR); r_frame.pack(pady=10, fill="x")
        def create_row(label_text):
            row = tk.Frame(r_frame, bg=BG_COLOR); row.pack(pady=12, anchor="center")
            tk.Label(row, text=label_text, font=("Arial", 14, "bold"), bg=BG_COLOR, width=35, anchor="e").pack(side="left", padx=5)
            ent = tk.Entry(row, font=("Arial", 14), width=35); ent.pack(side="left", ipady=6)
            # Исправлено: очищаем поле перед вставкой нового пути
            tk.Button(row, text="ОБЗОР", bg="#34495e", fg="white", width=12, font=("Arial", 10, "bold"), 
                      command=lambda e=ent: (e.delete(0, tk.END), e.insert(0, filedialog.askopenfilename()))).pack(side="left", padx=10, ipady=3)
            return ent
        self.e1 = create_row("Основной реестр (Реестр из МОС):"); self.e2 = create_row("Второй реестр (1-я форма):")
        tk.Button(self.main_area, text="ЗАПУСТИТЬ ОБРАБОТКУ", bg="#2980b9", fg="white", font=("Arial", 20, "bold"), padx=70, pady=25, 
                  command=self.start_reestr_process).pack(pady=20)
        desc = "Принцип работы: сопоставление данных двух файлов для выявления расхождений."
        self.add_magic_footer(self.main_area, desc)

    def start_reestr_process(self):
        path1 = self.e1.get().strip()
        path2 = self.e2.get().strip()
        
        if not path1:
            messagebox.showerror("Ошибка", "Выберите Документ №1 (Реестр из МОС)")
            return
            
        if not os.path.exists(path1):
            messagebox.showerror("Ошибка", f"Файл №1 не найден:\n{path1}")
            return
            
        if path2 and not os.path.exists(path2):
            messagebox.showerror("Ошибка", f"Файл №2 не найден:\n{path2}")
            return
            
        if path1 == path2:
            messagebox.showerror("Ошибка", "Документ №1 и Документ №2 не могут быть одинаковыми файлами!")
            return
            
        try:
            reestr_engine.run_reestr_process(path1, path2, APP_VERSION)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # --- МОДУЛЬ 3: ХРАНЕНИЕ ---
    def show_storage_module(self):
        if not self.db_folder:
            self.db_folder = storage_engine.get_save_directory()
            if not self.db_folder: return
            self.data = storage_engine.load_data(self.db_folder)
        self.clear_main(); self.current_module = "storage"; self.update_colors(self.btn_storage)
        top = tk.Frame(self.main_area, bg=BG_COLOR, pady=10); top.pack(fill="x", anchor="nw")
        self.search_ent = tk.Entry(top, font=("Arial", 14), width=30, bd=1, relief="solid"); self.search_ent.pack(side="left", padx=(20, 0), ipady=5)
        tk.Button(top, text="ВСТАВИТЬ", bg="#f39c12", fg="white", font=("Arial", 10, "bold"), command=lambda: self.quick_paste(self.search_ent)).pack(side="left", padx=5, ipady=3)
        tk.Button(top, text="ПОИСК ПО НОМЕРУ ЗАКАЗА", bg="#3498db", fg="white", font=("Arial", 10, "bold"), command=self.search_order).pack(side="left", padx=5, ipady=3)
        container = tk.Frame(self.main_area, bg=BG_COLOR); container.pack(expand=True, fill="both")
        self.canvas = tk.Canvas(container, bg=BG_COLOR, highlightthickness=0); self.scroll_frame = tk.Frame(self.canvas, bg=BG_COLOR)
        sb = tk.Scrollbar(container, command=self.canvas.yview); self.canvas.configure(yscrollcommand=sb.set)
        self.canvas.pack(side="left", expand=True, fill="both"); sb.pack(side="right", fill="y")
        self.canvas_frame_id = self.canvas.create_window((0,0), window=self.scroll_frame, anchor="nw")
        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.canvas_frame_id, width=e.width))
        self.root.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        self.render_storage(); self.add_magic_footer(self.main_area)

    def reset_highlights(self):
        for w in self.cell_widgets.values(): 
            try: w.config(highlightthickness=0)
            except: pass

    # --- ИСПРАВЛЕННЫЙ ПОИСК ---
    def search_order(self):
        self.reset_highlights()
        q = self.search_ent.get().strip()
        if not q: return
        
        found_info = []
        for s_idx, section in enumerate(self.data["sections"]):
            for c_idx, cell in enumerate(section["cells"]):
                if cell["order_id"] == q:
                    key = f"{s_idx}_{c_idx}"
                    if key in self.cell_widgets:
                        self.cell_widgets[key].config(highlightbackground="red", highlightthickness=4)
                    found_info.append(f"{section['name']}: {cell['cell_id']}")
        
        if found_info:
            top = tk.Toplevel(self.root); top.title("Найдено"); top.geometry("600x400")
            tk.Label(top, text="ЗАКАЗ НАЙДЕН!", font=("Arial", 20, "bold"), fg="red").pack(pady=20)
            tk.Label(top, text="\n".join(found_info), font=("Arial", 16, "bold"), justify="center").pack(pady=10)
            tk.Button(top, text="ОК", font=("Arial", 12, "bold"), width=15, command=top.destroy).pack(side="bottom", pady=20)
        else:
            messagebox.showwarning("Поиск", "Заказ не найден!")

    def render_storage(self):
        if self.current_module != "storage" or not self.canvas: return
        for w in self.scroll_frame.winfo_children(): w.destroy()
        self.cell_widgets = {} 
        cols = max(1, (self.main_area.winfo_width() - 60) // 230) 
        
        # Функция для открытия ссылки на заказ
        def on_order_click(oid):
            if oid:
                storage_engine.open_order_link(oid)
        
        for s_idx, section in enumerate(self.data["sections"]):
            tk.Frame(self.scroll_frame, bg="black", height=2).pack(fill="x", pady=(10,0))
            head = tk.Frame(self.scroll_frame, bg="#34495e", pady=3); head.pack(fill="x")
            lbl_head = tk.Label(head, text=f" {section['name']}", font=("Arial", 12, "bold"), fg="white", bg="#34495e", cursor="hand2")
            lbl_head.pack(side="left", padx=10); lbl_head.bind("<Button-1>", lambda e, si=s_idx: self.rename_section(si))
            tk.Button(head, text="[ + ] ДОБАВИТЬ ЯЧЕЙКУ", bg="#27ae60", fg="white", font=("Arial", 8, "bold"), command=lambda si=s_idx: self.add_cell(si)).pack(side="right", padx=5)
            tk.Button(head, text="[ - ] УДАЛИТЬ ЯЧЕЙКУ", bg="#c0392b", fg="white", font=("Arial", 8, "bold"), command=lambda si=s_idx: self.confirm_remove_cell(si)).pack(side="right", padx=5)
            grid_f = tk.Frame(self.scroll_frame, bg=BG_COLOR); grid_f.pack(fill="x", padx=10, pady=10)
            for c_idx, cell in enumerate(section["cells"]):
                box = tk.Frame(grid_f, bg="#2c3e50", padx=1, pady=1, highlightthickness=0); box.grid(row=c_idx//cols, column=c_idx%cols, padx=12, pady=12)
                self.cell_widgets[f"{s_idx}_{c_idx}"] = box
                hdr = tk.Frame(box, bg="#ecf0f1"); hdr.pack(fill="x")
                tk.Button(hdr, text="<", font=("Arial", 8), width=3, command=lambda si=s_idx, ci=c_idx: self.move_cell(si, ci, -1)).pack(side="left")
                tk.Button(hdr, text=f"Ячейка №{cell['cell_id']}", font=("Arial", 9, "bold"), bg="#ecf0f1", relief="flat", command=lambda si=s_idx, ci=c_idx: self.rename_cell(si, ci)).pack(side="left", expand=True, fill="x")
                tk.Button(hdr, text=">", font=("Arial", 8), width=3, command=lambda si=s_idx, ci=c_idx: self.move_cell(si, ci, 1)).pack(side="right")
                mid = tk.Frame(box, bg="#2c3e50"); mid.pack()
                
                # Создаем кнопку заказа с возможностью клика для открытия ссылки
                if cell["order_id"]:
                    btn = tk.Button(mid, text=cell["order_id"], bg="#3498db", fg="white", 
                                    font=("Arial", 12, "bold"), width=15, height=2, relief="flat",
                                    command=lambda oid=cell["order_id"]: on_order_click(oid),
                                    cursor="hand2")
                else:
                    btn = tk.Button(mid, text="ПУСТО", bg="#2ecc71", fg="white", 
                                    font=("Arial", 12, "bold"), width=15, height=2, relief="flat")
                btn.pack(side="left")
                
                tk.Button(mid, text="i", bg="#f1c40f" if cell.get("note") else "#bdc3c7", font=("Arial", 10, "bold"), width=3, height=2, command=lambda si=s_idx, ci=c_idx: self.edit_note(si, ci)).pack(side="right")
                tk.Button(box, text="ОЧИСТИТЬ" if cell["order_id"] else "ПРИСВОИТЬ", bg="#e74c3c" if cell["order_id"] else "#f39c12", fg="white", font=("Arial", 9), command=lambda si=s_idx, ci=c_idx, f=bool(cell["order_id"]): self.confirm_clear_cell(si, ci) if f else self.assign_cell(si, ci)).pack(fill="x")

    def edit_note(self, si, ci):
        item = self.data["sections"][si]["cells"][ci]
        win = tk.Toplevel(self.root); win.title("Заметка"); win.geometry("550x500"); win.grab_set()
        txt = tk.Text(win, height=18, width=60, font=("Arial", 10)); txt.insert("1.0", item.get("note", "")); txt.pack(pady=10, padx=10)
        btn_f = tk.Frame(win); btn_f.pack(pady=10)
        tk.Button(btn_f, text="ВСТАВИТЬ", bg="#f39c12", fg="white", font=("Arial", 10, "bold"), command=lambda: self.quick_paste(txt)).pack(side="left", padx=10)
        tk.Button(btn_f, text="СОХРАНИТЬ", bg="#27ae60", fg="white", font=("Arial", 10, "bold"), command=lambda: self.save_note(si, ci, txt.get("1.0", tk.END).strip(), win)).pack(side="left", padx=10)

    def save_note(self, si, ci, val, win):
        self.data["sections"][si]["cells"][ci]["note"] = val; self.save_refresh(); win.destroy()

    def assign_cell(self, si, ci):
        win = tk.Toplevel(self.root); win.title("Номер заказа"); win.geometry("450x180"); win.grab_set()
        e = tk.Entry(win, font=("Arial", 14), width=22); e.pack(pady=20)
        tk.Button(win, text="ВСТАВИТЬ", bg="#f39c12", fg="white", font=("Arial", 9, "bold"), command=lambda: self.quick_paste(e)).pack()
        tk.Button(win, text="ПРИСВОИТЬ", bg="#27ae60", fg="white", font=("Arial", 11, "bold"), command=lambda: [self.data["sections"][si]["cells"][ci].update({"order_id": e.get().strip()}), self.save_refresh(), win.destroy()]).pack(pady=15)

    def confirm_remove_cell(self, si):
        if self.data["sections"][si]["cells"] and messagebox.askyesno("Удаление", "Удалить ячейку?"):
            self.data["sections"][si]["cells"].pop(); self.save_refresh()
    def confirm_clear_cell(self, si, ci):
        if messagebox.askyesno("Очистка", "Очистить ячейку?"):
            self.data["sections"][si]["cells"][ci]["order_id"] = ""; self.save_refresh()
    def move_cell(self, si, ci, s):
        t = ci + s
        if 0 <= t < len(self.data["sections"][si]["cells"]):
            self.data["sections"][si]["cells"][ci], self.data["sections"][si]["cells"][t] = self.data["sections"][si]["cells"][t], self.data["sections"][si]["cells"][ci]; self.save_refresh()
    def rename_section(self, si):
        r = simpledialog.askstring("Блок", "Имя:", initialvalue=self.data["sections"][si]["name"])
        if r: self.data["sections"][si]["name"] = r; self.save_refresh()
    def rename_cell(self, si, ci):
        r = simpledialog.askstring("Ячейка", "Имя:", initialvalue=self.data["sections"][si]["cells"][ci]["cell_id"])
        if r: self.data["sections"][si]["cells"][ci]["cell_id"] = r.strip(); self.save_refresh()
    def add_cell(self, si): self.data["sections"][si]["cells"].append({"cell_id": "?", "order_id": "", "note": ""}); self.save_refresh()
    def save_refresh(self): storage_engine.save_data(self.data, self.db_folder); self.render_storage()

if __name__ == "__main__":
    root = tk.Tk(); app = LogisticsApp(root); root.mainloop()
