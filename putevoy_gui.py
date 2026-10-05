"""
Графический интерфейс генератора путевых листов (Windows 10/11).
Перетащите путевой лист за прошлый месяц в окно — или список магазинов
(.xlsx), чтобы обновить справочник.
"""
import os
import shutil
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except Exception:          # без библиотеки — работает выбор файла по клику
    HAS_DND = False

import putevoy_core as C

APP_TITLE = "Путевые листы"
FONT = ("Segoe UI", 10)
MUTED = "#6b7280"
ERROR = "#b91c1c"
OK = "#15803d"
DROP_BG = "#f3f4f6"
DROP_BG_LOADED = "#ecfdf5"


# ------------------------------------------------------------ пути
def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def shops_path():
    """shops.json лежит рядом с программой; при первом запуске копируется
    из встроенной в exe копии."""
    path = os.path.join(app_dir(), "shops.json")
    if not os.path.exists(path):
        bundled = os.path.join(getattr(sys, "_MEIPASS", app_dir()), "shops.json")
        if os.path.exists(bundled) and os.path.abspath(bundled) != os.path.abspath(path):
            try:
                shutil.copyfile(bundled, path)
            except OSError:
                return bundled
    return path


# ------------------------------------------------------------ окно
class App:
    def __init__(self, root):
        self.root = root
        self.db_path = shops_path()
        self.db = C.load_db(self.db_path)
        self.info = None
        self.changes = {}
        self.out_path = None

        root.title(APP_TITLE)
        root.minsize(600, 640)
        root.option_add("*Font", FONT)
        style = ttk.Style(root)
        style.configure("Big.TButton", font=(FONT[0], 11, "bold"), padding=(16, 8))
        style.configure("Link.TButton", foreground="#1d4ed8", padding=0, relief="flat")

        main = ttk.Frame(root, padding=16)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)

        # --- зона перетаскивания
        self.drop = tk.Label(main, bg=DROP_BG, fg="#374151", relief="groove", bd=2,
                             height=4, cursor="hand2", justify="center", font=(FONT[0], 11))
        self.drop.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 14))
        self.drop.bind("<Button-1>", lambda e: self.browse())
        self._set_drop_text()
        if HAS_DND:
            for w in (self.drop, root):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self.on_drop)

        # --- поля
        self.office_var = tk.StringVar()
        self.liters_var = tk.StringVar()
        self.norm_var = tk.StringVar()
        self.rest_var = tk.StringVar(value="—")
        self.weekends_var = tk.BooleanVar(value=False)

        ttk.Label(main, text="Адрес выезда").grid(row=1, column=0, sticky="w", padx=(0, 12))
        self.office_entry = ttk.Entry(main, textvariable=self.office_var)
        self.office_entry.grid(row=1, column=1, sticky="ew")
        self.office_status = ttk.Label(main, text="", foreground=MUTED, font=(FONT[0], 9))
        self.office_status.grid(row=2, column=1, sticky="w", pady=(2, 10))
        self.office_var.trace_add("write", lambda *a: self.update_office_status())

        ttk.Label(main, text="Выдано топлива, л").grid(row=3, column=0, sticky="w", pady=4)
        self.liters_entry = ttk.Entry(main, textvariable=self.liters_var, width=12)
        self.liters_entry.grid(row=3, column=1, sticky="w", pady=4)
        self.liters_entry.bind("<Return>", lambda e: self.generate())

        ttk.Label(main, text="Норма расхода, л/100 км").grid(row=4, column=0, sticky="w", pady=4)
        self.norm_entry = ttk.Entry(main, textvariable=self.norm_var, width=12)
        self.norm_entry.grid(row=4, column=1, sticky="w", pady=4)

        ttk.Label(main, text="Остаток на начало, л").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Label(main, textvariable=self.rest_var).grid(row=5, column=1, sticky="w", pady=4)

        self.weekends_check = ttk.Checkbutton(main, text="Разрешить поездки в выходные",
                                              variable=self.weekends_var)
        self.weekends_check.grid(row=6, column=1, sticky="w", pady=(6, 2))

        consts = ttk.Frame(main)
        consts.grid(row=7, column=1, sticky="w", pady=(2, 14))
        self.consts_btn = ttk.Button(consts, text="Изменить постоянные данные (водитель, автомобиль…)",
                                     style="Link.TButton", command=self.edit_constants, cursor="hand2")
        self.consts_btn.pack(side="left")
        self.consts_note = ttk.Label(consts, text="", foreground=MUTED, font=(FONT[0], 9))
        self.consts_note.pack(side="left", padx=8)

        self.gen_btn = ttk.Button(main, text="Сформировать путевой лист", style="Big.TButton",
                                  command=self.generate)
        self.gen_btn.grid(row=8, column=0, columnspan=2, pady=(0, 12))

        # --- результат
        box = ttk.Frame(main)
        box.grid(row=9, column=0, columnspan=2, sticky="nsew")
        main.rowconfigure(9, weight=1)
        self.result = tk.Text(box, height=12, font=("Consolas", 9), wrap="none", relief="solid",
                              bd=1, state="disabled", bg="#fafafa")
        scroll = ttk.Scrollbar(box, command=self.result.yview)
        self.result.configure(yscrollcommand=scroll.set)
        self.result.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        actions = ttk.Frame(main)
        actions.grid(row=10, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.open_btn = ttk.Button(actions, text="Открыть файл", command=self.open_file)
        self.open_btn.pack(side="left")
        self.folder_btn = ttk.Button(actions, text="Показать в папке", command=self.show_in_folder)
        self.folder_btn.pack(side="left", padx=8)

        # --- справочник
        foot = ttk.Frame(main)
        foot.grid(row=11, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        self.db_label = ttk.Label(foot, text="", foreground=MUTED, font=(FONT[0], 9))
        self.db_label.pack(side="left")
        ttk.Button(foot, text="Обновить справочник из Excel…", style="Link.TButton",
                   command=self.browse_shops, cursor="hand2").pack(side="right")
        self._update_db_label()

        self._set_form_state(False)
        self.open_btn.state(["disabled"])
        self.folder_btn.state(["disabled"])

    # ------------------------------------------------------------ вспомогательное
    def _set_drop_text(self, lines=None, loaded=False):
        if lines is None:
            how = "Перетащите сюда" if HAS_DND else "Нажмите здесь, чтобы выбрать"
            lines = [f"{how} путевой лист за прошлый месяц (.xlsm)",
                     "или нажмите, чтобы выбрать файл" if HAS_DND else ""]
        self.drop.configure(text="\n".join(l for l in lines if l),
                            bg=DROP_BG_LOADED if loaded else DROP_BG)

    def _set_form_state(self, enabled):
        flag = ["!disabled"] if enabled else ["disabled"]
        for w in (self.office_entry, self.liters_entry, self.norm_entry, self.weekends_check,
                  self.consts_btn, self.gen_btn):
            w.state(flag)

    def _show_result(self, text):
        self.result.configure(state="normal")
        self.result.delete("1.0", "end")
        self.result.insert("1.0", text)
        self.result.configure(state="disabled")

    def _update_db_label(self):
        self.db_label.configure(text=f"Справочник: {len(self.db['shops'])} магазинов")

    # ------------------------------------------------------------ файлы
    def browse(self):
        path = filedialog.askopenfilename(
            title="Путевой лист за прошлый месяц",
            filetypes=[("Excel", "*.xlsm *.xlsx"), ("Все файлы", "*.*")])
        if path:
            self.open_path(path)

    def browse_shops(self):
        path = filedialog.askopenfilename(title="Список магазинов",
                                          filetypes=[("Excel", "*.xlsx *.xlsm"), ("Все файлы", "*.*")])
        if path:
            self.import_shops(path)

    def on_drop(self, event):
        paths = self.root.tk.splitlist(event.data)
        if paths:
            self.open_path(paths[0])

    def open_path(self, path):
        if not path.lower().endswith((".xlsm", ".xlsx")):
            messagebox.showerror(APP_TITLE, "Нужен файл Excel (.xlsm или .xlsx).")
            return
        if C.is_trip_sheet(path):
            self.load_prev(path)
        elif C.is_shop_list(path):
            self.import_shops(path)
        else:
            messagebox.showerror(APP_TITLE, "Файл не похож ни на путевой лист, ни на список магазинов.")

    def load_prev(self, path):
        try:
            info = C.read_prev(path)
        except C.UserError as e:
            messagebox.showerror(APP_TITLE, str(e))
            return
        self.info = info
        self.changes = {}
        self.out_path = None
        head, sub = C.describe_prev(info)
        self._set_drop_text([f"✓ {head}", sub, f"Файл: {os.path.basename(path)}"], loaded=True)
        self.office_var.set(info["office"])
        self.norm_var.set(str(info["norm"]).replace(".", ","))
        self.rest_var.set(str(info["rest_end"]).replace(".", ","))
        self.liters_var.set("")
        self.consts_note.configure(text="")
        self._set_form_state(True)
        self.open_btn.state(["disabled"])
        self.folder_btn.state(["disabled"])
        self._show_result("")
        self.update_office_status()
        self.liters_entry.focus_set()
        if info["warning"]:
            messagebox.showwarning(APP_TITLE, info["warning"])

    def import_shops(self, path):
        if not messagebox.askyesno(
                APP_TITLE, f"Заменить справочник магазинов списком из файла\n«{os.path.basename(path)}»?\n\n"
                           f"Сейчас в справочнике {len(self.db['shops'])} магазинов."):
            return
        try:
            rep = C.import_shops(path, self.db)
            C.save_db(self.db, self.db_path)
        except C.UserError as e:
            messagebox.showerror(APP_TITLE, str(e))
            return
        except OSError as e:
            messagebox.showerror(APP_TITLE, f"Не удалось сохранить справочник: {e}")
            return
        def few(names):
            return ", ".join(names[:5]) + (" …" if len(names) > 5 else "")
        msg = [f"Справочник обновлён: {rep['total']} магазинов."]
        if rep["added"]:
            msg.append(f"Новые ({len(rep['added'])}): {few(rep['added'])}")
        if rep["removed"]:
            msg.append(f"Удалены ({len(rep['removed'])}): {few(rep['removed'])}")
        if rep["changed"]:
            msg.append(f"Сменили адрес ({len(rep['changed'])}): {few(rep['changed'])}")
        messagebox.showinfo(APP_TITLE, "\n\n".join(msg))
        self._update_db_label()
        self.update_office_status()

    # ------------------------------------------------------------ адрес выезда
    def update_office_status(self):
        if not self.info:
            self.office_status.configure(text="")
            return
        office = self.office_var.get().strip()
        try:
            town, points, _ = C.compute_points(self.db, office)
        except C.UserError:
            self.office_status.configure(
                text="Не удалось определить город — укажите его, например «г. Мурманск, …»",
                foreground=ERROR)
            return
        note = ("взят из прошлого листа — проверьте, не изменился ли"
                if office == self.info["office"] else "адрес изменён")
        self.office_status.configure(
            text=f"{town} · магазинов для поездок: {len(points)} · {note}",
            foreground=MUTED if points else ERROR)

    # ------------------------------------------------------------ постоянные данные
    def edit_constants(self):
        if not self.info:
            return
        win = tk.Toplevel(self.root)
        win.title("Постоянные данные")
        win.transient(self.root)
        win.grab_set()
        frm = ttk.Frame(win, padding=16)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)
        ttk.Label(frm, text="Измените то, что поменялось. Остальное оставьте как есть.",
                  foreground=MUTED).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        consts = self.info["consts"]
        vars_ = {}
        for i, (cell, label, _) in enumerate(C.FIELDS, start=1):
            ttk.Label(frm, text=label).grid(row=i, column=0, sticky="w", padx=(0, 12), pady=2)
            v = tk.StringVar(value=C.format_field(self.changes.get(cell, consts.get(cell))))
            ttk.Entry(frm, textvariable=v, width=60).grid(row=i, column=1, sticky="ew", pady=2)
            vars_[cell] = v

        def save():
            new = {}
            try:
                for cell, label, kind in C.FIELDS:
                    text = vars_[cell].get().strip()
                    if text != C.format_field(consts.get(cell)).strip():
                        new[cell] = C.parse_field(text, kind) if text else ""
            except C.UserError as e:
                messagebox.showerror(APP_TITLE, str(e), parent=win)
                return
            self.changes = new
            self.consts_note.configure(text=f"изменено полей: {len(new)}" if new else "")
            win.destroy()

        btns = ttk.Frame(frm)
        btns.grid(row=len(C.FIELDS) + 1, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Отмена", command=win.destroy).pack(side="right")
        ttk.Button(btns, text="Сохранить", command=save).pack(side="right", padx=8)

    # ------------------------------------------------------------ генерация
    def generate(self):
        if not self.info:
            return
        try:
            office = self.office_var.get().strip()
            if not office:
                raise C.UserError("Укажите адрес выезда.")
            liters = C.parse_decimal(self.liters_var.get(), "Выдано топлива")
            norm = C.parse_decimal(self.norm_var.get(), "Норма расхода")
            res = C.generate(self.info, self.db, office, liters, norm, self.weekends_var.get())
        except C.UserError as e:
            messagebox.showerror(APP_TITLE, str(e))
            return

        out = filedialog.asksaveasfilename(
            title="Сохранить путевой лист",
            initialdir=os.path.dirname(os.path.abspath(self.info["path"])),
            initialfile=C.output_name(self.info, res, self.changes),
            defaultextension=".xlsm", filetypes=[("Excel с макросами", "*.xlsm")])
        if not out:
            return
        if os.path.abspath(out) == os.path.abspath(self.info["path"]):
            messagebox.showerror(APP_TITLE, "Нельзя перезаписать путевой лист за прошлый месяц — "
                                            "выберите другое имя файла.")
            return
        try:
            C.write_result(self.info, res, out, self.changes)
        except C.UserError as e:
            messagebox.showerror(APP_TITLE, str(e))
            return
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Не удалось сохранить файл:\n{e}")
            return
        self.out_path = out
        self._show_result(f"Готово: {os.path.basename(out)}\n\n" + C.summary_text(res))
        self.open_btn.state(["!disabled"])
        self.folder_btn.state(["!disabled"])

    def open_file(self):
        if self.out_path:
            try:
                os.startfile(self.out_path)
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"Не удалось открыть файл: {e}")

    def show_in_folder(self):
        if self.out_path:
            subprocess.Popen(["explorer", "/select,", os.path.normpath(self.out_path)])


def main():
    try:                                   # чёткий текст на мониторах с масштабированием
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
