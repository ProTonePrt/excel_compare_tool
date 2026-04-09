import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from comparator import ExcelComparator


UI_TEXT = {
    "title": "Сравнение Excel: Транспортные средства",
    "header": "Excel Compare Tool",
    "subheader": "Сравнение списков ТС по гос. номеру или VIN",
    "select_old": "Выбрать старый файл",
    "select_new": "Выбрать новый файл",
    "compare": "Сравнить",
    "save_report": "Сохранить отчёт",
    "status_ready": "Готово. Выберите файлы для сравнения.",
    "status_old_selected": "Старый файл выбран.",
    "status_new_selected": "Новый файл выбран.",
    "status_compare_start": "Сравнение выполняется...",
    "status_compare_done": "Сравнение завершено.",
    "status_saved": "Отчёт сохранён.",
    "status_error": "Ошибка. Проверьте сообщение.",
    "status_no_report": "Сначала выполните сравнение.",
    "error_title": "Ошибка",
    "warning_title": "Предупреждение",
    "info_title": "Информация",
    "choose_both_files": "Выберите старый и новый файл перед сравнением.",
    "report_empty": "Нет отчёта для сохранения.",
    "saved_ok": "Отчёт успешно сохранён.",
    "not_found_column": (
        "Не удалось найти колонку с гос. номером или VIN.\n"
        "Переименуйте столбец, например в 'Гос номер', 'VIN', 'Рег. номер'."
    ),
}


COLORS = {
    "bg": "#F3F4F6",
    "card": "#FFFFFF",
    "muted": "#6B7280",
    "text": "#111827",
    "accent": "#2563EB",
    "accent_hover": "#1D4ED8",
    "line": "#E5E7EB",
    "ok": "#16A34A",
}


class ExcelCompareApp:
    """Современный интерфейс в стиле Windows 11 на ttk/tkinter."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(UI_TEXT["title"])
        self.root.geometry("980x680")
        self.root.minsize(820, 560)
        self.root.configure(bg=COLORS["bg"])

        self.comparator = ExcelComparator()
        self.old_file_path = ""
        self.new_file_path = ""
        self.last_report = ""

        self.status_var = tk.StringVar(value=UI_TEXT["status_ready"])
        self._busy = False
        self._setup_styles()
        self._create_widgets()

    def _setup_styles(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")

        style.configure("Card.TFrame", background=COLORS["card"], relief="flat")
        style.configure("Header.TLabel", background=COLORS["bg"], foreground=COLORS["text"], font=("Segoe UI", 18, "bold"))
        style.configure("SubHeader.TLabel", background=COLORS["bg"], foreground=COLORS["muted"], font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background=COLORS["card"], foreground=COLORS["muted"], font=("Segoe UI", 9))
        style.configure("File.TLabel", background=COLORS["card"], foreground=COLORS["text"], font=("Segoe UI", 10))
        style.configure("Status.TLabel", background=COLORS["bg"], foreground=COLORS["muted"], font=("Segoe UI", 9))
        style.configure("Line.TFrame", background=COLORS["line"])
        style.configure("TProgressbar", troughcolor=COLORS["line"], background=COLORS["accent"], bordercolor=COLORS["line"], lightcolor=COLORS["accent"], darkcolor=COLORS["accent"])

    def _create_widgets(self):
        root_pad = tk.Frame(self.root, bg=COLORS["bg"])
        root_pad.pack(fill=tk.BOTH, expand=True, padx=20, pady=16)

        ttk.Label(root_pad, text=UI_TEXT["header"], style="Header.TLabel").pack(anchor="w")
        ttk.Label(root_pad, text=UI_TEXT["subheader"], style="SubHeader.TLabel").pack(anchor="w", pady=(2, 14))

        card = ttk.Frame(root_pad, style="Card.TFrame", padding=16)
        card.pack(fill=tk.BOTH, expand=True)

        top_grid = ttk.Frame(card, style="Card.TFrame")
        top_grid.pack(fill=tk.X)
        top_grid.columnconfigure(1, weight=1)

        self.btn_old = tk.Button(
            top_grid,
            text=UI_TEXT["select_old"],
            command=self.select_old_file,
            bg=COLORS["card"],
            fg=COLORS["text"],
            activebackground="#EEF2FF",
            activeforeground=COLORS["text"],
            relief="flat",
            bd=0,
            padx=14,
            pady=8,
            font=("Segoe UI", 10),
            cursor="hand2",
        )
        self.btn_old.grid(row=0, column=0, padx=(0, 12), pady=(0, 10), sticky="w")
        self._bind_hover(self.btn_old, COLORS["card"], "#EEF2FF")

        self.lbl_old = ttk.Label(top_grid, text="Файл не выбран", style="File.TLabel")
        self.lbl_old.grid(row=0, column=1, sticky="we", pady=(0, 10))

        self.btn_new = tk.Button(
            top_grid,
            text=UI_TEXT["select_new"],
            command=self.select_new_file,
            bg=COLORS["card"],
            fg=COLORS["text"],
            activebackground="#EEF2FF",
            activeforeground=COLORS["text"],
            relief="flat",
            bd=0,
            padx=14,
            pady=8,
            font=("Segoe UI", 10),
            cursor="hand2",
        )
        self.btn_new.grid(row=1, column=0, padx=(0, 12), pady=(0, 8), sticky="w")
        self._bind_hover(self.btn_new, COLORS["card"], "#EEF2FF")

        self.lbl_new = ttk.Label(top_grid, text="Файл не выбран", style="File.TLabel")
        self.lbl_new.grid(row=1, column=1, sticky="we", pady=(0, 8))

        ttk.Frame(card, style="Line.TFrame", height=1).pack(fill=tk.X, pady=(8, 12))

        action_row = ttk.Frame(card, style="Card.TFrame")
        action_row.pack(fill=tk.X, pady=(0, 8))

        self.btn_compare = tk.Button(
            action_row,
            text=f"▶ {UI_TEXT['compare']}",
            command=self.compare_files,
            bg=COLORS["accent"],
            fg="white",
            activebackground=COLORS["accent_hover"],
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=16,
            pady=9,
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
        )
        self.btn_compare.pack(side=tk.LEFT, padx=(0, 10))
        self._bind_hover(self.btn_compare, COLORS["accent"], COLORS["accent_hover"])

        self.btn_save = tk.Button(
            action_row,
            text=f"💾 {UI_TEXT['save_report']}",
            command=self.save_report,
            bg=COLORS["card"],
            fg=COLORS["text"],
            activebackground="#EEF2FF",
            activeforeground=COLORS["text"],
            relief="flat",
            bd=0,
            padx=14,
            pady=9,
            font=("Segoe UI", 10),
            cursor="hand2",
        )
        self.btn_save.pack(side=tk.LEFT)
        self._bind_hover(self.btn_save, COLORS["card"], "#EEF2FF")

        self.progress = ttk.Progressbar(card, mode="indeterminate", length=160)
        self.progress.pack(fill=tk.X, pady=(2, 10))
        self.progress.stop()

        output_frame = tk.Frame(card, bg=COLORS["line"], bd=0, highlightthickness=0)
        output_frame.pack(fill=tk.BOTH, expand=True, pady=(4, 10))

        self.output = ScrolledText(
            output_frame,
            wrap=tk.WORD,
            font=("Consolas", 10),
            bg="#FAFAFB",
            fg=COLORS["text"],
            bd=0,
            padx=10,
            pady=10,
            insertbackground=COLORS["text"],
        )
        self.output.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        status_row = ttk.Frame(root_pad, style="Card.TFrame")
        status_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(status_row, textvariable=self.status_var, style="Status.TLabel").pack(side=tk.LEFT, anchor="w")

    @staticmethod
    def _bind_hover(button: tk.Button, normal: str, hover: str):
        button.bind("<Enter>", lambda _: button.config(bg=hover))
        button.bind("<Leave>", lambda _: button.config(bg=normal))

    def _set_busy(self, is_busy: bool):
        self._busy = is_busy
        state = tk.DISABLED if is_busy else tk.NORMAL
        self.btn_compare.config(state=state)
        self.btn_old.config(state=state)
        self.btn_new.config(state=state)
        if is_busy:
            self.progress.start(10)
        else:
            self.progress.stop()

    def set_status(self, text: str):
        self.status_var.set(text)
        self.root.update_idletasks()

    def select_old_file(self):
        path = filedialog.askopenfilename(
            title=UI_TEXT["select_old"],
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.old_file_path = path
            self.lbl_old.config(text=os.path.basename(path))
            self.set_status(UI_TEXT["status_old_selected"])

    def select_new_file(self):
        path = filedialog.askopenfilename(
            title=UI_TEXT["select_new"],
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.new_file_path = path
            self.lbl_new.config(text=os.path.basename(path))
            self.set_status(UI_TEXT["status_new_selected"])

    def format_report(self, result: dict) -> str:
        """Формирует текст отчёта в требуемом формате."""
        old_name = os.path.basename(self.old_file_path)
        new_name = os.path.basename(self.new_file_path)
        unchanged_count = len(result["unchanged"])
        added = result["added"]
        removed = result["removed"]
        total_changes = len(added) + len(removed)

        lines = [
            "📊 ОТЧЁТ О СРАВНЕНИИ",
            f"Старый файл: {old_name}",
            f"Новый файл: {new_name}",
            f"Найдено колонок: \"{result['col_old']}\" → \"{result['col_new']}\"",
            f"Нормализовано ID: старый={result.get('old_total_valid', 0)}, новый={result.get('new_total_valid', 0)}",
            f"✅ Без изменений: {unchanged_count} ТС",
        ]

        if total_changes == 0:
            lines.append("✅ Изменений не обнаружено. Списки идентичны.")
            return "\n".join(lines)

        lines.append("🟢 НА УЧАСТОК ПОСТУПИЛИ:")
        if added:
            for item in added:
                lines.append(f"{self._id_prefix(item)} {item}")
        else:
            lines.append("—")

        lines.append("🔴 С УЧАСТКА ВЫБЫЛИ:")
        if removed:
            for item in removed:
                lines.append(f"{self._id_prefix(item)} {item}")
        else:
            lines.append("—")

        lines.append(f"💡 Всего изменений: {total_changes}")
        return "\n".join(lines)

    @staticmethod
    def _id_prefix(value: str) -> str:
        """
        Подписывает строку как VIN, если это похоже на VIN (17 символов),
        иначе как гос. номер.
        """
        if len(value) == 17:
            return "vin номер"
        return "гос номер"

    def compare_files(self):
        if not self.old_file_path or not self.new_file_path:
            messagebox.showwarning(UI_TEXT["warning_title"], UI_TEXT["choose_both_files"])
            self.set_status(UI_TEXT["status_error"])
            return

        try:
            self._set_busy(True)
            self.set_status(UI_TEXT["status_compare_start"])

            old_df = self.comparator.load_file(self.old_file_path)
            new_df = self.comparator.load_file(self.new_file_path)

            result = self.comparator.compare(old_df, new_df)
            report = self.format_report(result)
            self.last_report = report

            self.output.delete("1.0", tk.END)
            self.output.insert(tk.END, report)
            self.set_status(UI_TEXT["status_compare_done"])
        except ValueError as exc:
            message_text = str(exc)
            if "Не удалось найти колонку" in message_text:
                message_text = UI_TEXT["not_found_column"]
            messagebox.showerror(UI_TEXT["error_title"], message_text)
            self.set_status(UI_TEXT["status_error"])
        except FileNotFoundError as exc:
            messagebox.showerror(UI_TEXT["error_title"], str(exc))
            self.set_status(UI_TEXT["status_error"])
        except PermissionError as exc:
            messagebox.showerror(UI_TEXT["error_title"], str(exc))
            self.set_status(UI_TEXT["status_error"])
        except Exception as exc:
            messagebox.showerror(
                UI_TEXT["error_title"],
                f"Произошла непредвиденная ошибка:\n{exc}\n\n{traceback.format_exc()}",
            )
            self.set_status(UI_TEXT["status_error"])
        finally:
            self._set_busy(False)

    def save_report(self):
        if not self.last_report.strip():
            messagebox.showinfo(UI_TEXT["info_title"], UI_TEXT["report_empty"])
            self.set_status(UI_TEXT["status_no_report"])
            return

        save_path = filedialog.asksaveasfilename(
            title=UI_TEXT["save_report"],
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile="report.txt",
        )

        if not save_path:
            return

        try:
            with open(save_path, "w", encoding="utf-8") as file:
                file.write(self.last_report)
            messagebox.showinfo(UI_TEXT["info_title"], UI_TEXT["saved_ok"])
            self.set_status(UI_TEXT["status_saved"])
        except PermissionError:
            messagebox.showerror(
                UI_TEXT["error_title"],
                "Не удалось сохранить файл. Проверьте права доступа или закройте файл, если он открыт.",
            )
            self.set_status(UI_TEXT["status_error"])
        except Exception as exc:
            messagebox.showerror(UI_TEXT["error_title"], f"Ошибка при сохранении отчёта: {exc}")
            self.set_status(UI_TEXT["status_error"])


def main():
    root = tk.Tk()
    app = ExcelCompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
