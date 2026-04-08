import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

from comparator import ExcelComparator


# Все строки интерфейса вынесены в переменные для удобной локализации.
UI_TEXT = {
    "title": "Сравнение Excel: Транспортные средства",
    "select_old": "Выбрать старый файл",
    "select_new": "Выбрать новый файл",
    "compare": "▶ СРАВНИТЬ",
    "save_report": "💾 Сохранить отчёт",
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


class ExcelCompareApp:
    """Главный класс GUI-приложения."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(UI_TEXT["title"])
        self.root.geometry("900x600")
        self.root.minsize(760, 520)

        self.comparator = ExcelComparator()
        self.old_file_path = ""
        self.new_file_path = ""
        self.last_report = ""

        self._create_widgets()

    def _create_widgets(self):
        top_frame = tk.Frame(self.root, padx=12, pady=10)
        top_frame.pack(fill=tk.X)

        self.btn_old = tk.Button(top_frame, text=UI_TEXT["select_old"], command=self.select_old_file, width=24)
        self.btn_old.grid(row=0, column=0, padx=(0, 8), pady=5, sticky="w")

        self.lbl_old = tk.Label(top_frame, text="—", anchor="w")
        self.lbl_old.grid(row=0, column=1, padx=4, pady=5, sticky="we")

        self.btn_new = tk.Button(top_frame, text=UI_TEXT["select_new"], command=self.select_new_file, width=24)
        self.btn_new.grid(row=1, column=0, padx=(0, 8), pady=5, sticky="w")

        self.lbl_new = tk.Label(top_frame, text="—", anchor="w")
        self.lbl_new.grid(row=1, column=1, padx=4, pady=5, sticky="we")

        top_frame.columnconfigure(1, weight=1)

        action_frame = tk.Frame(self.root, padx=12, pady=4)
        action_frame.pack(fill=tk.X)

        self.btn_compare = tk.Button(
            action_frame,
            text=UI_TEXT["compare"],
            command=self.compare_files,
            font=("Segoe UI", 10, "bold"),
            width=18,
        )
        self.btn_compare.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_save = tk.Button(action_frame, text=UI_TEXT["save_report"], command=self.save_report, width=18)
        self.btn_save.pack(side=tk.LEFT)

        self.output = ScrolledText(self.root, wrap=tk.WORD, font=("Consolas", 10))
        self.output.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        self.status_var = tk.StringVar(value=UI_TEXT["status_ready"])
        self.status_bar = tk.Label(
            self.root,
            textvariable=self.status_var,
            bd=1,
            relief=tk.SUNKEN,
            anchor="w",
            padx=8,
            pady=4,
        )
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def set_status(self, text: str):
        """Обновляет текст в статус-баре."""
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
