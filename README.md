# gendoc2

import tkinter as tk
from tkinter import messagebox
from docx import Document
from datetime import datetime
import os
import sys
from petrovich.main import Petrovich
from petrovich.enums import Case, Gender

# --- корректная дата ---
def format_date():
    months = [
        "января", "февраля", "марта", "апреля", "мая", "июня",
        "июля", "августа", "сентября", "октября", "ноября", "декабря"
    ]
    now = datetime.now()
    day = f"«{now.day:02d}»"
    month = months[now.month - 1]
    year = now.year
    return f"{day} {month} {year} г."

# --- определение пола по отчеству ---
def detect_gender(middle_name: str) -> Gender:
    middle_name = middle_name.strip().lower()
    if middle_name.endswith("ич"):
        return Gender.MALE
    elif middle_name.endswith("на"):
        return Gender.FEMALE
    else:
        return Gender.FEMALE  # безопасный вариант

# --- склонение ФИО в родительный падеж ---
petrovich = Petrovich()

def to_genitive(fio: str) -> str:
    parts = fio.strip().split()
    if len(parts) != 3:
        return fio  # безопасный откат

    last, first, middle = parts
    gender = detect_gender(middle)

    try:
        last_g = petrovich.lastname(last, Case.GENITIVE, gender)
        first_g = petrovich.firstname(first, Case.GENITIVE, gender)
        middle_g = petrovich.middlename(middle, Case.GENITIVE, gender)
        return f"{last_g} {first_g} {middle_g}"
    except Exception:
        return fio

# --- универсальная замена плейсхолдера ---
def replace_placeholder(paragraph, placeholder, value):
    if placeholder not in paragraph.text:
        return False
    text = paragraph.text.replace(placeholder, value)
    paragraph.clear()
    paragraph.add_run(text)
    return True

# --- путь к template.docx (чтобы работало и в exe) ---
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- создание документа ---
def create_document():
    fio = fio_entry.get().strip()
    inn = inn_entry.get().strip()

    if not fio or not inn:
        messagebox.showerror("Ошибка", "Заполните ФИО и ИНН")
        return

    try:
        doc = Document(resource_path("template.docx"))

        date_text = format_date()
        fio_gen = to_genitive(fio)

        # словарь плейсхолдеров
        placeholders = {
            "{{DATE}}": date_text,
            "{{FIO}}": fio,
            "{{INN}}": inn,
            "{{FIO_GEN}}": fio_gen
        }

        # замена плейсхолдеров в параграфах
        for paragraph in doc.paragraphs:
            for key, value in placeholders.items():
                replace_placeholder(paragraph, key, value)

        # сохраняем документ
        output_name = f"Документ_{fio.split()[0]}.docx"
        doc.save(output_name)

        messagebox.showinfo("Готово", f"Документ создан:\n{output_name}")

    except Exception as e:
        messagebox.showerror("Ошибка", str(e))

# --- интерфейс ---
root = tk.Tk()
root.title("Заполнение документа")
root.geometry("420x260")
root.resizable(False, False)

tk.Label(root, text="ФИО", font=("Arial", 11)).pack(pady=(20, 5))
fio_entry = tk.Entry(root, width=45)
fio_entry.pack()

tk.Label(root, text="ИНН", font=("Arial", 11)).pack(pady=(15, 5))
inn_entry = tk.Entry(root, width=45)
inn_entry.pack()

tk.Button(
    root,
    text="Создать документ",
    font=("Arial", 11, "bold"),
    command=create_document
).pack(pady=30)

root.mainloop()
