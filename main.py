import os
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import json
from tkinterdnd2 import TkinterDnD, DND_FILES

TEMPLATES_FILE = 'email_templates.json'

class OutlookApp:
    def __init__(self):
        self.templates = self.load_templates()
        self.template_names = list(self.templates.keys())

        self.root = TkinterDnD.Tk()
        self.root.title("Outlook Email Creator")

        # Поля для ввода
        tk.Label(self.root, text="Получатель").grid(row=0, column=0)
        self.recipient_entry = self.create_entry_with_context_menu(self.root)
        self.recipient_entry.grid(row=0, column=1)

        tk.Label(self.root, text="Копия").grid(row=1, column=0)
        self.cc_entry = self.create_entry_with_context_menu(self.root)
        self.cc_entry.grid(row=1, column=1)

        tk.Label(self.root, text="Тема").grid(row=2, column=0)
        self.subject_entry = self.create_entry_with_context_menu(self.root)
        self.subject_entry.grid(row=2, column=1)

        tk.Label(self.root, text="Текст").grid(row=3, column=0)
        self.body_text = self.create_text_with_context_menu(self.root)
        self.body_text.grid(row=3, column=1)

        # Кнопка для добавления вложений
        tk.Button(self.root, text="Добавить вложение", command=self.add_attachment).grid(row=4, column=0)
        self.attachments_list = tk.Listbox(self.root, width=50, height=5)
        self.attachments_list.grid(row=4, column=1)

        # Кнопка отправки письма
        tk.Button(self.root, text="Отправить письмо", command=self.send_email).grid(row=5, column=1)

        # Выбор шаблона
        tk.Label(self.root, text="Выбрать шаблон").grid(row=6, column=0)
        self.selected_template = tk.StringVar(self.root)
        self.selected_template.set("Нет шаблона")

        # Исправленный вызов OptionMenu
        self.template_menu = tk.OptionMenu(self.root, self.selected_template, self.template_names)
        self.template_menu.grid(row=6, column=1)

        # Привязка события к изменению выбора шаблона
        self.selected_template.trace("w", self.on_template_selected)

        # Кнопки для работы с шаблонами
        tk.Button(self.root, text="Сохранить как шаблон", command=self.save_template).grid(row=7, column=0)
        tk.Button(self.root, text="Удалить шаблон", command=self.delete_template).grid(row=7, column=1)
        tk.Button(self.root, text="Очистить", command=self.clear_fields).grid(row=7, column=2)

        # Обработка вложений
        self.attachments = []

        # Поддержка drag-and-drop
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.drop)

        # Обновляем меню шаблонов
        self.update_template_menu()

        self.root.mainloop()

    def create_entry_with_context_menu(self, root):
        entry = tk.Entry(root, width=50)
        self.add_context_menu(entry)
        return entry

    def create_text_with_context_menu(self, root):
        text = tk.Text(root, height=10, width=50)
        self.add_context_menu(text)
        return text

    def add_context_menu(self, widget):
        context_menu = tk.Menu(widget, tearoff=0)
        context_menu.add_command(label="Копировать", command=lambda: self.copy(widget))
        context_menu.add_command(label="Вставить", command=lambda: self.paste(widget))
        context_menu.add_command(label="Вырезать", command=lambda: self.cut(widget))
        widget.bind("<Button-3>", lambda event: context_menu.tk_popup(event.x_root, event.y_root))

    def add_attachment(self):
        filepaths = filedialog.askopenfilenames(title="Выберите файлы")
        if filepaths:
            for filepath in filepaths:
                self.attachments.append(filepath)
                self.attachments_list.insert(tk.END, filepath)

    def drop(self, event):
        files = self.root.tk.splitlist(event.data)
        for file in files:
            self.attachments.append(file)
            self.attachments_list.insert(tk.END, file)

    def send_email(self):
        recipient = self.recipient_entry.get()
        cc = self.cc_entry.get()
        subject = self.subject_entry.get()
        body = self.body_text.get("1.0", tk.END).strip()

        if not recipient or not subject:
            messagebox.showerror("Ошибка", "Поля 'Получатель' и 'Тема' обязательны для заполнения!")
            return

        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = recipient
            mail.CC = cc
            mail.Subject = subject
            mail.Body = body

            # Добавляем вложения
            for attachment in self.attachments:
                mail.Attachments.Add(attachment)

            mail.Send()
            messagebox.showinfo("Успех", "Письмо отправлено!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось отправить письмо: {e}")

    def save_template(self):
        template_name = simpledialog.askstring("Шаблон", "Введите название шаблона")
        if template_name:
            self.templates[template_name] = {
                "recipient": self.recipient_entry.get(),
                "cc": self.cc_entry.get(),
                "subject": self.subject_entry.get(),
                "body": self.body_text.get("1.0", tk.END).strip(),
                "attachments": self.attachments
            }
            self.save_templates()
            self.update_template_menu()
            messagebox.showinfo("Успех", f"Шаблон '{template_name}' сохранен!")

    def load_template(self, template_name):
        if template_name in self.templates:
            template = self.templates[template_name]
            self.recipient_entry.delete(0, tk.END)
            self.recipient_entry.insert(0, template["recipient"])

            self.cc_entry.delete(0, tk.END)
            self.cc_entry.insert(0, template["cc"])

            self.subject_entry.delete(0, tk.END)
            self.subject_entry.insert(0, template["subject"])

            self.body_text.delete("1.0", tk.END)
            self.body_text.insert("1.0", template["body"])

            self.attachments_list.delete(0, tk.END)
            self.attachments = template["attachments"]
            for attachment in self.attachments:
                self.attachments_list.insert(tk.END, attachment)

    def delete_template(self):
        template_name = self.selected_template.get()
        if template_name in self.templates:
            del self.templates[template_name]
            self.save_templates()
            self.update_template_menu()

            messagebox.showinfo("Успех", f"Шаблон '{template_name}' удален!")

    def clear_fields(self):
        self.recipient_entry.delete(0, tk.END)
        self.cc_entry.delete(0, tk.END)
        self.subject_entry.delete(0, tk.END)
        self.body_text.delete("1.0", tk.END)
        self.attachments_list.delete(0, tk.END)
        self.attachments = []

    def update_template_menu(self):
        self.template_names = list(self.templates.keys())
        menu = self.template_menu['menu']
        menu.delete(0, 'end')

        for name in self.template_names:
            menu.add_command(label=name, command=lambda value=name: self.selected_template.set(value))

        if self.template_names:
            self.selected_template.set(self.template_names[0])
        else:
            self.selected_template.set("Нет шаблона")

    def on_template_selected(self, *args):
        selected_template_name = self.selected_template.get()
        if selected_template_name != "Нет шаблона":
            self.load_template(selected_template_name)

    def load_templates(self):
        if os.path.exists(TEMPLATES_FILE):
            with open(TEMPLATES_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {}

    def save_templates(self):
        with open(TEMPLATES_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.templates, f, indent=4, ensure_ascii=False)

if __name__ == "__main__":
    OutlookApp()
