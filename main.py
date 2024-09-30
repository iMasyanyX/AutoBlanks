import win32com.client as win32
import tkinter as tk
from tkinter import ttk
import tkinter.messagebox as messagebox
import os

# Функция для чтения конфигурационного файла
def read_config():
    config = {}
    with open('config.txt', 'r', encoding='utf-8') as file:
        lines = file.readlines()
        shop = ''
        for line in lines:
            if line.strip():  # Пропускаем пустые строки
                if ':' in line:
                    shop, files = line.split(':')
                    shop = shop.strip()
                    config[shop] = {'files': {}}
                else:
                    file_name, copies = line.split(',')
                    file_name = file_name.strip()
                    copies = int(copies.strip())
                    # Преобразуем относительный путь в абсолютный путь
                    file_path = os.path.join(os.getcwd(), file_name)
                    config[shop]['files'][file_path] = copies
    return config

# Функция для печати Excel документа на выбранном принтере
def print_excel_file(file_path, num_copies):
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(file_path)
        worksheet = workbook.ActiveSheet
        worksheet.PrintOut(From=1, To=1, Copies=num_copies)
        workbook.Close(False)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при печати: {str(e)}")
    finally:
        excel.Quit()

# Основная функция для создания графического интерфейса
def main():
    window = tk.Tk()
    window.title("Печать документов")
    window.geometry("150x600")

    config = read_config()

    def select_shop(shop_name):
        for file_path, num_copies in config[shop_name]['files'].items():
            print_excel_file(file_path, num_copies)

        messagebox.showinfo("Готово", "Печать завершена!")

    for shop_name in config.keys():
        shop_button = ttk.Button(window, text=shop_name, command=lambda shop=shop_name: select_shop(shop))
        shop_button.pack(pady=5)

    window.mainloop()

if __name__ == '__main__':
    main()