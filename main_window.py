import tkinter as tk
from tkinter import filedialog as fd
from services import main_process
import sys
import os

class SomeClass():
    def __init__(self):
        self.param = None

    def set_param(self, param):
        self.param = param

    def get_param(self):
        return self.param
    
    def del_param(self):
        self.param = None
    
dialog = SomeClass()

win = tk.Tk()
win.title("Помогатор")

def img_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
    
#Настройка окна 
path_logo = img_path("icon.png")
win_width = 300
win_height = 100
screen_width = win.winfo_screenwidth()
screen_height = win.winfo_screenheight()
win.geometry(f"{win_width}x{win_height}+{int((screen_width-win_width)/2)}+{int((screen_height-win_height)/2)}")
win.resizable(False, False)
logo = tk.PhotoImage(file=path_logo)
win.iconphoto(False,logo)
    
# Получение пути к рабочему каталогу  
def directory():
    dir_path = fd.askdirectory(title="Выбор папки")
    dialog.set_param(dir_path)
    
#Основная функция
def working_def():
    print("STARTING!!!")
  
#Подсказка_1
label_choice = tk.Label(win, text="Выбери рабочую папку:", font=("Arial", 12))
label_choice.place(relx=0.05, rely=0.15)

#Подсказка_2
label_start = tk.Label(win, text="Нажми на 'Запуск!':", font=("Arial", 12))
label_start.place(relx=0.05, rely=0.55)

#Рекламка
label_adv = tk.Label(win, fg="#9400D3", text="telegram: @prokopyeff", font=("Arial", 8))
label_adv.place(relx=0.05, rely=0.8)

#Кнопка выбора рабочего каталога
choice_dialog_btn = tk.Button(win, text='Выбор', font=("Arial", 12), command = lambda:directory())
choice_dialog_btn.place(relx=0.75, rely=0.1)

#Кнопка запуска основной функции
start_btn = tk.Button(win, text = "Запуск!", command=lambda:main_process(dialog.get_param(), dialog, win), font=("Arial", 12))
"""Добавить параметр state для дизактивации кнопки на время работы алгоритма"""
start_btn.place(relx=0.74, rely=0.5)

def run():
    win.mainloop()