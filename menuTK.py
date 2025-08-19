import locale
import tkinter as tk
from tkinter import Tk, Button, Text, Frame, BOTH, filedialog, END, Label, Entry, messagebox, Checkbutton
import json

# Мои импорты
import perechen_elementov
import specification
import vedomost_pokupnih_izdeliy
import ERI

locale.setlocale(locale.LC_ALL, ('ru_RU', 'UTF-8'))

files_to_open = []
data = {}

# Проверяем, если файл Profile.json

try:
    with open("Profile.json", "r") as f:
        data = json.load(f)
        data = {
            "Project_Name": "",
            "Templates_Path": "",

            "Razrab": "",
            "Proveril": "",
            "N_control": "",
            "Utverdil": "",

            "Scheme": True,
            "PE": True,
            "DK": True,
            "I1": True,
            "I2": True,

            "SB_count": 1,
            "Precent": 10
        }


except (FileExistsError, FileNotFoundError):
    with open("Profile.json", "w") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
        data = {
            "Project_Name": "",
            "Templates_Path": "",

            "Razrab": "",
            "Proveril": "",
            "N_control": "",
            "Utverdil": "",
            "Scheme": True,

            "PE": True,
            "DK": True,
            "I1": True,
            "I2": True,

            "SB_count": 1,

            "Precent": 10
        }


class UI(Frame):
    def __init__(self):
        super().__init__()

        self.lb_Utverdil = Label(self, width=30, height=1,
                                 text="Утвердил")
        self.ent_Utverdil = Entry(self, width=30)
        self.lb_Control = Label(self, width=30, height=1,
                                text="Н. контроль")
        self.ent_Control = Entry(self, width=30)
        self.lb_Checked = Label(self, width=30, height=1,
                                text="Проверил")
        self.ent_Checked = Entry(self, width=30)
        self.lb_Razrab = Label(self, width=30, height=1,
                               text="Разработчик")
        self.ent_Razrab = Entry(self, width=30)
        self.lb_Project = Label(self, width=30, height=1,
                                text="Название проекта")
        self.ent_Project = Entry(self, width=30)
        self.lb_Templates = Label(self, width=30, height=1,
                                  text="Путь к шаблонам")
        self.ent_Templates = Entry(self, width=30)
        self.btn_ERI = Button(self, text="ЭРИ",
                              command=self.createERI)
        self.btn_TEO = Button(self, text="ТЭО")
        self.btn_VP = Button(self, text="ВП",
                             command=self.createVedomost)
        self.btn_Specefication = Button(self, text="Спецификация",
                                        command=self.createSpecification)
        self.btn_Perechen = Button(self, text="Перечень",
                                   command=self.createPerechen)
        self.btn_AddFile = Button(self, text="Добавить файл/-ы",
                                  command=self.onOpen)
        self.btn_RemoveFiles = Button(self, text="Очистить",
                                      command=self.clearFiles)
        self.txt_FilesListView = Text(self, width=30, height=10)
        self.is_civ = tk.BooleanVar()
        self.c1 = Checkbutton(self, text='Гражданский', variable=self.is_civ, onvalue=1, offvalue=0)

        self.initUI()

    def initUI(self):
        self.master.title("Генератор документов")
        self.pack(fill=BOTH, expand=1)

        self.columnconfigure([0, 1], weight=1)
        self.rowconfigure([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11], weight=1)

        self.btn_RemoveFiles.grid(column=0, row=0,
                                  padx=2, pady=2)

        self.txt_FilesListView.grid(column=0, row=1, rowspan=3,
                                    padx=2, pady=2)

        self.btn_AddFile.grid(column=0, row=4,
                              padx=2, pady=2)

        self.btn_Perechen.grid(column=1, row=0,
                               padx=2, pady=2, sticky="ew")
        self.btn_Specefication.grid(column=1, row=1,
                                    padx=2, pady=2, sticky="ew")
        self.btn_VP.grid(column=1, row=2,
                         padx=2, pady=2, sticky="ew")
        self.btn_TEO.grid(column=1, row=3,
                          padx=2, pady=2, sticky="ew")
        self.btn_ERI.grid(column=1, row=4,
                          padx=2, pady=2, sticky="ew")

        self.ent_Templates.grid(column=0, row=5,
                                padx=2, pady=2)
        self.ent_Templates.bind("<Key>", self.onModified)

        self.lb_Templates.grid(column=1, row=5,
                               padx=2, pady=2)

        self.ent_Project.grid(column=0, row=6,
                              padx=2, pady=2)
        self.ent_Project.bind("<Key>", self.onModified)

        self.lb_Project.grid(column=1, row=6,
                             padx=2, pady=2)

        self.ent_Razrab.grid(column=0, row=7,
                             padx=2, pady=2)
        self.ent_Razrab.bind("<Key>", self.onModified)

        self.lb_Razrab.grid(column=1, row=7,
                            padx=2, pady=2)

        self.ent_Checked.grid(column=0, row=8,
                              padx=2, pady=2)
        self.ent_Checked.bind("<Key>", self.onModified)

        self.lb_Checked.grid(column=1, row=8,
                             padx=2, pady=2)

        self.ent_Control.grid(column=0, row=9,
                              padx=2, pady=2)
        self.ent_Control.bind("<Key>", self.onModified)

        self.lb_Control.grid(column=1, row=9,
                             padx=2, pady=2)

        self.ent_Utverdil.grid(column=0, row=10,
                               padx=2, pady=2)
        self.ent_Utverdil.bind("<Key>", self.onModified)

        self.lb_Utverdil.grid(column=1, row=10,
                              padx=2, pady=2)

        self.c1.grid(column=1, row=11,
                     padx=2, pady=2)

    def onOpen(self):
        if not self.txt_FilesListView.compare("end-1c", "==", "1.0"):
            self.txt_FilesListView.delete("0.0", END)
        for file in filedialog.askopenfilenames(multiple=1, filetypes=[("Excel files", ".xlsx .xls")]):
            files_to_open.append(file)

        if files_to_open != '':
            for file in files_to_open:
                self.txt_FilesListView.insert(END, file.replace("\\\\", "/").split("/")[-1] + "\n")

    def clearFiles(self):
        files_to_open.clear()
        if not self.txt_FilesListView.compare("end-1c", "==", "1.0"):
            self.txt_FilesListView.delete("0.0", END)

    def createPerechen(self):
        if len(files_to_open) != 0:
            perechen_elementov.execute(files_to_open, self.is_civ.get())

            self.txt_FilesListView.insert(END, "\nПЭ создан(-ы)!")
        else:
            messagebox.showerror("Ошибка", "Не выбрано ни одного файла!")

    def createSpecification(self):
        if len(files_to_open) != 0:
            specification.execute(files_to_open, self.is_civ.get())

            self.txt_FilesListView.insert(END, "\nСП создана(-ы)!")
        else:
            messagebox.showerror("Ошибка", "Не выбрано ни одного файла!")

    def createVedomost(self):
        if len(files_to_open) != 0:
            vedomost_pokupnih_izdeliy.execute(files_to_open)

            self.txt_FilesListView.insert(END, "\nВП создана!")
        else:
            messagebox.showerror("Ошибка", "Не выбрано ни одного файла!")

    def createERI(self):
        if len(files_to_open) != 0:
            ERI.execute(files_to_open)

            self.txt_FilesListView.insert(END, "\nЭРИ создан!")
        else:
            messagebox.showerror("Ошибка", "Не выбрано ни одного файла!")

    @staticmethod
    def onModified(event):
        data['Templates_Path'] = ui.ent_Templates.get()
        data['Project_Name'] = ui.ent_Project.get()

        data['Razrab'] = ui.ent_Razrab.get()
        data['Proveril'] = ui.ent_Checked.get()
        data['N_control'] = ui.ent_Control.get()
        data['Utverdil'] = ui.ent_Utverdil.get()

        with open("Profile.json", "w") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

root = Tk()
ui = UI()

def on_closing():
    global data

    data['Templates_Path'] = ui.ent_Templates.get()
    data['Project_Name'] = ui.ent_Project.get()

    data['Razrab'] = ui.ent_Razrab.get()
    data['Proveril'] = ui.ent_Checked.get()
    data['N_control'] = ui.ent_Control.get()
    data['Utverdil'] = ui.ent_Utverdil.get()

    with open("Profile.json", "r") as f:
        data = json.load(f)

    root.destroy()

def main():
    ui.ent_Templates.insert(0, data['Templates_Path'])
    ui.ent_Project.insert(0, data['Project_Name'])

    ui.ent_Razrab.insert(0, data['Razrab'])
    ui.ent_Checked.insert(0, data['Proveril'])
    ui.ent_Control.insert(0, data['N_control'])
    ui.ent_Utverdil.insert(0, data['Utverdil'])

    root.geometry("400x400")
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()
