from threading import Thread
from time import time
from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox

from src.consts import path_template_file, path_out_dir
from src.logic import read_file, make_invoice_file

window = Tk()
file_path = StringVar()
start_button = None
message = StringVar()
message_label = None


def check_button_state():
    global start_button
    start_button["state"] = NORMAL
    if file_path.get() is '':
        start_button["state"] = DISABLED


def select_file():
    global file_path
    file_name = fd.askopenfile(title='Выберите файл данных', filetypes=[('xlsx files', ['.xlsx'])])
    file_path.set(file_name.name)
    check_button_state()


def work():
    global message

    start_time = time()

    message.set('Чтение данных')
    invoices = read_file(file_path.get())
    message.set('Чтение данных завершено')
    i = 0
    size = len(invoices)
    for invoice in invoices:
        make_invoice_file([""] + [str(a) for a in invoice], path_template_file, path_out_dir)
        i += 1
        current_time = time()
        message.set('Выполнено ' + str(i) + '/' + str(size) + ', прошло: ' + str(int(current_time - start_time)) + 'с')
        if i == size:
            messagebox.showinfo("Готово", "Обработка завершена")
            window.quit()


def start():
    global file_path
    global message

    start_button["state"] = DISABLED
    message.set('Обрабротка начата')
    thread = Thread(target=work)
    thread.start()


def main():
    global window
    global file_path
    global start_button
    global message
    global message_label

    window.title("Создание накладных")
    window.geometry('332x111')

    file_label = Label(text="Путь к файлу:")
    file_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
    file_input = Entry(textvariable=file_path)
    file_input.grid(row=0, column=1, padx=5, pady=5)
    file_button = Button(window, text="Выбрать файл", command=select_file)
    file_button.grid(column=2, row=0, padx=5, pady=5, sticky="e")

    message_label = Label(textvariable=message)
    message_label.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="w")
    message.set("Заполните поля, нажмите кнопку."[0:35])

    start_button = Button(window, text="Начать", command=start)
    start_button.grid(column=2, row=2, padx=5, pady=5, sticky="e")
    start_button["state"] = DISABLED

    window.mainloop()


if __name__ == '__main__':
    main()
