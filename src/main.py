from threading import Thread
from time import time
from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox

from src.consts import path_template_file, path_out_dir
from src.logic import read_file, make_invoice_file


def start():
    global file_path
    global message

    start_button["state"] = DISABLED
    message.set('Обрабротка начата')
    thread = Thread(target=work)
    thread.start()


window = Tk()
file_path = StringVar()
start_button = Button(window, text="Начать", width=12, command=start)
message = StringVar()
message_label = Label(textvariable=message)


def check_button_state():
    global start_button
    start_button["state"] = NORMAL
    if file_path.get() == '':
        start_button["state"] = DISABLED


def select_file():
    global file_path
    file_name = fd.askopenfile(title='Выберите файл данных', filetypes=[('xlsx files', ['.xlsx'])])
    file_path.set(file_name.name)
    check_button_state()


def work():
    global message
    global file_path

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


def main():
    global window
    global file_path
    global start_button
    global message
    global message_label

    window.title("Создание накладных")
    window.geometry('580x140')

    col1x = 20
    col2x = 440
    row1y = 20
    row2y = 50
    row3y = 90
    file_label = Label(text="Путь к файлу со списком данных:")
    file_label.place(x=col1x, y=row1y)
    file_input = Entry(textvariable=file_path, width=50)
    file_input.place(x=col1x, y=row2y+4)
    file_button = Button(window, text="Выбрать файл", width=12, command=select_file)
    file_button.place(x=col2x, y=row2y)

    message_label.place(x=col1x, y=row3y+4)
    message.set('Заполните поле и нажмите "Начать"'[:35])

    start_button.place(x=col2x, y=row3y)
    start_button["state"] = DISABLED

    window.mainloop()


if __name__ == '__main__':
    main()
