# -*- coding: cp936 -*-
import Tkinter as tk
import tkFileDialog

import ExcelProcess

window = tk.Tk()
window.title('Copy excel')
window.geometry('500x300')
cust_input = tk.Entry(window)
# Entry�ĵ�һ�������Ǹ����ڣ��������window
# *��ʾ������ı���Ϊ�Ǻţ���Entry���ɼ����ݣ���ΪNone���ʾΪ�����ı���ԭ��ʽ�ɼ�
cust_input.grid(row=1, column=1, sticky='W', pady=20, padx=10,)


def insert_point():
    var = e.get()
    print var


OPEN_FILE_NAME = ''


def open_files():
    global OPEN_FILE_NAME
    OPEN_FILE_NAME = tkFileDialog.askopenfilename(initialdir='d:/')
    openFileLabel.config(text=OPEN_FILE_NAME)
    # ExcelProcess.copyExcel(filename, e.get())


def copy_excel():
    ExcelProcess.copyExcel(OPEN_FILE_NAME, cust_input.get())


fileButton = tk.Button(window, text=u'ѡ��excel�ļ�...', width=20, command=open_files)
fileButton.grid(row=0, column=0, sticky='W')

openFileLabel = tk.Label(window, text='')
openFileLabel.grid(row=0, column=1, pady=20, padx=10, sticky='W')

inputLabel = tk.Label(window, text=u'������customer code: ')
inputLabel.grid(row=1, column=0, sticky='W', pady=20, padx=10)

fileButton = tk.Button(window, text=u'����excel', width=20, command=copy_excel)
fileButton.grid(row=2, column=0, sticky='W')

window.mainloop()
