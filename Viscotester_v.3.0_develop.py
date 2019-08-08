from tkinter import *
import tkinter.scrolledtext as ScrolledText
import threading
from time import sleep, asctime
import logging
import serial
import re
from collections import OrderedDict
from os import startfile, path
from statistics import mean, stdev
import xlsxwriter
import datetime
from math import log10


class PrintResults(logging.Handler):
    def __init__(self, text):
        super().__init__()
        self.text = text

    def emit(self, record):
        msg = self.format(record)
        def append_text():
            self.text.configure(state='normal')
            self.text.insert(END, f'{msg}\n')
            self.text.configure(state='disabled')
            self.text.yview(END)
        self.text.after(0, append_text)


class Gui(Frame):
    def __init__(self, master):
        super().__init__()
        self.master = master        
        self.master.title('Viscotester')

        self.grid(column=0, row=0, sticky='ew')  # tentar descobrir o motivo dessa linha
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=1)
        self.grid_columnconfigure(3, weight=1)

        self.main_title = Label(self, text='Viscotester 3.0', font=('Helvetica', 18))
        self.main_title.grid(column=0, row=0, sticky='ew', columnspan=4)

        self.filename_text = Label(self, text='Nome do arquivo')
        self.filename_text.grid(column=0, row=1)
        self.filename = Entry(self, text='Nome do arquivo', width=40)
        self.filename.grid(column=0, row=2)

        self.init_button = Button(self, text="Iniciar análise", width=28)
        self.init_button.grid(column=1, row=1)

        self.stop_analysis = Button(self, text="Parar análise", width=28)
        self.stop_analysis.grid(column=1, row=2)

        self.save_workbook = Button(self, text="Salvar dados", width=28)
        self.save_workbook.grid(column=2, row=1)

        self.launch_workbook = Button(self, text="Abrir planilha", width=28)
        self.launch_workbook.grid(column=2, row=2)
            
        st = ScrolledText.ScrolledText(self, state='disabled')
        st.configure(font='TkFixedFont')
        st.grid(column=0, row=5, sticky='w', columnspan=4)

        text_handler = PrintResults(st)

        logging.basicConfig(filename='test.log',
                            level=logging.INFO,
                            format='%(asctime)s - %(message)s')
        logger = logging.getLogger()
        logger.addHandler(text_handler)


class Viscotester:
    def __init__(self, values):
        self.values = values


def values_storager(self, *values):
    print(values)
    rpm_value, cp_value, torque_value = (values[0], values[1], values[2])

    if rpm_value not in registers.keys():
        registers[rpm_value] = [[cp_value], [torque_value]]
    elif rpm_value in registers.keys():
        registers[rpm_value][0].append(cp_value)
        registers[rpm_value][1].append(torque_value)
    return registers

def job():
    ser = serial.Serial('COM1', 9600)
    while True:
        serial_object = ser.readline().split()
        print(serial_object)
        rpm_value, cp_value, torque_value = (
                                             float(serial_object[3]), 
                                             int(serial_object[7]), 
                                             float(serial_object[5])
                                            )
        values = (rpm_value, cp_value, torque_value)
        print(values)
        print('Entrando na função de armazenamento')
        registers = values_storager(*values)
        print(registers)

        visc = Viscotester(values)

        prints = f'RPM: {rpm_value:.>15} /// cP: {cp_value:.>15} /// Torque: {torque_value:.>15}%'
        logging.info(prints)
        


if __name__ == '__main__':
    registers = dict()

    root = Tk()
    Gui(root)
    
    t1 = threading.Thread(target=job)
    t1.start()
    
    root.mainloop()
    t1.join()