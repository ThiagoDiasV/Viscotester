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

        self.grid(column=0, row=0, sticky='ew')
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=1)
        self.grid_columnconfigure(3, weight=1)

        self.main_title = Label(
                                self,
                                text='Viscotester 3.0',
                                font=('Helvetica', 18)
                                )
        self.main_title.grid(column=0, row=0, sticky='ew', columnspan=4)

        self.filename_text = Label(self, text='Nome do arquivo')
        self.filename_text.grid(column=0, row=1)
        self.filename = Entry(self, text='Nome do arquivo', width=40)
        self.filename.grid(column=0, row=2)

        self.init_button = Button(
                                  self,
                                  text="Iniciar análise", width=28,
                                  command=self.initialize_viscotester
                                  )
        self.init_button.grid(column=1, row=1)

        self.stop_analysis = Button(
                                    self,
                                    text="Parar análise",
                                    width=28,
                                    command=lambda: visc.change_active_status()
                                    )
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

    def initialize_viscotester(self):
        visc = Viscotester()
        visc.active_status = True
        return visc


class Viscotester:

    def __init__(self):
        self.thread = threading.Thread(target=self.job)
        self.thread.start()
        self.stop_threads = threading.Event()
        self.active_status = False

    @property
    def active_status(self):
        return active_status

    @active_status.setter
    def active_status(self, value):
        self.active_status = value

    def change_active_status(self):
        if self.active_status:
            self.active_status = False
        else:
            self.active_status = True

    def active_status(self, value):
        self.active_status = value

    def job(self):

        def validator(serial_object):
            if len(serial_object) == 8:
                return True

        def values_storager(*args):
            rpm_value, cp_value, torque_value = (args[0], args[1], args[2])

            if rpm_value not in registers.keys():
                registers[rpm_value] = [[cp_value], [torque_value]]
            elif rpm_value in registers.keys():
                registers[rpm_value][0].append(cp_value)
                registers[rpm_value][1].append(torque_value)
            return registers

        registers = dict()
        ser = serial.Serial('COM1', 9600)

        while True:
            print(self.active_status)
            if self.active_status:
                serial_object = ser.readline().split()
                if validator(serial_object):
                    rpm, cp, torque = (
                                        float(serial_object[3]),
                                        int(serial_object[7]),
                                        float(serial_object[5])
                                      )
                    values = (rpm, cp, torque)
                    registers = values_storager(*values)
                    prints = (f'RPM: {rpm:.>15} /// '
                              f'cP: {cp:.>15} /// '
                              f'Torque: {torque:.>15}%')
                    print(registers)
                else:
                    prints = f'ERRO!'
                logging.info(prints)



if __name__ == '__main__':

    root = Tk()
    gui = Gui(root)
    root.mainloop()
