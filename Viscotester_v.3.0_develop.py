"""
Viscotester: a Python script to process data from a viscosimeter
Visco Tester 6L Haake.
The documentation is in English but the program is used in a Brazilian
laboratory, so the language of the prints is Portuguese-BR.
This program is made specifically for Visco Tester 6L Haake and Windows OS.
A viscosimeter is a equipment used to measure the viscosity of liquids and
fluids. The equipment use tools named spindles. The spindle is immersed
in the substance that will be evaluated and is rotated at different
rotations.
The output of equipment are the rotation per minute (RPM) parameter,
the viscosity (cP) and the torque (%) value. The torque value is calculated
based on the speed and the geometry of the spindle.
"""

import tkinter as tk
import tkinter.scrolledtext as ScrolledText
from tkinter import filedialog
import threading
import logging
import serial
import re
from collections import OrderedDict
from os import startfile, path
from statistics import mean, stdev
import xlsxwriter
import datetime
from math import log10
import PIL.Image
import PIL.ImageTk


class PrintResults(logging.Handler):

    def __init__(self, text):
        """
        Initializes the class that prints the results
        on the tkinter GUI screen.
        """
        super().__init__()
        self.text = text

    def emit(self, record):
        """
        Records the message on variable 'msg'.
        """
        msg = self.format(record)

        def append_text():
            self.text.configure(state='normal')
            self.text.insert(tk.END, f'{msg}\n')
            self.text.configure(state='disabled')
            self.text.yview(tk.END)
        self.text.after(0, append_text)


class Gui(tk.Frame):

    def __init__(self, master):
        """
        Creates the GUI of Viscotester.
        """
        super().__init__()
        self.master = master
        self.master.title('Viscotester')

        self.grid(column=0, row=4, sticky='ew')
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=1)
        self.grid_columnconfigure(3, weight=1)

        self.main_title = tk.Label(
                                self,
                                text='Viscotester 6L HAAKE 3.0',
                                font=('Helvetica', 24)
                                )
        self.main_title.grid(column=0, row=0, sticky='ew', columnspan=4)

        self.img = PIL.ImageTk.PhotoImage(PIL.Image.open(
                            'Viscotester/panel_photo.jpg')
                            )
        self.panel = tk.Label(self, image=self.img)
        self.panel.grid(column=0, row=1, sticky='ew', columnspan=4)

        self.filename_text = tk.Label(self, text='Nome do arquivo')
        self.filename_text.grid(column=0, row=2)
        self.filename = tk.Entry(self, text='Nome do arquivo', width=40)
        self.filename.grid(column=0, row=3)

        self.init_button = tk.Button(
                                  self,
                                  text="Iniciar análise", width=28,
                                  command=self.initialize_viscotester_readings
                                  )
        self.init_button.grid(column=1, row=2)

        self.stop_analysis = tk.Button(
                                    self,
                                    text="Parar análise",
                                    width=28,
                                    command=self.stop_analysis
                                    )
        self.stop_analysis.grid(column=1, row=3)

        self.save_workbook = tk.Button(
                                    self, 
                                    text="Salvar análise e abrir planilha", 
                                    width=28,
                                    command=self.save_workbook
                                    )
        self.save_workbook.grid(column=2, row=2)

        self.open_another_workbook = tk.Button(
                                      self,
                                      text="Abrir outra planilha",
                                      width=28,
                                      command=self.launch_workbook
                                      )
        self.open_another_workbook.grid(column=2, row=3)

        st = ScrolledText.ScrolledText(self, state='disabled')
        st.configure(font='TkFixedFont')
        st.grid(column=0, row=5, sticky='w', columnspan=4)

        text_handler = PrintResults(st)

        # Change this logic
        logging.basicConfig(filename=f'Viscotester/{self.filename.get()}.log',
                            level=logging.INFO,
                            format='%(asctime)s - %(message)s')
        logger = logging.getLogger()
        logger.addHandler(text_handler)

    def initialize_viscotester_readings(self):
        """
        Change the active_status property of an instance of
        Viscotester class to True, initializing the readings
        and the recording of data.
        """
        viscotester.active_status = True

    def stop_analysis(self):
        """
        Change the active_status property of an instance of
        Viscotester class to False, stopping the readings
        and the recording of data.
        """
        viscotester.active_status = False

    def save_workbook(self):
        """
        Creates an instance of Results_Workbook class,
        which will create a workbook to the recorded data.
        """
        Results_Workbook(self.filename)

    def launch_workbook(self):
        """
        Allows the user to open any workbook on his computer.
        """
        another_workbook = filedialog.askopenfile().name
        startfile(another_workbook)


class Viscotester:

    def __init__(self):
        """
        Initialize an instance of Viscotester class, but the
        default active_status is False. The readings will only
        appear on screen of the GUI when the active_status will
        set to True.
        """
        self._thread = threading.Thread(target=self.job)
        self._stop_threads = threading.Event()
        self._active_status = False
        self._registers = dict()

    @property
    def active_status(self):
        return self._active_status

    @active_status.setter
    def active_status(self, status):
        if self._active_status != status:
            self._active_status = status
            if self._active_status:
                self._thread.start()

    @property
    def registers(self):
        return self._registers

    def job(self):

        def validator(serial_object):
            if len(serial_object) == 8:
                return True

        def values_storager(*args):
            rpm, cp, torque = (args[0], args[1], args[2])

            if rpm not in self._registers.keys():
                self._registers[rpm] = [[cp], [torque]]
            elif rpm in self._registers.keys():
                self._registers[rpm][0].append(cp)
                self._registers[rpm][1].append(torque)
            return self._registers

        self.ser = serial.Serial('COM1', 9600)

        while self.active_status:
            self.serial_object = self.ser.readline().split()
            if validator(self.serial_object):
                rpm, cp, torque = (
                                    float(self.serial_object[3]),
                                    int(self.serial_object[7]),
                                    float(self.serial_object[5])
                                    )
                values = (rpm, cp, torque)
                self._registers = values_storager(*values)
                prints = (f'RPM: {rpm:.>15} /// '
                            f'cP: {cp:.>15} /// '
                            f'Torque: {torque:.>15}%')
            else:
                prints = f'ERRO!'
            logging.info(prints)

    def sort_registers_values(self):
        self._registers = dict(
            OrderedDict(sorted(viscotester.registers.items()))
            )
        return self._registers


class Results_Workbook:
    def __init__(self, filename):
        self.file_name = self.check_filename(
                        gui.filename.get().replace(' ', '')
                        )
        self.workbook_name = self.set_workbook_name_path(self.file_name)

        self.workbook = xlsxwriter.Workbook(self.workbook_name)

        registers = viscotester.sort_registers_values()
        self.worksheet_maker(self.workbook, self.file_name, self.workbook_name,
                             **{str(k): v for k, v in registers.items()})
        self.workbook.close()
        startfile(self.workbook_name)

    def check_filename(self, filename):
        regexp = re.compile(r'[\\/|<>*:?\"[\]]')
        self.file_name = re.sub(regexp, '', filename)
        return self.file_name

    def date_storage(self):
        '''
        A function to create a tuple with the today's date.
        This date will be in one cell of the workbook that
        will be created and in the name of the xlsx file.
        '''

        date = datetime.date.today()
        date_today = (date.day, date.month, date.year)
        return date_today

    def set_workbook_name_path(self, filename):
        date_today = self.date_storage()
        day, month, year = (date_today[0],
                            date_today[1],
                            date_today[2])

        if path.isdir('C:/Users/UFC/Desktop/Viscosidades/'):
            my_path = 'C:/Users/UFC/Desktop/Viscosidades/'
        elif path.isdir('C:/Users/UFC/Desktop/'):
            my_path = 'C:/Users/UFC/Desktop/'
        else:
            my_path = ''
        workbook_name = (
            f'{my_path}'
            f'{filename}_{day:02d}'
            f'{month:02d}{year:04d}'
            '.xlsx'
        )
        return workbook_name

    def data_processor(self, **registers):
        '''
        Processes the data of registers dict to delete outliers.
        The cutoff parameter are (average - standard deviation) and
        (average + standard deviation).
        A for loop perform iteration on values of registers dict and exclude
        outliers.
        '''

        for value in registers.values():
            if len(value[0]) > 1:
                mean_value = mean(value[0])
                std_value = stdev(value[0])
                if std_value != 0:
                    cp_list = [x for x in value[0]
                               if (x > mean_value - std_value)]
                    cp_list = [x for x in cp_list
                               if (x < mean_value + std_value)]
                    value[0] = cp_list
        return registers

    def logarithm_values_maker(self, **registers):
        '''
        Calculates the base-10 logarithm of the processed values.
        The dict comprehension below is only to transform RPM values
        in float types again, because the **kwargs only accept string
        type as keys, and is necessary that RPM values are float type,
        not string.
        A new list (cp_list) is created to receive the cP values.
        A iteration is made on keys of registers dict using for loop
        to make a list with two lists inside of it. The first list
        will store the base-10 logarithm values of RPM values. The second
        list will store the base-10 logarithm values of cP values.
        This function returns this logarithm_list.
        '''

        registers = {float(k): v for k, v in registers.items()}
        cp_list = list()
        for value in registers.values():
            cp_list.append(mean(value[0]))
        for key in registers.keys():
            logarithm_list = [[log10(k) for k in registers.keys()
                               if mean(registers[k][0]) != 0],
                              [log10(v) for v in cp_list if v != 0]]
        return logarithm_list

    def worksheet_maker(self, workbook, filename, workbook_name, **registers):
        '''
        This function creates new worksheets inside the created workbook and put
        the values in columns.
        In each worksheet:
        Column 'A' will store the sample name and the date.
        Columns 'B', 'C', and 'D' will store all read data (RPM, cP and Torque
        values).
        Columns 'F', 'G', 'H', and 'I' will store the processed data, without
        outliers, respectively: RPM, average cP, standard deviation and relative
        standard deviation.
        Columns 'K' and 'L' will receive log10 values of processed RPM and cP
        values.
        Finally, in columns 'M', 'N' and 'O', the cells 'M2', 'N2' and 'O2' will
        receive intercept, slope and R squared values of log10 values.
        Each worksheet will have two charts, one for processed data and other for
        log10 data.
        '''

        self.worksheet_name = filename[:30]
        self.worksheet = workbook.add_worksheet(self.worksheet_name)
        self.bold = workbook.add_format({'bold': True})
        self.italic = workbook.add_format({'italic': True})
        self.float_format = workbook.add_format({'num_format': '0.0000'})
        self.mean_format = workbook.add_format({'num_format': '0.00'})
        self.percentage_format = workbook.add_format({'num_format': '0.00%'})
        self.worksheet.set_column(0, 15, 16)
        self.worksheet.set_column(4, 4, 25)
        self.worksheet.set_column(9, 9, 20)
        self.worksheet.write('A1', f'{self.worksheet_name}', self.bold)
        self.worksheet.write('A2', 'Data', self.italic)
        self.date_today = self.date_storage()
        self.worksheet.write('A3',
                        f'{self.date_today[0]:02d}/{self.date_today[1]:02d}/'
                        f'{self.date_today[2]:04d}')
        self.worksheet.write('B1', 'RPM', self.bold)
        self.worksheet.write('C1', 'cP', self.bold)
        self.worksheet.write('D1', 'Torque(%)', self.bold)
        self.worksheet.write('E1', 'Processamento dos dados >>', self.bold)
        self.worksheet.write('F1', 'RPM', self.bold)
        self.worksheet.write('G1', 'Médias: cP', self.bold)
        self.worksheet.write('H1', 'Desvio padrão: cP', self.bold)
        self.worksheet.write('I1', 'DP (%): cP', self.bold)
        self.worksheet.write('J1', 'Escala logarítmica >>', self.bold)
        self.worksheet.write('K1', 'RPM Log10', self.bold)
        self.worksheet.write('L1', 'cP Log10', self.bold)
        self.worksheet.write('M1', 'Intercepto', self.bold)
        self.worksheet.write('N1', 'Inclinação', self.bold)
        self.worksheet.write('O1', 'R²', self.bold)

        # The for loop below puts the read values inside .xlsx cells.
        # RPM, cP and torque values will be stored on cols 1, 2 and 3.

        self.row = 1
        self.col = 1

        for key, value in registers.items():
            for cp in value[0]:
                self.worksheet.write(self.row, self.col, float(key))
                self.worksheet.write(self.row, self.col + 1, cp)
                self.row += 1
            self.row -= len(value[0])
            for torque in value[1]:
                self.worksheet.write(self.row, self.col + 2, torque)
                self.row += 1
        self.processed_registers = self.data_processor(**registers)

        # The for loop below puts the processed values inside .xlsx cells.
        # RPM, mean(cP), stdev and stdev% will be stored on cols 5, 6, 7 and 8.

        self.row = self.col = 1
        for key, value in self.processed_registers.items():
            if mean(value[0]) != 0:
                self.worksheet.write(self.row, self.col + 4, float(key))
                if len(value[0]) > 1:
                    self.worksheet.write(self.row, self.col + 5, mean(value[0]), self.mean_format)
                    self.worksheet.write(self.row, self.col + 6, stdev(value[0]), self.float_format)
                    self.worksheet.write(self.row, self.col + 7,
                                    (stdev(value[0])/(mean(value[0]))),
                                    self.percentage_format)
                else:
                    self.worksheet.write(self.row, self.col + 5, value[0][0], self.mean_format)
                    self.worksheet.write(self.row, self.col + 6, 0)
                    self.worksheet.write(self.row, self.col + 7, 0)
                self.row += 1

        self.log_list = self.logarithm_values_maker(**self.processed_registers)

        # write_column() function below puts the log10 values inside .xlsx cells.

        self.worksheet.write_column('K2', self.log_list[0], self.float_format)
        self.worksheet.write_column('L2', self.log_list[1], self.float_format)
        self.worksheet.write_array_formula(
                                      'M2:M2', '{=INTERCEPT(L2:L20, K2:K20)}',
                                      self.float_format
                                      )
        self.worksheet.write_array_formula(
                                      'N2:N2', '{=SLOPE(L2:L20, K2:K20)}',
                                      self.float_format
                                      )
        self.worksheet.write_array_formula(
                                      'O2:O2', '{=RSQ(K2:K20, L2:L20)}',
                                      self.float_format
                                      )

        self.chart_1 = self.workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
        self.chart_1.add_series({
            'categories': f'={self.worksheet_name.replace(" ", "")}'
                          f'!$F2$:$F${len(self.processed_registers.keys()) + 1}',
            'values': f'={self.worksheet_name.replace(" ", "")}'
                      f'!$G$2:$G${len(self.processed_registers.values()) + 1}',
            'line': {'color': 'green'}
            })
        self.chart_1.set_title({'name': f'{self.worksheet_name}'})
        self.chart_1.set_x_axis({
            'name': 'RPM',
            'name_font': {'size': 14, 'bold': True},
        })
        self.chart_1.set_y_axis({
            'name': 'cP',
            'name_font': {'size': 14, 'bold': True},
        })
        self.chart_1.set_size({
            'width': 500,
            'height': 400
        })
        self.worksheet.insert_chart(self.row + 2, 5, self.chart_1)

        self.chart_2 = self.workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
        self.chart_2.add_series({
            'categories': f'={self.worksheet_name.replace(" ", "")}'
                          f'!$K$2:$K${len(self.processed_registers.keys()) + 1}',
            'values': f'={self.worksheet_name.replace(" ", "")}'
                      f'!$L$2:$L${len(self.processed_registers.values()) + 1}',
            'line': {'color': 'blue'},
            'trendline': {
                'type': 'linear',
                'display_equation': True,
                'display_r_squared': True,
                'line': {
                    'color': 'red',
                    'width': 1,
                    'dash_type': 'long_dash',
                },
            },
        })
        self.chart_2.set_title({'name': f'Curva escala log: {self.worksheet_name}'})
        self.chart_2.set_x_axis({
            'name': 'RPM',
            'name_font': {'size': 14, 'bold': True},
        })
        self.chart_2.set_y_axis({
            'name': 'cP',
            'name_font': {'size': 14, 'bold': True},
        })
        self.chart_2.set_size({
            'width': 500,
            'height': 400
        })
        self.worksheet.insert_chart(self.row + 2, 10, self.chart_2)


if __name__ == '__main__':

    root = tk.Tk()
    root.resizable(0, 0)
    viscotester = Viscotester()
    gui = Gui(root)
    root.mainloop()
