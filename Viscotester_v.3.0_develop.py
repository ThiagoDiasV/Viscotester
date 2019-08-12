from tkinter import *
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

        self.grid(column=0, row=4, sticky='ew')
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=1)
        self.grid_columnconfigure(3, weight=1)

        self.main_title = Label(
                                self,
                                text='Viscotester 6L HAAKE 3.0',
                                font=('Helvetica', 24)
                                )
        self.main_title.grid(column=0, row=0, sticky='ew', columnspan=4)

        self.img = PIL.ImageTk.PhotoImage(PIL.Image.open('Viscotester/panel_photo.jpg'))
        self.panel = Label(self, image=self.img)
        self.panel.grid(column=0, row=1, sticky='ew', columnspan=4)

        self.filename_text = Label(self, text='Nome do arquivo')
        self.filename_text.grid(column=0, row=2)
        self.filename = Entry(self, text='Nome do arquivo', width=40)
        self.filename.grid(column=0, row=3)

        self.init_button = Button(
                                  self,
                                  text="Iniciar análise", width=28,
                                  command=self.initialize_viscotester_readings
                                  )
        self.init_button.grid(column=1, row=2)

        self.stop_analysis = Button(
                                    self,
                                    text="Parar análise",
                                    width=28,
                                    command=self.stop_analysis
                                    )
        self.stop_analysis.grid(column=1, row=3)

        self.save_workbook = Button(
                                    self, 
                                    text="Salvar análise e abrir planilha", 
                                    width=28,
                                    command=self.save_workbook
                                    )
        self.save_workbook.grid(column=2, row=2)

        self.open_another_workbook = Button(
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

        logging.basicConfig(filename='test.log',
                            level=logging.INFO,
                            format='%(asctime)s - %(message)s')
        logger = logging.getLogger()
        logger.addHandler(text_handler)

    def initialize_viscotester_readings(self):
        viscotester.active_status = True

    def stop_analysis(self):
        viscotester.active_status = False

    def save_workbook(self):
        Results_Workbook(self.filename)

    def launch_workbook(self):
        another_workbook = filedialog.askopenfile().name
        startfile(another_workbook)


class Viscotester:

    def __init__(self):
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
                print(self.registers)
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
        self.file_name = gui.filename.get().replace(' ', '')
        self.workbook_name = self.set_workbook_name_path(self.file_name)

        self.workbook = xlsxwriter.Workbook(self.workbook_name)

        registers = viscotester.sort_registers_values()
        self.worksheet_maker(self.workbook, self.file_name, self.workbook_name,
                             **{str(k): v for k, v in registers.items()})
        self.workbook.close()
        startfile(self.workbook_name)

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

        worksheet = workbook.add_worksheet()
        worksheet_name = filename
        bold = workbook.add_format({'bold': True})
        italic = workbook.add_format({'italic': True})
        float_format = workbook.add_format({'num_format': '0.0000'})
        mean_format = workbook.add_format({'num_format': '0.00'})
        percentage_format = workbook.add_format({'num_format': '0.00%'})
        worksheet.set_column(0, 15, 16)
        worksheet.set_column(4, 4, 25)
        worksheet.set_column(9, 9, 20)
        worksheet.write('A1', f'{worksheet_name}', bold)
        worksheet.write('A2', 'Data', italic)
        date_today = self.date_storage()
        worksheet.write('A3',
                        f'{date_today[0]:02d}/{date_today[1]:02d}/'
                        f'{date_today[2]:04d}')
        worksheet.write('B1', 'RPM', bold)
        worksheet.write('C1', 'cP', bold)
        worksheet.write('D1', 'Torque(%)', bold)
        worksheet.write('E1', 'Processamento dos dados >>', bold)
        worksheet.write('F1', 'RPM', bold)
        worksheet.write('G1', 'Médias: cP', bold)
        worksheet.write('H1', 'Desvio padrão: cP', bold)
        worksheet.write('I1', 'DP (%): cP', bold)
        worksheet.write('J1', 'Escala logarítmica >>', bold)
        worksheet.write('K1', 'RPM Log10', bold)
        worksheet.write('L1', 'cP Log10', bold)
        worksheet.write('M1', 'Intercepto', bold)
        worksheet.write('N1', 'Inclinação', bold)
        worksheet.write('O1', 'R²', bold)

        # The for loop below puts the read values inside .xlsx cells.
        # RPM, cP and torque values will be stored on cols 1, 2 and 3.

        row = 1
        col = 1

        for key, value in registers.items():
            for cp in value[0]:
                worksheet.write(row, col, float(key))
                worksheet.write(row, col + 1, cp)
                row += 1
            row -= len(value[0])
            for torque in value[1]:
                worksheet.write(row, col + 2, torque)
                row += 1
        processed_registers = self.data_processor(**registers)

        # The for loop below puts the processed values inside .xlsx cells.
        # RPM, mean(cP), stdev and stdev% will be stored on cols 5, 6, 7 and 8.

        row = col = 1
        for key, value in processed_registers.items():
            if mean(value[0]) != 0:
                worksheet.write(row, col + 4, float(key))
                if len(value[0]) > 1:
                    worksheet.write(row, col + 5, mean(value[0]), mean_format)
                    worksheet.write(row, col + 6, stdev(value[0]), float_format)
                    worksheet.write(row, col + 7,
                                    (stdev(value[0])/(mean(value[0]))),
                                    percentage_format)
                else:
                    worksheet.write(row, col + 5, value[0][0], mean_format)
                    worksheet.write(row, col + 6, 0)
                    worksheet.write(row, col + 7, 0)
                row += 1

        log_list = self.logarithm_values_maker(**processed_registers)

        # write_column() function below puts the log10 values inside .xlsx cells.

        worksheet.write_column('K2', log_list[0], float_format)
        worksheet.write_column('L2', log_list[1], float_format)
        worksheet.write_array_formula(
                                      'M2:M2', '{=INTERCEPT(L2:L20, K2:K20)}',
                                      float_format
                                      )
        worksheet.write_array_formula(
                                      'N2:N2', '{=SLOPE(L2:L20, K2:K20)}',
                                      float_format
                                      )
        worksheet.write_array_formula(
                                      'O2:O2', '{=RSQ(K2:K20, L2:L20)}',
                                      float_format
                                      )

        chart_1 = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
        chart_1.add_series({
            'categories': f'={worksheet_name.replace(" ", "")}'
                          f'!$F2$:$F${len(processed_registers.keys()) + 1}',
            'values': f'={worksheet_name.replace(" ", "")}'
                      f'!$G$2:$G${len(processed_registers.values()) + 1}',
            'line': {'color': 'green'}
            })
        chart_1.set_title({'name': f'{worksheet_name}'})
        chart_1.set_x_axis({
            'name': 'RPM',
            'name_font': {'size': 14, 'bold': True},
        })
        chart_1.set_y_axis({
            'name': 'cP',
            'name_font': {'size': 14, 'bold': True},
        })
        chart_1.set_size({
            'width': 500,
            'height': 400
        })
        worksheet.insert_chart(row + 2, 5, chart_1)

        chart_2 = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
        chart_2.add_series({
            'categories': f'={worksheet_name.replace(" ", "")}'
                          f'!$K$2:$K${len(processed_registers.keys()) + 1}',
            'values': f'={worksheet_name.replace(" ", "")}'
                      f'!$L$2:$L${len(processed_registers.values()) + 1}',
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
        chart_2.set_title({'name': f'Curva escala log: {worksheet_name}'})
        chart_2.set_x_axis({
            'name': 'RPM',
            'name_font': {'size': 14, 'bold': True},
        })
        chart_2.set_y_axis({
            'name': 'cP',
            'name_font': {'size': 14, 'bold': True},
        })
        chart_2.set_size({
            'width': 500,
            'height': 400
        })
        worksheet.insert_chart(row + 2, 10, chart_2)


if __name__ == '__main__':

    root = Tk()
    viscotester = Viscotester()
    gui = Gui(root)
    root.mainloop()
