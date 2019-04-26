'''
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
'''

# Imports
import re
from collections import OrderedDict
from os import startfile, path
from statistics import mean, stdev
from time import sleep
import colorama
from colorama import Fore, Style
import serial
import xlsxwriter
import datetime
from math import log10


colorama.init(autoreset=True, convert=True)


def initial_menu():
    '''
    Prints an initial menu at the screen.
    '''

    print(Fore.GREEN + '-' * 90)
    print(Fore.BLUE + '#' * 37 + Fore.CYAN + ' VISCOTESTER 6L '
          + Style.RESET_ALL + Fore.BLUE + '#' * 37)
    print(Fore.BLUE + '#' * 35 + Fore.CYAN + ' INSTRUÇÕES DE USO '
          + Style.RESET_ALL + Fore.BLUE + '#' * 36)
    print(Fore.GREEN + '-' * 90)
    print('1 - Ligue o aparelho e realize o ' + Fore.BLUE + 'AUTO TEST',
          'pressionando a tecla ' + Fore.GREEN + 'START')
    print('2 - Observe se não há nenhum fuso acoplado ao aparelho antes de '
          'pressionar ' + Fore.GREEN + 'START')
    print('3 - Aguarde o ' + Fore.BLUE + 'AUTO TEST ' + Style.RESET_ALL +
          'ser finalizado e em seguida pressione ' + Fore.GREEN + 'START')
    print('4 - Adicione o fuso correto e selecione o fuso correto no aparelho '
          'pressionando ' + Fore.YELLOW + 'ENTER')
    print('5 - Selecione a RPM desejada e pressione ' + Fore.YELLOW + 'ENTER')
    print('6 - Observe se o fuso correto está acoplado ao aparelho e '
          'pressione ' + Fore.GREEN + 'START')
    print(Fore.GREEN + '-' * 90)
    print(Fore.BLUE + '#' * 90)
    print(Fore.BLUE + '#' * 90)
    print(Fore.GREEN + '-' * 90)


def final_menu():
    '''
    Prints some informations if the maximum torque is obtained from the
    Viscotester and require the user to press STOP on the equipment.
    '''

    print('Torque máximo atingido')
    print('Leituras não são mais possíveis de serem feitas')
    print('Pressione ' + Fore.RED + 'STOP' + Style.RESET_ALL +
          ' no aparelho e ' + Fore.GREEN + 'aguarde')


def regex_name_validation(name):
    '''
    Does a validation on sample name and worksheet name using regex
    to avoid errors on the file that will be created.
    The input is the name that the user typed to the program.
    The function repeats the requirement of the name if the user used
    forbidden characters (like \\/|<>*:?").
    Returns the name that will be used.
    '''

    regexp = re.compile(r'[\\/|<>*:?\"[\]]')
    while regexp.search(name):
        print(Fore.RED + 'Você digitou um caractere não permitido '
              'para nome de arquivo ou de planilha.')
        print(Fore.RED + 'Saiba que você não pode usar nenhum dos '
              'caracteres abaixo: ')
        print(Fore.RED + r' [ ] \  / | < > * : " ?')
        name = str(input('Digite novamente um nome para a amostra '
                         'sem caracteres proibidos: '))
    return name


def file_name_function():
    '''
    Require the name of the sample to put on the xlsx filename.
    The regex_name_validation() function is used here to avoid errors.
    '''

    file_name = str(input('Digite um nome para o arquivo (.xlsx) '
                          'que será gerado: ')).strip()
    file_name = regex_name_validation(file_name)
    return file_name


def serial_object_creator(time_set):
    '''
    At each rotation of the equipment this function creates a serial object.
    This is important because at each rotation the timeout to close serial
    port should change. This occurs because the time to break the while loop
    is dependent of the rotation of equipment.
    The time to closing port responsibility is assigned to 'time_set'
    variable.
    The data of serial port will be assigned to 'ser' variable.
    The class serial.Serial receive 'COM1' as port parameter because this
    program is used on Windows OS. Baudrate parameter is 9600 and timeout
    parameter is equal to 'time_set' variable. The variable 'time_set' is
    defined in timer_for_closing_port() function below.
    Of 'serial_object', index [3] is the RPM value, index [5] is the torque
    value and the index [7] is the viscosity (cP) value.
    '''

    ser = serial.Serial('COM1', 9600, timeout=time_set)
    serial_object = ser.readline().split()

    return serial_object


def timer_for_closing_port(serial_object):
    '''
    Defines a new time for closing serial port. This times depends on the
    rotation per minute parameter of equipment.
    The possible values for rotation per minute parameter of the equipment
    are: 0.3, 0.5, 0.6, 1, 1.5, 2, 2.5, 3, 4, 5, 6, 10, 12, 20, 30, 50, 60,
    100 and 200 RPMs.
    When the rotation per minute (RPM) parameter of equipment is lower than
    6 RPMs, the 'time_for_closing' value is defined by the 'if' statement
    below.
    If the value of RPM is above 6 and below 100, 'time_for_closing' value
    is defined by the 'elif' statement. Finally, if the RPM value is 100
    or 200 RPMs, 'time_for_closing' value is defined by 'else' statement.
    These differences on calculation of 'time_for_closing' variable occurs
    because this variable is responsible to finish the loop that controls
    the program, and at high rotations the probability of errors increase.
    The 'float(object[3])' value below is the RPM parameter. 'float' function
    is necessary because the equipment send to the computer bytes literals.
    '''

    rpm_value = float(serial_object[3])

    if rpm_value <= 6:
        time_for_closing = 2.5*(60/rpm_value)
    elif rpm_value < 100:
        time_for_closing = 3.5*(60/rpm_value)
    else:
        time_for_closing = 25*(60/rpm_value)
    return time_for_closing


def torque_validator(serial_object):
    '''
    Returns a boolean value that depends on the torque of equipment.
    '''

    cp_value = serial_object[7]

    if cp_value == b'off':
        return False
    else:
        return True


def readings_printer(serial_object):
    '''
    Prints the results of the equipment readings at the screen.
    As said before, the indexes 3, 5 and 7 represents the RPM
    values, the torque values and the cP values respectively.
    '''

    rpm_value, cp_value, torque_value = (
                                         float(serial_object[3]), 
                                         int(serial_object[7]), 
                                         float(serial_object[5])
                                         )

    print(f' RPM: {rpm_value:.>20} /// cP: {cp_value:.>20} '
          f'/// Torque: {torque_value:.>20}%')


def values_storager(serial_object):
    '''
    Storages the readings inside a dict named 'registers'.
    The keys are the RPM values. The values are two lists, the first
    list receives the cP values and the second list receives the torque
    values. Each key have two lists representing cP and torque values.
    The 'object' parameter is the serial_object of serial_object_creator()
    function.
    The return is the dict registers with new values.
    '''

    rpm_value, cp_value, torque_value = (
                                         float(serial_object[3]), 
                                         int(serial_object[7]), 
                                         float(serial_object[5])
                                         )

    if rpm_value not in registers.keys():
        registers[rpm_value] = [[cp_value], [torque_value]]
    elif rpm_value in registers.keys():
        registers[rpm_value][0].append(cp_value)
        registers[rpm_value][1].append(torque_value)
    return registers


def data_processor(**registers):
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
                cp_list = [x for x in value[0] if (x > mean_value - std_value)]
                cp_list = [x for x in cp_list if (x < mean_value + std_value)]
                value[0] = cp_list
    return registers


def logarithm_values_maker(**registers):
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


def date_storage():
    '''
    A function to create a tuple with the today's date.
    This date will be in one cell of the workbook that
    will be created and in the name of the xlsx file.
    '''

    date = datetime.date.today()
    date_today = (date.day, date.month, date.year)
    return date_today


def workbook_maker(file_name):
    '''
    This function creates a workbook in format .xlsx. and returns it.
    The else statement below is because if some user delete the folder
    'Viscosidades', the workbooks will be saved on Desktop.
    '''

    date_today = date_storage()
    if path.isdir('C:/Users/UFC/Desktop/Viscosidades/'):
        workbook = xlsxwriter.Workbook(
                                'C:/Users/UFC/Desktop/Viscosidades/'
                                f'{file_name}_{date_today[0]:02d}'
                                f'{date_today[1]:02d}{date_today[2]:04d}'
                                '.xlsx')
    else:
        workbook = xlsxwriter.Workbook(
                                'C:/Users/UFC/Desktop/'
                                f'{file_name}_{date_today[0]:02d}'
                                f'{date_today[1]:02d}{date_today[2]:04d}'
                                '.xlsx')
    return workbook


def worksheet_name_function():
    '''
    This function records the name of each worksheet using the name of the
    sample evaluated.
    '''
    sample_name = str(input('Digite o nome da amostra: ')).strip()
    sample_name = regex_name_validation(sample_name)
    return sample_name


def worksheet_maker(workbook, worksheet_name, **registers):
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

    worksheet = workbook.add_worksheet(f'{worksheet_name.replace(" ", "")}')
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
    date_today = date_storage()
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
    processed_registers = data_processor(**registers)

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

    log_list = logarithm_values_maker(**processed_registers)

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


def workbook_close_function(workbook):
    '''
    A simple function to close the created workbook.
    '''

    workbook.close()


def workbook_launcher(workbook):
    '''
    A simple function to launch the workbook for user to see his results.
    '''

    date_today = date_storage()
    if path.isdir('C:/Users/UFC/Desktop/Viscosidades/'):
        startfile('C:/Users/UFC/Desktop/Viscosidades/'
                  f'{file_name}_{date_today[0]:02d}'
                  f'{date_today[1]:02d}{date_today[2]:04d}'
                  '.xlsx')
    else:
        startfile('C:/Users/UFC/Desktop/'
                  f'{file_name}_{date_today[0]:02d}'
                  f'{date_today[1]:02d}{date_today[2]:04d}'
                  '.xlsx')


# Init.
initial_menu()
file_name = file_name_function()
workbook = workbook_maker(file_name)
repeat_option = ''
regex_repeat = re.compile(r'[NS]')
while repeat_option != 'N':
    repeat_option = ''
    worksheet_name = worksheet_name_function()
    sleep(2.5)
    print('Aguarde que em instantes o programa se inicializará.')
    sleep(2.5)
    print('Ao finalizar suas leituras, pressione ' + Fore.RED + 'STOP '
          + Style.RESET_ALL + 'no aparelho.')
    sleep(2.5)
    print('Ao pressionar ' + Fore.RED +
          'STOP' + Style.RESET_ALL +
          ', o programa levará alguns segundos para preparar sua planilha. '
          'Aguarde.')
    registers = dict()  # The registered values will be stored in this dict.
    time = 300  # First timeout value. Will change after the first rotation.
    sleep(5)  # Delay the beginning of the script. This helps to avoid errors.
    print(Fore.GREEN + '-' * 90)
    print(Fore.BLUE + '#' * 40 + Fore.CYAN + ' LEITURAS '
          + Fore.BLUE + '#' * 40)
    print(Fore.GREEN + '-' * 90)
    while True:
        try:
            object = serial_object_creator(time)
            time = timer_for_closing_port(object)
            if torque_validator(object):
                if not object:
                    print('Torque máximo atingido ou erro no aparelho')
                else:
                    readings_printer(object)
                    registers = values_storager(object)
            else:
                final_menu()

        except KeyboardInterrupt:
            print('Programa interrompido por atalho de teclado')
            break

        except IndexError:  # This exception finishes the loop.
            print('Foi pressionado ' + Fore.RED + 'STOP'
                  + Style.RESET_ALL + ' no aparelho')
            registers = dict(OrderedDict(sorted(registers.items())))
            break

    worksheet_maker(
                    workbook, worksheet_name,
                    **{str(k): v for k, v in registers.items()}
                    )
    print('Você quer ler outra amostra?')
    print('Responda com "S" para se sim ou "N" para se não.')
    print('Se você quiser ler outra amostra, coloque a nova amostra, retire e limpe o fuso e,')
    print('após isso, responda abaixo após pressionar '
          + Fore.GREEN + 'START' + Style.RESET_ALL + ' no aparelho:')
    while not regex_repeat.search(repeat_option):
        repeat_option = str(input('[S/N]: ')).strip().upper()
        if repeat_option == 'S':
            print('Pressione ' + Fore.GREEN + 'START')
            sleep(5)


workbook_close_function(workbook)
workbook_launcher(workbook)
print(Fore.GREEN + 'OBRIGADO POR USAR O VISCOTESTER 6L SCRIPT')
