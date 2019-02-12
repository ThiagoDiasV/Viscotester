import re
from collections import OrderedDict
from os import startfile
from statistics import mean, stdev
from time import sleep
import colorama
from colorama import Fore, Style
import serial
import xlsxwriter
import datetime

colorama.init(autoreset=True, convert=True)


def initial_menu():
    '''
    Prints an initial menu at the screen
    '''

    print(Fore.GREEN + '-' * 90)
    print(Fore.BLUE + '#' * 37 + Fore.CYAN + ' VISCOTESTER 6L ' + Style.RESET_ALL  + Fore.BLUE + '#' * 37)
    print(Fore.BLUE + '#' * 35 + Fore.CYAN + ' INSTRUÇÕES DE USO ' + Style.RESET_ALL + Fore.BLUE + '#' * 36)
    print(Fore.GREEN + '-' * 90)
    print('1 - Ligue o aparelho e realize o ' + Fore.BLUE + 'AUTO TEST' , 'pressionando a tecla ' + Fore.GREEN + 'START')
    print('2 - Observe se não há nenhum fuso acoplado ao aparelho antes de pressionar ' + Fore.GREEN + 'START')
    print('3 - Aguarde o ' + Fore.BLUE + 'AUTO TEST ' +Style.RESET_ALL + 'ser finalizado e em seguida pressione ' + Fore.GREEN + 'START')
    print('4 - Adicione o fuso correto e selecione o fuso correto no aparelho pressionando ' + Fore.YELLOW + 'ENTER')
    print('5 - Selecione a RPM desejada e pressione ' + Fore.YELLOW + 'ENTER')
    print('6 - Observe se o fuso correto está acoplado ao aparelho e pressione ' + Fore.GREEN + 'START')
    print(Fore.GREEN + '-' * 90)
    print(Fore.BLUE + '#' * 90)
    print(Fore.BLUE + '#' * 90)
    print(Fore.GREEN + '-' * 90)


def final_menu():
    '''
    Prints some informations if the maximum torque is obtained from the Viscotester and require the user to press STOP on the equipment
    '''

    print('Torque máximo atingido')
    print('Leituras não são mais possíveis de serem feitas')
    print('Pressione ' + Fore.RED + 'STOP' + Style.RESET_ALL + ' no aparelho')


def regex_name_validation(name):
    '''
    Does a regex validation on sample name and worksheet name
    '''

    regexp = re.compile(r'[\\/|<>*:?"]')
    while regexp.search(name):
        print(Fore.RED + 'Você digitou um caractere não permitido para nome de arquivo.')
        print(Fore.RED + 'Saiba que você não pode usar nenhum dos caracteres abaixo: ')
        print(Fore.RED + ' \ / | < > * : " ?')
        name = str(input('Digite novamente um nome para a amostra sem caracteres proibidos: '))
    return name


def sample_name_function():
    '''
    Require the name of the sample to put on the xlsx filename
    '''

    sample_name = str(input('Digite um nome para o arquivo será gerado: ')).strip()
    sample_name = regex_name_validation(sample_name)
    sleep(2.5)
    print('Aguarde que em instantes o programa se inicializará.')
    sleep(2.5)
    print('Ao finalizar suas leituras, pressione ' + Fore.RED + 'STOP ' + Style.RESET_ALL + 'no aparelho.')
    sleep(2.5)
    print('Ao pressionar ' + Fore.RED + 
          'STOP' + Style.RESET_ALL + 
          ', o programa levará alguns segundos para preparar sua planilha. Aguarde.')
    return sample_name


def serial_object_creator(time_set):
    '''
    At each rotation of the equipment this function creates a serial object 
    This is important because at each rotation the timeout to close serial port should change
    '''

    ser = serial.Serial('COM1', 9600, timeout=time_set)
    serial_object = ser.readline().split()
    return serial_object


def timer_for_closing_port(object):
    '''
    Defines a new time for closing serial port. This times depends on the rotation per minute parameter of equipment
    '''

    if float(object[3]) <= 6:
        time_for_closing = 2*(60/float(object[3]))
    elif float(object[3]) < 100:
        time_for_closing = 3*(60/float(object[3]))
    else:
        time_for_closing = 25*(60/float(object[3]))
    return time_for_closing


def torque_validator(serial_object):
    if serial_object[7] == b'off':
        return False
    else:
        return True


def readings_printer(object):
    '''
    Prints the results of the equipment readings on the screen
    '''

    print(f' RPM: {float(object[3]):.>20} /// cP: {int(object[7]):.>20} /// Torque: {float(object[5]):.>20}%')


def values_storager(object):
    '''
    Storages the readings inside a dict named 'registers' 
    '''

    if float(object[3]) not in registers.keys():
        registers[float(object[3])] = [[int(object[7])], [float(object[5])]]
    elif float(object[3]) in registers.keys():
        registers[float(object[3])][0].append(int(object[7]))
        registers[float(object[3])][1].append(float(object[5]))
    return registers


def data_processor(**registers):
    '''
    Processes the data of registers to delete outliers
    '''

    for value in registers.values():
        if len(value[0]) > 1:    
            mean_value = mean(value[0])
            std_value = stdev(value[0])
            if std_value != 0:
                cp_list = [x for x in value[0] if (x > mean_value - std_value)]
                cp_list = [x for x in cp_list if (x < mean_value + std_value)]
                value[0] = cp_list
            else:
                pass
        else:
            pass
    return registers


def date_storage():
    '''
    A function to create a tuple with the today's date
    '''

    date = datetime.date.today()
    hour = str(datetime.datetime.now().hour)
    minute = str(datetime.datetime.now().minute)
    date_today = (date.day, date.month, date.year, hour, minute)
    return date_today
    

def workbook_maker(sample_name):
    '''
    This function creates a workbook in format .xlsx
    '''

    date_today = date_storage()    
    workbook = xlsxwriter.Workbook(f'{sample_name}_{date_today[0]:02d}{date_today[1]:02d}{date_today[2]:04d}.xlsx')
    return workbook


def worksheet_maker(workbook, **registers):
    '''
    This function creates new worksheets inside the created workbook and put the values in columns
    Inside this function I put a feature to put the processed data in other columns
    After that, this function creates a chart with the processed data
    '''

    worksheet_name = str(input('Digite o nome da amostra: ')).strip()
    worksheet_name = regex_name_validation(worksheet_name)
    worksheet = workbook.add_worksheet(f'{worksheet_name}')
    bold = workbook.add_format({'bold': True})
    italic = workbook.add_format({'italic': True})
    worksheet.set_column(0, 8, 20)
    worksheet.set_column(4, 4, 25)
    worksheet.write('A1', f'{worksheet_name}', bold)
    worksheet.write('A2', 'Data', italic)
    date_today = date_storage() 
    worksheet.write('A3', f'{date_today[0]:02d}/{date_today[1]:02d}/{date_today[2]:04d}')
    worksheet.write('B1', 'RPM', bold)
    worksheet.write('C1', 'cP', bold)
    worksheet.write('D1', 'Torque(%)', bold)
    worksheet.write('E1', 'Processamento dos dados >>', bold)
    worksheet.write('G1', 'RPM', bold)
    worksheet.write('H1', 'cP', bold)
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
    row = col = 1
    for key, value in processed_registers.items():
        if mean(value[0]) != 0:    
            worksheet.write(row, col + 5, float(key))
            if len(value[0]) > 1:
                worksheet.write(row, col + 6, mean(value[0]))
            else:
                worksheet.write(row, col + 6, value[0][0]) 
            row += 1
    chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
    chart.add_series({
        'categories': f'={worksheet_name}!$G2$:$G${len(processed_registers.keys()) + 1}', 
        'values': f'={worksheet_name}!$H$2:$H${len(processed_registers.values()) + 1}', 
        'line': {'color': 'green'}
        })
    chart.set_title({'name': f'{worksheet_name}'})
    chart.set_x_axis({
        'name': 'RPM',
        'name_font': {'size': 14, 'bold': True},
    })
    chart.set_y_axis({
        'name': 'cP',
        'name_font': {'size': 14, 'bold': True},
    })
    worksheet.insert_chart('F18', chart)


def workbook_close_function(workbook):
    '''
    A simple function to close the created workbook
    '''

    workbook.close()


def workbook_launcher(workbook):
    '''
    A simple function to launch the workbook for user see his results
    '''

    date_today = date_storage()    
    startfile(f'{sample_name}_{date_today[0]:02d}{date_today[1]:02d}{date_today[2]:04d}.xlsx')   


initial_menu()
sample_name = sample_name_function() 
workbook = workbook_maker(sample_name)
repeat_option = ''
while repeat_option != 'N':
    registers = dict()  # Registros das leituras serão armazenados neste dicionário.
    time = 250  # Tempo para timeout da porta inicial alto para evitar bugs na inicialização do programa.
    sleep(5)  # Tempo de espera para evitar que bugs que o aparelho gera na inicialização possam dar crash no programa.
    print(Fore.GREEN + '-' * 90)
    print(Fore.BLUE + '#' * 40 + Fore.CYAN + ' LEITURAS ' + Fore.BLUE + '#' * 40)
    print(Fore.GREEN + '-' * 90)
    while True:
        try:
            object = serial_object_creator(time)
            time = timer_for_closing_port(object)
            if torque_validator(object):
                if object == False:
                    print('Torque máximo atingido ou erro no aparelho')
                else:
                    readings_printer(object)
                    registers = values_storager(object)
            else:
                final_menu()

        except KeyboardInterrupt:
            print('Programa interrompido por atalho de teclado')
            break

        except IndexError:
            print('Foi pressionado ' + Fore.RED + 'STOP' + Style.RESET_ALL + ' no aparelho')
            registers = dict(OrderedDict(sorted(registers.items())))
            break


    worksheet_maker(workbook, **{str(k): v for k , v in registers.items()})
    print('Você quer ler outra amostra?')
    print('Responda com "S" para se sim ou "N" para se não')
    print('Se você quiser ler outra amostra,\nresponda após pressionar ' + Fore.GREEN + 'START' + Style.RESET_ALL +  ' no aparelho')
    repeat_option = str(input('[S/N]: ')).strip().upper()[0]
workbook_close_function(workbook)
workbook_launcher(workbook) 
print(Fore.GREEN + 'OBRIGADO POR USAR O VISCOTESTER 6L SCRIPT')
