import re
from collections import OrderedDict
from os import startfile
from statistics import mean, stdev
from time import sleep
import serial
import xlsxwriter


def initial_menu():
    print('-' * 90)
    print('#' * 37, 'VISCOTESTER 6L', '#' * 37)
    print('#' * 35, 'INSTRUÇÕES DE USO', '#' * 36)
    print('-' * 90)
    print('1 - Ligue o aparelho, realize o AUTO TEST pressionando a tecla START;')
    print('2 - Observe se não há nenhum fuso acoplado ao aparelho, se sim, pressione START;')
    print('3 - Aguarde o AUTO TEST ser finalizado;')
    print('4 - Adicione o fuso correto, selecione o fuso correto no aparelho pressionando ENTER;')
    print('5 - Selecione a RPM desejada e pressione ENTER;')
    print('6 - Observe se o fuso correto está acoplado ao aparelho e pressione START.')
    print('-' * 90)
    print('#' * 90)
    print('#' * 90)
    print('-' * 90)


def sample_name():
    regexp = re.compile(r'[\\/|<>*:?"]')
    sample_name = str(input('Digite o nome da amostra: ')).strip()
    while regexp.search(sample_name):
        print('Você digitou um caractere não permitido para nome de arquivo.')
        print('Saiba que você não pode usar nenhum dos caracteres abaixo: ')
        print(' \ / | < > * : " ?')
        sample_name = str(input('Digite novamente um nome para a amostra sem caracteres proibidos: '))
    sleep(2.5)
    print('Aguarde que em instantes o programa se inicializará.')
    sleep(2.5)
    print('Ao finalizar suas leituras, pressione STOP no aparelho.')
    sleep(2.5)
    print('Ao pressionar STOP, o programa levará alguns segundos para preparar sua planilha. Aguarde...')
    return sample_name


def serial_object_creator(time_set):
    ser = serial.Serial('COM1', 9600, timeout=time_set)
    serial_object = ser.readline().split()
    return serial_object


def timer_for_closing_port(object):
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
    print(f' RPM: {float(object[3]):.>20} /// cP: {int(object[7]):.>20} /// Torque: {float(object[5]):.>20}%')


def values_storager(object):
    if float(object[3]) not in registers.keys():
        registers[float(object[3])] = [[int(object[7])], [float(object[5])]]
    elif float(object[3]) in registers.keys():
        registers[float(object[3])][0].append(int(object[7]))
        registers[float(object[3])][1].append(float(object[5]))
    return registers


def sheet_maker(sample_name, **registers):    
    if len(registers) > 0:


        def data_processor(**registers):
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


        workbook = xlsxwriter.Workbook(f'{sample_name}.xlsx')
        worksheet = workbook.add_worksheet()
        bold = workbook.add_format({'bold': True})
        worksheet.set_column(0, 8, 20)
        worksheet.set_column(4, 4, 25)
        worksheet.write('A1', f'{sample_name}', bold)
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
            'categories': f'=Sheet1!$G2$:$G${len(processed_registers.keys()) + 1}', 
            'values': f'=Sheet1!$H$2:$H${len(processed_registers.values()) + 1}', 
            'line': {'color': 'green'}
            })
        chart.set_title({'name': f'{sample_name}'})
        chart.set_x_axis({
            'name': 'RPM',
            'name_font': {'size': 14, 'bold': True},
        })
        chart.set_y_axis({
            'name': 'cP',
            'name_font': {'size': 14, 'bold': True},
        })
        worksheet.insert_chart('F18', chart)
        workbook.close()
        print('Aguarde que uma planilha será aberta com seus resultados.')
        startfile(f'{sample_name}.xlsx')
        return workbook


    else:
        print('Nenhuma planilha será gerada por falta de dados.')


registers = dict()  # Registros das leituras serão armazenados neste dicionário.
time = 200  # Tempo para timeout da porta inicial alto para evitar bugs na inicialização do programa.


initial_menu()
sample_name = sample_name()

sleep(5)  # Tempo de espera para evitar que bugs que o aparelho gera na inicialização possam dar crash no programa.
print('-' * 90)
print('#' * 39, ' LEITURAS ', '#' * 39)
print('-' * 90)
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
            print('Torque máximo atingido')
            print('Leituras não são possíveis de serem feitas')
            print('Pressione STOP no aparelho')

    except KeyboardInterrupt:
        print('Programa interrompido por atalho de teclado')
        break

    except IndexError:
        print('Foi pressionado STOP no aparelho')
        registers = dict(OrderedDict(sorted(registers.items())))
        break


sheet_maker(sample_name, **{str(k): v for k , v in registers.items()})

print('OBRIGADO POR USAR O VISCOTESTER 6L SCRIPT')
sleep(10)
