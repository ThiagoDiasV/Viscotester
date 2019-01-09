import serial
# import statistics
import xlsxwriter
# import os
from time import sleep

print('#VISCOTESTER#')


def workbook_create():  # Cria a planilha onde serão armazenados os dados
    workbook_name = str(input('Nome planilha: '))
    workbook = xlsxwriter.Workbook(f'{workbook_name}.xlsx')
    info_sample = str(input('Nome da amostra: '))
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    italic = workbook.add_format({'italic': True})
    worksheet.write(0, 0, 'ID amostra', bold)
    worksheet.write(1, 0, f'{info_sample}', italic)
    worksheet.write(0, 1, 'RPM', bold)
    worksheet.write(0, 2, f'cP: {info_sample}', bold)
    worksheet.write(0, 3, 'Desvio Padrão', bold)


def serial_object_create():  # Cria o objeto serial, no qual serão atribuídas as leituras do aparelho
    ser = serial.Serial('COM1', 9600)
    serial_object = ser.readline.split()
    return serial_object


def reading_samples(serial_lines, counting, listaRPM, listaCP):
    listaCP = []
    if serial_lines[7] == b'off':  # Caso haja erro de torque máximo
        print('Leitura não realizada. Erro de leitura.')
        counting += 1
        print(counting)
        return counting
    else:
        print(serial_lines)
        if float(serial_lines[3]) not in listaRPM:
            listaRPM.append(float(serial_lines[3]))
            print(listaRPM)
            return listaRPM
        else:
            
            listaRPM.append([serial_lines[7]])
            print(listaRPM, listaCP)
            return listaCP


def error_simulator():
    lista = ['bla', 'bla', 'bla', '5', 'bla', 'bla', 'bla', b'off']
    print(lista)
    return lista


def successful_reading():
    lista = ['bla', 'bla', 'bla', '5', 'bla', 'bla', 'bla', '5000']
    print(lista)
    return lista


def successful_reading_2():
    lista = ['bla', 'bla', 'bla', '6', 'bla', 'bla', 'bla', '5000']
    print(lista)
    return lista


# workbook_create() Cria a planilha

# serial_lines = serial_object_create().split()

serial_lines = successful_reading()
serial_lines_2 = successful_reading_2()  # Quando rodar o programa, em vez de serial_lines, o objeto parâmetro de reading samples será serial_object

counting = 0
listaRPM = []
listaCP = []

while True:  # Criar outra condição mais semântica pra isso
    successful_reading()
    counting = reading_samples(serial_lines, counting, listaRPM, listaCP)
    sleep(1)
    successful_reading()
    counting = reading_samples(serial_lines, counting, listaRPM, listaCP)
    sleep(1)
    successful_reading_2()
    counting = reading_samples(serial_lines_2, counting, listaRPM, listaCP)
    sleep(1)

