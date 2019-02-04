import serial
import xlsxwriter
from statistics import mean, stdev
from os import startfile
# import seaborn as sns


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
    print('6 - Observe se o fuso correto está acoplado ao aparelho e pressione START')
    print('-' * 90)
    print('#' * 90)
    print('#' * 90)
    print('-' * 90)


def sample_name():
    sample_name = str(input('Digite o nome da amostra: '))
    return sample_name


def serial_object_creation(time_set):
    ser = serial.Serial('COM1', 9600, timeout=time_set)
    serial_object = ser.readline().split()
    return serial_object


def time_for_closing_port(object):
    time_for_closing = (60/float(object[3])) + 10
    return time_for_closing


def torque_validation(serial_object):
    if serial_object[7] == b'off':
        return False
    else:
        return True


def printing_readings(object):
    print(f'RPM: {float(object[3])} / cP: {int(object[7])} / Torque(%): {float(object[5])}')


def storing_values(object):
    if str(float(object[3])) not in registers.keys():
        registers[str(float(object[3]))] = [[int(object[7])], [float(object[5])]]
    elif str(float(object[3])) in registers.keys():
        registers[str(float(object[3]))][0].append(int(object[7]))
        registers[str(float(object[3]))][1].append(float(object[5]))
    print(registers)
    return registers


def sheet_maker(sample_name, **registers):


    def data_processor(**registers):
        for value in registers.values():
            mean_value = mean(value[0])
            std_value = stdev(value[0])
            cp_list = [x for x in value[0] if (x > mean_value - std_value)]
            cp_list = [x for x in cp_list if (x < mean_value + std_value)]
            value[0] = cp_list
        return registers


    workbook = xlsxwriter.Workbook(f'{sample_name}.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    worksheet.write('A1', f'{sample_name}', bold)
    worksheet.write('B1', 'RPM', bold)
    worksheet.write('C1', 'cP', bold)
    worksheet.write('D1', 'Torque(%)', bold)
    worksheet.write('E1', 'Processamento dos dados >>', bold)
    worksheet.write('H1', 'RPM', bold)
    worksheet.write('I1', 'cP', bold)
    row = 1
    col = 1
    for key, value in registers.items():
        worksheet.write(row, col, float(key))
        for cp in value[0]:
            worksheet.write(row, col + 1, cp)
            row += 1
        row -= len(value[0])
        for torque in value[1]:
            worksheet.write(row, col + 2, torque)
            row += 1
    processed_registers = data_processor(**registers)
    row = col = 1
    for key, value in processed_registers.items():
        worksheet.write(row, col + 6, float(key))
        worksheet.write(row, col + 7, mean(value[0]))
        row += 1
    workbook.close()
    return workbook


registers = dict()
time = 200


initial_menu()
sample_name = sample_name()


while True:
    try:
        object = serial_object_creation(time)
        time = time_for_closing_port(object)
        if torque_validation(object):
            if object == False:
                print('Torque máximo atingido ou erro no aparelho')
            else:
                printing_readings(object)
                registros = storing_values(object)
        else:
            print('Torque máximo atingido')
            print('Leituras não são possíveis de serem feitas')
            print('Pressione STOP no aparelho')

    except KeyboardInterrupt:
        print(f'Resultados registrados: {registers}')
        print('Programa interrompido por atalho de teclado')
        break

    except IndexError:
        print(f'Resultados registrados: {registers}')
        print('Foi pressionado STOP no aparelho')
        break

print('Aguarde que uma planilha será aberta com seus resultados.')

sheet_maker(sample_name, **registers)

startfile(f'{sample_name}.xlsx')

print('FIM DO PROGRAMA')
