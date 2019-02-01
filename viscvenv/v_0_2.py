import serial


def serial_object_creation(time_set):
    ser = serial.Serial('COM1', 9600, timeout=time_set)
    serial_object = ser.readline().split()
    return serial_object


def time_for_closing_port(object):
    time_for_closing = (60/int(object[3])) + 5
    return time_for_closing


def torque_validation(serial_object):
    if serial_object[7] == b'off':
        return False
    else:
        return True


def printing_readings(object):
    print(f'RPM: {int(object[3])} / cP: {float(object[7])} / Torque(%): {float(object[5])}')


def storing_values(object):
    if int(object[3]) not in registers.keys():
        registers[int(object[3])] = [[float(object[7])], [float(object[5])]]
    elif int(object[3]) in registers.keys():
        registers[int(object[3])][0].append(float(object[7]))
        registers[int(object[3])][1].append(float(object[5]))
    print(registers)
    return registers


registers = dict()

time = 200
while True:
    try:
        object = serial_object_creation(time)
        time = time_for_closing_port(object)
        if torque_validation(object):
            if object == False:
                print('O aparelho foi stoppado')
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


print('FIM DO PROGRAMA')
