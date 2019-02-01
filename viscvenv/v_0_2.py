import serial


def serial_object_creation():
    ser = serial.Serial('COM1', 9600, timeout=60)
    serial_object = ser.readline().split()
    return serial_object


def time_for_closing_port(ser):
    pass


def torque_validation(serial_object):
    if serial_object[7] == b'off':
        return False
    else:
        return True


def printing_readings(object):
    print(f'RPM: {int(object[3])} / cP: {float(object[7])} / Torque(%): {float(object[5])}')


def storing_RPMs(object):
    pass


while True:
    try:
        object = serial_object_creation()
        if torque_validation(object):
            printing_readings(object)
        else:
            print('Torque máximo atingido')
            print('Leituras não são possíveis de serem feitas')
            print('Pressione STOP no aparelho')

    except KeyboardInterrupt:
        print('Programa interrompido')
        break


print('Aqui temos o fim do programa')

