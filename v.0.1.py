import serial
import statistics
import xlsxwriter
import os

print('\033[1;36m-=-\033[m'*20)
print('                      \033[1;32mVisco Tester 6L\033[m')
print('\033[1;36m-=-\033[m'*20)
print('\033[1;31m ### INSTRUÇÕES DE USO DO APARELHO ###\033[m')
print('\033[1;37m1 - Peça orientação aos funcionários do laboratório para o uso do viscosímetro; '
      '\n2 - O aparelho possui 4 fusos disponíveis para uso. A escolha do fuso depende da viscosidade da amostra; '
      '\n3 - Não retire ou coloque o fuso no aparelho sem a orientação dos funcionários do laboratório; '
      '\n4 - Abaixo serão solicitadas algumas informações sobre sua amostra; '
      '\n5 - Quando as leituras começarem, aguarde pelo menos umas 5 leituras em cada velocidade; '
      '\n6 - Caso as leituras estejam variando bastante, aguarde um pouco mais além das 5 leituras acima, '
      '\n    até que elas fiquem relativamente constantes;'
      '\n7 - Ao final das leituras, quando for registrado o torque máximo ou sejam realizadas leituras a 200 RPM'
      '\n    o programa finalizará e mostrará uma planilha do Excel com seus resultados. \033[m')
print('\033[1;31m #####################################\033[m')

nomeplanilha = str(input('Digite um nome para o arquivo Excel a ser criado: '))
workbook = xlsxwriter.Workbook(f'{nomeplanilha}.xlsx')
infoamostra = str(input('Nome da amostra: '))
print('Pressione \033[1;31mSTART\033[m no viscosímetro para iniciar e aguarde as leituras serem feitas. ')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True})
italic = workbook.add_format({'italic': True})

worksheet.write(0, 0, 'ID amostra', bold)
worksheet.write(1, 0, f'{infoamostra}', italic)
worksheet.write(0, 1, 'RPM', bold)
worksheet.write(0, 2, f'cP: {infoamostra}', bold)
worksheet.write(0, 3, 'Desvio Padrão', bold)

row = 1
col = 0
ser = serial.Serial('COM1', 9600)
listavisc = []
listaRPM = []
resultados = []
resultadosvisc = []
teste = []
somaleituras = 0
somafinal = 0
pos = 0
mediaviscfinal = []
desviolist = []
valoresfinais = []
a = (ser.readline().split())
print('Leituras sendo realizadas...')
a = (ser.readline().split())
teste.append(a[7])
if teste[0] == b'off':
    print('Torque máximo registrado. \nLeituras não são possíveis de serem feitas. \nPressione \033[1;31mSTOP\033[m no aparelho.')
else:
    print('Preparando o programa. Aguarde mais uns instantes...')
    RPM = float(a[3])
    torque = float(a[5])
    viscosidade = int(a[7])
    listaRPM.append(float(a[3]))
teste.clear()
while True:
    a = (ser.readline().split())
    teste.append(a[7])
    if teste[0] == b'off':
        if somafinal == 0:
            somaleituras = 0
        print('Leitura não realizada. Erro!')
        if float(a[3]) < 6 and len(listaRPM) >= 2:
            somafinal += 2
            somaleituras += 7
        else:
            somafinal += 1
            somaleituras += 1
    if float(a[3]) < RPM:
        listavisc.clear()
        listaRPM.remove(listaRPM[len(listaRPM)-1])
        print('A velocidade de rotação do fuso foi diminuída. \nRecalculando parâmetros...')
        RPM = float(a[3])
        if RPM not in listaRPM:
            listaRPM.append(float(a[3]))
        if RPM in listaRPM and len(resultadosvisc) > 0:
            resultadosvisc.remove(resultadosvisc[len(resultadosvisc)-1])
    elif RPM != float(a[3]) and teste[0] != b'off':
        somaleituras = 0
        resultadosvisc.append(listavisc[:])
        listavisc.clear()
        print(f'RPM = {float(a[3])} / cP = {int(a[7])}')
        if int(a[7]) != 0:
            listavisc.append(int(a[7]))
        listaRPM.append(float(a[3]))
        RPM = float(a[3])
        torque = float(a[5])
        somaleituras += 1
    elif RPM == float(a[3]) and teste[0] != b'off':
        print(f'RPM = {float(a[3])} / cP = {int(a[7])}')
        if int(a[7]) != 0:
            listavisc.append(int(a[7]))
        somaleituras += 1
        if RPM == 200 and somaleituras > 6:
            resultadosvisc.append(listavisc[:])
    print(listaRPM)
    print(listavisc)
    print(resultadosvisc)
    if torque > 98 or somafinal >= 2 or RPM == 200 and somaleituras >= 7:
        resultadosvisc.append(listavisc[:])
        listaRPM.append(float(a[3]))
        if len(listaRPM) > len(resultadosvisc):
            del(listaRPM[len(listaRPM) - 1])
        print('\033[1;34mLeituras finalizadas\033[m')
        print('\033[1;31mPRESSIONE STOP NO APARELHO! \033[m')
        print('Uma planilha do Excel será aberta com os resultados. \nAguarde.')
        for c in resultadosvisc:
            if len(c) >= 2:
                media = statistics.mean(c)
                desvio = statistics.stdev(c)
                limsuperior = media + desvio
                liminferior = media - desvio
            for i, j in enumerate(c):
                while j > limsuperior or j < liminferior and len(c) >= 2:
                    c.remove(j)
                    media = statistics.mean(c)
                    desvio = statistics.stdev(c)
                    limsuperior = media + desvio
                    liminferior = media - desvio
                    j = c[i]
                break
            desviolist.append(desvio)
        for c in resultadosvisc:
            mediafinal = statistics.mean(c)
            mediaviscfinal.append(mediafinal)
        worksheet.write_column('B2', listaRPM)
        worksheet.write_column('C2', mediaviscfinal)
        worksheet.write_column('D2', desviolist)
        workbook.close()
        os.startfile(f'{nomeplanilha}.xlsx')
        break
    teste.clear()
