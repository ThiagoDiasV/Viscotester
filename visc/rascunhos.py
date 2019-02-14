from math import log10

RPM = (0.3, 0.5, 0.6, 1, 1.5, 2, 2.5, 3, 4, 5, 6, 10, 12, 20, 30, 50, 60, 100, 200)
cP = (1000, 900, 800, 700, 600, 500, 400, 300, 200, 100, 80, 60, 40, 20)

lista = tuple(log10(x) for x in RPM)
lista2 = tuple(log10(x) for x in cP)
print(type(lista))
for i in lista:
    print(i)
print(type(lista2))
for i in lista2:
    print(i)
