import re
regex = re.compile(r'[NnSs]')
op = ''
while not regex.search(op):
    op = str(input('[S/N]: ')).strip().upper()[0]
    print(op)
print('fim')