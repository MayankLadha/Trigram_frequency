import itertools
from itertools import permutations
import xlsxwriter

workbook = xlsxwriter.Workbook('Trigram_Frequency.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': 1})

fo = open("sample2.txt", "r")
freq = {}
b = True
alpha = list('abcdefghijklmnopqrstuvwxyz')
res = [''.join(p) for p in itertools.permutations(alpha, r=3)]

for x in res:
    freq[x]=0

while b:
    hs = fo.read(100)
    if len(hs) != 0:
        for x in res:
            freq[x] += hs.count(x)
    else:
        b = False

row=1
col=0
worksheet.set_row(0, 20, bold)
worksheet.set_column('A:A', 20)
worksheet.set_column('B:B', 20)
worksheet.write(0,0,"Trigram")
worksheet.write(0,1,"Frequency")

print ("Trigram : frequency")
for x in res:
    col=0
    print (x+"      : "+str(freq[x]))
    worksheet.write(row,col,x)
    col+=1
    worksheet.write(row,col,freq[x])
    row+=1

sorted_freq = sorted(freq.items(), key=lambda x: x[1], reverse=True)

row=1
col=5
worksheet.set_column('F:F', 20)
worksheet.set_column('G:G', 20)
worksheet.write(0,5,"Trigram")
worksheet.write(0,6,"Frequency")

for x in range(0,15600):
    col=5
    worksheet.write(row,col,sorted_freq[x][0])
    col+=1
    worksheet.write(row,col,sorted_freq[x][1])
    row+=1

workbook.close()
print ("Report successfully writtern to Trigram_Frequency.xlsx")
