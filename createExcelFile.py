from openpyxl import Workbook
import random as rd
import pycountry as pc

with open("Python24EneroLista.txt", encoding='utf8') as f: #// Read names from text file
    nombres = f.read().split("\n")

randCountry = list(pc.countries)
wb = Workbook()
ws = wb.active
poolPrize = 1000001
prizeList = []
#total = 0

for nombre in nombres: #// Create list with random nums that sum 1,000,000
    prize = rd.randrange(poolPrize)
    prizeList.append(prize)
    poolPrize -= prize
    #total += prize

#print(total)
rd.shuffle(prizeList)

ws.append(["Name", "Prize", "Country"])
for nombre, prize in zip(nombres, prizeList): #// For each name we append a "row" to the worksheet
    ws.append([nombre, prize, randCountry[rd.randrange(249)].name])

wb.save("lottery.xlsx") #// Save and create excel file