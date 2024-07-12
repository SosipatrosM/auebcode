import argparse
import pandas as pd
import sys
from openpyxl import Workbook
#hey

parser = argparse.ArgumentParser(description='Team Formation Program')
parser.add_argument('input', type=str, help='Input Excel file')
parser.add_argument('-o', '--output', type=str, help='Output Excel file')
args = parser.parse_args()  

# Read input Excel file
input_data = pd.read_excel(args.input)
data_list = input_data.values.tolist()

def exists_already(id, matrix):
    flag = False
    for i in range(len(matrix)):
        for j in matrix[i]:
            if (j == id):
                flag = True
                break
    return flag

row0 = [0, 0, ""]
pairs = [row0]
singles = [row0]
for i in range(len(data_list)):
    if (data_list[i][3] != None):
        if (not exists_already(data_list[i][0], pairs)):
            pairs.append([data_list[i][0], data_list[i][3], data_list[i][1]])
    else:
        if (not exists_already(data_list[i][0], singles)):
            singles.append([data_list[i][0], data_list[i][3], data_list[i][1]])

pairs.pop(0)
singles.pop(0)

if (len(pairs)<2):
    print("not enough pairs")
    sys.exit()

groups = [[pairs[0], pairs[1], 0, 0, 0, 0, 0]]

if (pairs[0][2]=="male"):
    groups[0][5] += 1
else:
    groups[0][6] += 1

for i in range(len(data_list)):
    gender = ""
    if (data_list[i][0]==pairs[0][1]):
        gender=data_list[i][1]

for i in range(1, len(pairs)):
    groups.append([0, 0, 0, 0, 0, 0, 0])
    groups[i][0] = pairs[i][0]
    groups[i][1] = pairs[i][1]
    if (pairs[i][2]=="male"):
        groups[i][5] += 1
    else:
        groups[i][6] += 1
    if (pairs[i][2]=="male"):
        groups[i][5] += 1
    else:
        groups[i][6] += 1

for i in range(len(groups)//2):
    if (groups[i][0] != pairs[i][0]):
        groups[i][2] = pairs[-1][0]
        groups[i][3] = pairs[-1][1]
        if (pairs[-1][0]=="male"):
            groups[i][5] += 1
        else:
            groups[i][6] += 1
        if (pairs[-1][1]=="male"):
            groups[i][5] += 1
        else:
            groups[i][6] += 1
        pairs.pop(-1)

for single in singles:
    for group in groups:
        if (group[5] + group[6] < 5):
            group[4] = single[0]
            if (single[0]=="male"):
                group[5] += 1
            else:
                group[6] += 1
            break
      

wb = Workbook()
ws = wb.active

for row in groups:
    ws.append([str(elem) for elem in row])

# Save the workbook
wb.save('teams.xlsx')

print("Team formation completed successfully!")
