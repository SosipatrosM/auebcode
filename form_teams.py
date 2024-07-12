import argparse
import pandas as pd
import sys
from openpyxl import Workbook
from copy import deepcopy

#cd C:\Users\Sosipatros\Documents\GitHub\auebcode
#python form_teams.py students.xlsx

parser = argparse.ArgumentParser(description='Team Formation Program')
parser.add_argument('input', type=str, help='Input Excel file')
parser.add_argument('-o', '--output', type=str, help='Output Excel file')
args = parser.parse_args()  

def is_number(value):
    try:
        # Try converting to numeric
        numeric_value = pd.to_numeric(value)
        # Explicitly return False if the value is NaN
        if pd.isna(numeric_value):
            return False
        return True
    except ValueError:
        return False

class Student:
    def __init__(self, id, gender, score, friendid):
        self.id = id
        self.gender = gender
        self.score = score
        self.friendid = friendid
        self.category = 0
        self.team = 0

# Read input Excel file
input_data = pd.read_excel(args.input)
data_list = input_data.values.tolist()

sorted_by_id = sorted(data_list, key=lambda x: x[0])
for i in range(0, len(sorted_by_id)-1):
    if sorted_by_id[i][0]==sorted_by_id[i+1][0]:
        print("ERROR: Duplicate students found. Terminating script.")
        sys.exit(1)
    

sorted_by_score = sorted(data_list, key=lambda x: x[2])

# Calculate boundaries
n = len(sorted_by_score)
b1 = n // 4   # First boundary: 25th percentile
b2 = n // 2   # Second boundary: 50th percentile (median)
b3 = 3 * n // 4   # Third boundary: 75th percentile

count_males = 0
count_notmales = 0
all_students = []
for i in range(len(sorted_by_score)):
    if sorted_by_score[i][1] != 'male' and sorted_by_score[i][1] != 'female':
        sorted_by_score[i][1] = 'not-specified'
    
    if sorted_by_score[i][1] == 'male':
        count_males += 1
    else:
        count_notmales += 1
    
    if not(is_number(sorted_by_score[i][2])):
        sorted_by_score[i][2] = 0
    else:
        sorted_by_score[i][3] = sorted_by_score[i][3]
    
    if not(is_number(sorted_by_score[i][3])):
        sorted_by_score[i][3] = 0
    else:
        sorted_by_score[i][3] = int(sorted_by_score[i][3])
    
    student = Student(sorted_by_score[i][0], sorted_by_score[i][1], sorted_by_score[i][2], sorted_by_score[i][3])
    if i<b1:
        student.category = 1
    elif i<b2:
        student.category = 10 #assigns one of the 4 categories. Condition works because sorted_by_score is sorted based on score
    elif i<b3:
        student.category = 100
    else:
        student.category = 1000
    all_students.append(student)
    print(student.id, student.gender, student.score, student.friendid, student.category)

def assign_based_on_gender(students, m, ml, fm, ct, teamed, teams): #Students DONT get to choose any member of their team, 2+ women per team of 4-5, balanced skills
    count_members = m               
    count_males = ml
    count_females = fm
    current_team = ct
    for i in range(len(students)):
        if students[i].team != 0: #Skip if student already assigned a team
            continue
        if count_members >= 4: 
            current_team +=1
            count_members = 1
            if students[i].gender == 'male':
                count_males = 1
                count_females = 0
            else:
                count_females = 1
                count_males = 0
            students[i].team = current_team
        else:
            if students[i].gender == 'male' and count_males >=2:
                continue
            elif students[i].gender == 'female' and count_females >=2:
                continue
            else:
                students[i].team = current_team
                if students[i].gender == 'male':
                    count_males +=1
                else:
                    count_females +=1
                count_members +=1
        teamed +=1 #INCORRECT RESULT, WHY?
        if current_team in teams:
            teams[current_team].append(students[i].id)
        else:
            teams[current_team] = [students[i].id]
            
    return students, count_members, count_males, count_females, current_team, teamed, teams

repeat = 0
stud, cmm, cml, cfm, ct, teamed, teams = deepcopy(all_students), 0, 0, 0, 1, 0, {}
while teamed <= len(all_students) and repeat<=len(all_students)-teamed:
    repeat +=1
    stud, cmm, cml, cfm, ct, teamed, teams = assign_based_on_gender(deepcopy(stud), cmm, cml, cfm, ct, teamed, deepcopy(teams))

print('Repetitions:', repeat)
print(cmm, cml, cfm, ct, teamed, len(teams))
for key in teams:
    print(key, teams[key])


wb = Workbook()
ws = wb.active

for key in teams:
    ws.append([str(elem) for elem in teams[key]])

# Save the workbook
wb.save('teams.xlsx')

print("Team formation completed successfully!")
