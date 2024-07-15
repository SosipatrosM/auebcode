import argparse
import pandas as pd
import sys
from openpyxl import Workbook
from openpyxl.styles import Font
from copy import deepcopy
import time

#Run in CMD using-> python form_teams.py students.xlsx -g , -g is optional

print("Processing", end="") #Loading screen
for _ in range(4):
    sys.stdout.write('.')
    sys.stdout.flush()
    time.sleep(0.5)

parser = argparse.ArgumentParser(description='Team Formation Program')
parser.add_argument('input', type=str, help='Input Excel file')
parser.add_argument('-o', '--output', type=str, help='Output Excel file')
parser.add_argument('-g', action='store_true', help='Optional flag if users needs gender printed')
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
    
def swap_gender(gender):
    if gender == 'male':
        return 'female'
    else:
        return 'male'

class Student:
    def __init__(self, id, gender, score, friendid):
        self.id = id
        self.gender = gender
        self.score = score
        self.friendid = friendid
        self.category = 0
        self.team = 0

class Team:

    def __init__(self, id, members):
        self.id = id
        # Ensure members is always a list
        if isinstance(members, list):
            self.members = members
        else:
            if members != 0:
                self.members = [members] 
            else:
                self.members = []
        self.males = 0
        self.females = 0
        self.scorebalance = 0
        for member in self.members:
            if member.gender == 'male':
                self.males +=1
            else:
                self.females +=1
            self.scorebalance += member.category
    
    def number_of(self, gender):
        if gender == 'male':
            return self.males
        else:
            return self.females

    def size(self):
        return len(self.members)
    
    def add_member(self, newmember):
        self.members.append(newmember)
        if newmember.gender == 'male':
            self.males += 1
        else:
            self.females += 1
        self.scorebalance += newmember.category
    
    def pop_member(self, oldmember):
        toberemoved = None
        for i in range(len(self.members)):
            if oldmember.id == self.members[i].id:
                toberemoved = self.members.pop(i)
                break
        if toberemoved:
            if oldmember.gender == 'male':
                self.males -= 1
            else:
                self.females -= 1
            self.scorebalance -= oldmember.category
        return toberemoved

#Check for header
preview_data = pd.read_excel(args.input, nrows=5)
if all(isinstance(x, str) for x in preview_data.columns):
    # Read the entire file with the first row as the header
    input_data = pd.read_excel(args.input)
    data_list = input_data.values.tolist()
else:
    # Read the entire file without treating the first row as the header
    input_data = pd.read_excel(args.input, header=None)
    data_list = input_data.values.tolist()


sorted_by_id = sorted(data_list, key=lambda x: x[0])
for i in range(len(sorted_by_id) - 1):
    if not is_number(sorted_by_id[i][0]):
        print("ERROR: Missing students. Row:", i) #LOGIC ISSUE
        sys.exit(1)
    if sorted_by_id[i][0]==sorted_by_id[i+1][0]:
        j = i+1
        print("ERROR: Duplicate students found. Rows:", i, j) #LOGIC ISSUE
        sys.exit(1)


sorted_by_score = sorted(data_list, key=lambda x: x[2])

# Calculate boundaries
n = len(sorted_by_score)
b1 = n // 4   # First boundary: 25th percentile
b2 = n // 2   # Second boundary: 50th percentile (median)
b3 = 3 * n // 4   # Third boundary: 75th percentile

count = {}
count['male'] = 0
count['female'] = 0
non_specified = []
all_students = []
for i in range(len(sorted_by_score)):
    if sorted_by_score[i][1] != 'male' and sorted_by_score[i][1] != 'female': #fixes value issues with gender, primarly NaN=no value
        sorted_by_score[i][1] = 'male'
        print("\nNot specified gender for id =", sorted_by_score[i][0], "\n")
        non_specified.append(i)
    
    if sorted_by_score[i][1] == 'male':
        count['male'] += 1
    else:
        count['female'] +=1
    
    if not(is_number(sorted_by_score[i][2])): #fixes value issues with score, primarly NaN
        sorted_by_score[i][2] = 0
    else:
        sorted_by_score[i][3] = sorted_by_score[i][3]
    
    if not(is_number(sorted_by_score[i][3])): #fixes value issues with friend id, primarly NaN
        sorted_by_score[i][3] = 0
    else:
        sorted_by_score[i][3] = int(sorted_by_score[i][3])
    
    student = Student(sorted_by_score[i][0], sorted_by_score[i][1], sorted_by_score[i][2], sorted_by_score[i][3])
    if i<b1:
        student.category = 1 #assigns one of the 4 categories. Condition works because sorted_by_score is sorted based on score
    elif i<b2:
        student.category = 10 
    elif i<b3:
        student.category = 100
    else:
        student.category = 1000
    all_students.append(student)

#INITIALIZE VARIABLES
remaining = {}
remaining['male'] = count['male']  #at some point we will need to know if we have run out of one gender
remaining['female'] = count['female']
new_team_id = 1
all_students[0].team = new_team_id
initial_members = [ all_students[0] ]
team1 = Team(new_team_id, initial_members)
all_teams = [team1]
for index, student in enumerate(all_students):
    remaining[student.gender] -= 1 #A student will be assigned a team ALWAYS or ignored in the next line
    if student.team != 0: #Ignore student if they have a team
        continue

    entered = False
    for team in all_teams: #Normally there should be an appropriate already-formed team
        if team.size() < 4 and ((team.number_of(student.gender)<2 and not (team.males == 3 and student.gender == 'female')) or (team.males==3 and student.gender == 'male') or (team.females==3 and student.gender == 'female') or (team.females==2 and student.gender=='female' and remaining['male']<=1)):
            entered = True
            student.team = team.id
            team.add_member(student)
            break

    #For very bad data e.g. too many consecutive men, we make exceptions and create same sex teams
    if not entered:
        extra = 0
        x = (count[swap_gender(student.gender)] - remaining[swap_gender(student.gender)])%2 #is "how many used of the opposite gender" odd or even?
        for i in range(index+1, len(all_students)): #Step1: check upcoming students
            if i >= index + 7-extra or remaining[swap_gender(student.gender)] <= 2-x: #If any of this two is true then we have to create a new team
                break
            if all_students[i].gender == student.gender:
                for each_team in all_teams:
                  if each_team.size() <4 and each_team.males*each_team.females==0 and not (each_team.males == 3 and student.gender == 'female'): #not a full team and the other gender is zero
                    entered = True
                    extra +=1 #ensures same sex teams
                    student.team = each_team.id
                    each_team.add_member(student)
                    break
            else:
                if all_students[i].team != 0:
                    continue
                else:
                    break

    if not entered: #all attempts to put him/her in an existing team failed. Now let's make a new one
        new_team_id += 1
        student.team = new_team_id
        initial_members = [student]
        new_team = Team(new_team_id, initial_members)
        all_teams.append(new_team)

#---------ATTEMPTS TO CREATE SOME TEAMS OF 5 IN ORDER TO FIX INCOMPLETE TEAMS ONLY IF STRAY STUDENTS ARE MAXIMUM 3

bestnum = len(all_students) // 4 #Ideal number of teams
extra_teams = len(all_teams) - bestnum
i = len(all_teams)-1
incomplete = []
extras = 0
while i >= 0 and len(incomplete) < extra_teams: #FINDS INCOMPLETE TEAMS EVERYWHERE DESPITE THEM USUALLY APPEARING NEAR THE END
    if len(all_teams[i].members) < 4:
        incomplete.append(all_teams[i].id)
        extras += len(all_teams[i].members)
    i -=1

i = len(all_teams) - 1

if extras <= 3:
    for j in incomplete:
        # Create a copy of the member list to avoid modifying the list while iterating
        members_to_move = all_teams[j-1].members.copy()
        for student in members_to_move:
            while i >= 0 and extras > 0:
                if len(all_teams[i].members) == 4:  # Look for a team that already has 4 members
                    student.team = all_teams[i].id  # Assign student to the new team
                    all_teams[i].add_member(all_teams[j-1].pop_member(student))  # Move student to the new team
                    extras -= 1  # Decrement extras since we assigned one student
                    break  # Move to the next student after successfully moving the current one
                i -= 1  # Move to the next team if current team is not suitable
all_teams = [team for team in all_teams if len(team.members) > 0] #delete empty teams

#----CHECKS ONE LAST TIME FOR INCOMPLETE TEAMS
incomplete = []
for team in all_teams:
    if team.size() < 4:
        incomplete.append(team.id)

isolated_females = []
for team in all_teams:
    if team.females == 1:
        isolated_females.append(team.id)

for index in non_specified:
    all_students[index].gender = 'not specified'

students_appeared = {} #checks for duplicate students
iteratable_teams = []
for team in all_teams:
    memberlist = [team.id]  # Start with the team ID
    for member in team.members:
        toprint = str(member.id)
        if args.g:
            gflag = "-"+member.gender[0]
            toprint += gflag
        memberlist.append(toprint)
        if member.id in students_appeared:
            students_appeared[member.id] +=1
        else:
            students_appeared[member.id] = 1
    missing = 5-len(team.members)
    for i in range(missing):
        memberlist.append(" ")
    memberlist.append(team.scorebalance)
    memberlist.append(team.males)
    memberlist.append(team.females)
    errorm = "ISSUE?" + (f"Student reappears {students_appeared[member.id]-1} time(s)!" if students_appeared[member.id] > 1 else "")
    if team.id in incomplete or team.id in isolated_females or len(errorm)> 8:
        memberlist.append(errorm)
    iteratable_teams.append(memberlist)
wb = Workbook()
ws = wb.active

headers = ["Team ID", "Member1", "Member2", "Member3", "Member4", "Member5", "Balance", "Males", "Females", "Comments"]
ws.append(headers)
for team in iteratable_teams:
    ws.append([str(elem) for elem in team])

for team_id in incomplete:
    print("\nTeam", team_id, "incomplete!")


# Save the workbook
wb.save('teams.xlsx')

print("\n\nDone!")
print("You will find a teams.xlsx file in the same folder as the script. Previous teams.xlsx files will be overwritten.")
print("Don't forget you can -g to ask the script to print each students gender e.g. for output evaluation")
