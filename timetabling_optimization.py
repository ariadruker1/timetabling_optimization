import time
import pandas as pd
from ortools.linear_solver import pywraplp


AllData = pd.ExcelFile("Problem Set 3 - Input Data.xlsx")
StudentCourseData = pd.read_excel(AllData, 'StudentCourse')
TeacherCourseData = pd.read_excel(AllData, 'TeacherCourse')
TeacherBlockData = pd.read_excel(AllData, 'TeacherDay')
CourseBlockData = pd.read_excel(AllData, 'CourseDay')

print(StudentCourseData)
print("")

print(TeacherCourseData)
print("")

print(TeacherBlockData)
print("")

print(CourseBlockData)


# This is how we read the above data.
#
# StudentCourseData[8][3] ='Y' means that Course C8 is desired by Student S3
# TeacherCourseData[9][5] ='Y' means that Course C9 can be taught by Teacher T5
# TeacherBlockData[1][5] ='N' means that Day D1 cannot be assigned to Teacher T5
# CourseBlockData[3][0] ='N' means that Day D3 cannot be assigned to Course C0


numStudents = 8
numTeachers = 6
numCourses = 12
numDays = 4

allStudents = range(numStudents)
allTeachers = range(numTeachers)
allCourses = range(numCourses)
allDays = range(numDays)

StudentList=['S0','S1','S2','S3','S4','S5','S6','S7']
TeacherList=['T0','T1','T2','T3','T4','T5']
CourseList=['C0','C1','C2','C3','C4','C5','C6','C7','C8','C9','C10','C11']
DayList=['D0','D1','D2','D3']

def _cell(x):  # normalize to uppercase string without spaces
    return str(x).strip().upper()

# Correct orientations (+1 to skip the label column)
desires         = lambda s, c: 1 if _cell(StudentCourseData.iloc[s, c + 1]) == 'Y' else 0
can_teach       = lambda t, c: 1 if _cell(TeacherCourseData.iloc[t, c + 1]) == 'Y' else 0
teacher_available = lambda t, d: 0 if _cell(TeacherBlockData.iloc[t, d + 1]) == 'N' else 1  # 'N' means unavailable
course_allowed    = lambda c, d: 0 if _cell(CourseBlockData.iloc[c, d + 1]) == 'N' else 1   # 'N' means blocked

solver = pywraplp.Solver('Timetabling Problem', pywraplp.Solver.CBC_MIXED_INTEGER_PROGRAMMING)

start_time = time.time()

# Define our binary variables for the students and teachers
X = {}
for s in allStudents:
    for c in allCourses:
        for d in allDays:
            X[s,c,d] = solver.BoolVar('X[%i,%i,%i]' % (s,c,d))

Y = {}
for t in allTeachers:
    for c in allCourses:
        for d in allDays:
            Y[t,c,d] = solver.BoolVar('Y[%i,%i,%i]' % (t,c,d))

CD = {}
for c in allCourses:
    for d in allDays:
        CD[c,d] = solver.BoolVar('CD[%i,%i]' % (c,d))
            

            
# Define our objective function
solver.Maximize(solver.Sum(desires(s, c) * X[s,c,d] for s in allStudents for c in allCourses for d in allDays))



# Each student must take one course on each day
for s in allStudents:
    for d in allDays:
        solver.Add(solver.Sum([X[s,c,d] for c in allCourses]) == 1)

for s in allStudents:
    for c in allCourses:
        for d in allDays:
            solver.Add(X[s, c, d] <= CD[c, d])

# Each course taught on exactly one day
for c in allCourses:
    solver.Add(solver.Sum(CD[c, d] for d in allDays) == 1)

# Exactly three courses per day
for d in allDays:
    solver.Add(solver.Sum(CD[c, d] for c in allCourses) == 3)

# A scheduled course has exactly one teacher on that day
for c in allCourses:
    for d in allDays:
        solver.Add(solver.Sum(Y[t, c, d] for t in allTeachers) == CD[c, d])

# Each teacher teaches exactly two courses total
for t in allTeachers:
    solver.Add(solver.Sum(Y[t, c, d] for c in allCourses for d in allDays) == 2)

# Qualification: teacher must be able to teach the course
for t in allTeachers:
    for c in allCourses:
        if not can_teach(t, c):
            for d in allDays:
                solver.Add(Y[t, c, d] == 0)

# Teacher-day availability: 'N' = unavailable
for t in allTeachers:
    for d in allDays:
        if teacher_available(t, d) == 0:
            for c in allCourses:
                solver.Add(Y[t, c, d] == 0)

# Course-day blocks: 'N' = cannot offer
for c in allCourses:
    for d in allDays:
        if course_allowed(c, d) == 0:
            solver.Add(CD[c, d] == 0)

# No teacher may teach more than one course per day
for t in allTeachers:
    for d in allDays:
        solver.Add(solver.Sum([Y[t,c,d] for c in allCourses]) <= 1)  


        
    
current_time = time.time() 
reading_time = current_time - start_time         
sol = solver.Solve()
solving_time = time.time() - current_time

print('Optimization Complete with Total Happiness Score of', round(solver.Objective().Value()))
print("")
print('Our program needed', round(solving_time,3), 'seconds to determine the optimal solution')
                



# Print Output for Students and Teachers

for s in allStudents:
    for c in allCourses:
        for d in allDays:
            if X[s,c,d].solution_value() == 1:
                print("Student", StudentList[s], "is taking Course", CourseList[c],
                      "on Day", DayList[d])
    print("")


for t in allTeachers:
    for c in allCourses:
        for d in allDays:
            if Y[t,c,d].solution_value() == 1:
                print("Teacher", TeacherList[t], "is teaching Course", CourseList[c],
                      "on Day", DayList[d])
    print("")

