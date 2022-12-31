from project import fullname

f1 = open("students\\" + fullname + "\courses.txt", "r")
f2 = open("students\\" + fullname + "\semesters.txt", "r")

#Read semesters from file and insert it into list semesters
semesters = []
for line in f2:
    line=line.strip()
    line = line[6:]
    if len(line) > 0:
        semesters.append(line)
f2.close()

#Read courses from file and insert it into list courses.
#Ensure there is only one dash between courses
courses = []
for line in f1:
    line=line.strip()
    #Create a list by splitting the a line. Each word is an item in list
    lineRec = line.split()
    courses.append(lineRec)
f1.close()

#Remove classes that are missing a grade
j = 0
while j < len(courses):
    if len(courses[j])>1:
        if (courses[j][-2] == 'A') or (courses[j][-2] == 'B') or (courses[j][-2] == 'C') or \
           (courses[j][-2] == 'D') or (courses[j][-2] == 'S') or (courses[j][-2] == 'SP') or \
           (courses[j][-2] == 'F') or (courses[j][-2] == 'W') or (courses[j][-2] == 'U') or \
           (courses[j][-2] == 'N') or (courses[j][-2] == 'I') or (courses[j][-2] == 'TR') or \
           (courses[j][-2] == 'P') or (courses[j][-2] == 'inprog'):                   
            #do nothing
            j += 1
            continue
        else:
            courses.pop(j)
            j -= 1
    j += 1

#Adjust courses to have course title as one element in list
#For example ['CHEM', '152', 'General', 'Chemistry', 'II', 'S', '0.000'] becomes
#['CHEM', '152', 'General Chemistry II ', 'S', '0.000']
j=0
for rec in courses:
    courseName = ""
    #For those who actually are course records not "-"
    if len(rec) > 1:
        size = len(rec)
        start = 2
        end = size-2
        popCount = 0
        #Create a course title in a single string
        for i in range (start, end):
            if (i<end):
                courseName = courseName + rec[i] + " "
            else:
                courseName = courseName + rec[i]
            #While creating string remove the title words from course record
            popCount = popCount + 1
        for itr in range(popCount):
            rec.pop(2)
        #insert back course name as one string title
        courses[j].insert(2, courseName)
    j=j+1

#Unify semester and course lists (Final Data structure view)
#add semester to end of list
i = 0
for s in semesters:
    while (courses[i][0] != "-"):
        print(courses[i])
        courses[i].append(s)
        i += 1
    i += 1
print("------------------------------")

#Then remove dashes
for c in courses:
    if c[0] == "-":
        courses.remove(c)

#join course acronym and number together
for c in courses:
    course = c[0]+ " " + c[1]
    c[0] = course
    c.pop(1)

#Replace 'inprog' with 'In progress'
for c in courses:
    if c[2] == 'inprog':
        c[2] = 'In progress'

#remove failed courses
#course, name, grade, credits, semester    
i = 0
while i < len(courses):
    c = courses[i]
    #For certain courses, D is a failing grade
    if c[0][:-4] == 'CSCI' or c[0][:-4] == "MATH" or \
         c[0] == "ENGL 110" or c[0] == "ENGL 111":
        if c[2] == 'F' or c[2] == 'U' or c[2] == 'N' or \
        c[2] == 'D' or c[2] == 'W' or c[2] == 'I':
            print(c)
            courses.pop(i)
            #decrement or else loop will skip
            #elements after popping an element
            i -= 1
    #For all other courses, D is not a failing grade
    elif c[2] == 'F' or c[2] == 'U' or c[2] == 'N' or \
         c[2] == 'W' or c[2] == 'I':
        print(c)
        courses.pop(i)
        i -= 1
    i += 1
print("------------------------------")

for c in courses:
    print(c)
