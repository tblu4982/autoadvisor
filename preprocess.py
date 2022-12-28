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
numberOfSeperators = 0
for line in f1:
    line=line.strip()
    #Create a list by splitting the a line. Each word is an item in list
    lineRec = line.split()
    #If this list is a course, because list will have more than one word
    if (len(lineRec) > 1):
        numberOfSeperators = 0
        courses.append(lineRec)
    #If this is a dash
    else:
        numberOfSeperators+=1
        #add only one dash as seperator. More than one dash in a row is ignored.
        if numberOfSeperators <= 1:
            courses.append(lineRec)
f1.close()

#Add "In Progress" as a grade for currently registered classes
for rec in courses:
    if len(rec)>1:
        if (rec[-2] == 'A') or (rec[-2] == 'B') or (rec[-2] == 'C') or \
           (rec[-2] == 'D') or (rec[-2] == 'S') or (rec[-2] == 'SP') or \
           (rec[-2] == 'F') or (rec[-2] == 'W') or (rec[-2] == 'U') or \
           (rec[-2] == 'N') or (rec[-2] == 'I') or (rec[-2] == 'TR') or \
           (rec[-2] == 'P'):                   
            #do nothing
            continue
        else:
            rec.insert(-1, 'In progress')

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

#checks for missing data entries
sem_buffer = 0
for c in courses:
    if len(c)>1:
        #check to see if course ids are missing
        if len(c[0]) != 4 and len(c[1]) != 4:
            print("Error! Missing or Incomplete data detected! Check data!")
            print("Missing course IDs!")
            print("Error found here:",c)
            exit()
        #exception handler to detect if last element in array can't be cast to float
        #which should only happen if the course credits info is missing
        try:
            if type(float(c[-1])) == float:
                continue
        #code will throw a ValueError if the last index in an array is not a number
        #which should only happen if the last index is missing course credits data
        except ValueError:
            print("Error! Missing or Incomplete data detected! Check data!")
            print("Missing value for course credits!")
            print("Error found here:",c)
            exit()
    #if we reach this, then we've reached the end of a course semester
    else:
        sem_buffer += 1
#if the number of course semesters logged via courses doesn't match
#the number of semesters, then we have a mismatch
if not sem_buffer == len(semesters):
    print("Error! Missing or Incomplete data detected! Check data!")
    print("Semester mismatch detected!")
    exit()

#Unify semester and course lists (Final Data structure view)
#add semester to end of list
i = 0
for s in semesters:
    while (courses[i][0] != "-"):
        print(courses[i])
        courses[i].append(s)
        i += 1
    i += 1
        
#for i in range(len(courses)-1):
#    print(courses[i])


#Then remove dashes
for c in courses:
    if c[0] == "-":
        courses.remove(c)

#join course acronym and number together
for c in courses:
    course = c[0]+ " " + c[1]
    c[0] = course
    c.pop(1)

#Find the current semester and future semesters on transcript    
curr_sems = []
#curr_sem will hold the first semester listed on transcript
curr_sem = courses[0]
prev_sem = courses[0]

#iterate through matrix to find current and future semesters
i = 0
for c in courses:
    #we will need to check if the semester being compared
    #has changed each time we iterate
    if not c[4] == curr_sem[4]:
        prev_sem = curr_sem
        curr_sem = c
    if prev_sem[2] == "In progress":
        #Unique issue: some transcripts will have courses
        #from a previous semester that have no grade marked as
        #'In progress', perform a check to see if there are classes
        #taken after such a course with a final grade
        if not curr_sem[2] == "In progress":
            continue
        #curr_sems will be empty the 1st time this condition passes
        #so we can immediately append it to the empty array
        elif len(curr_sems) == 0:
            curr_sems.append(prev_sem[4])
        #it is done this way because there may be multiple classes
        #from different semesters that will pass this check
        #the next part of this loop will filter duplicate semesters out
        else:
            #to find a unique semester, assume that it is one
            #and prove that it is not a unique semester
            uniqueSem = True
            if not curr_sem == prev_sem:
                for sem in curr_sems:
                    if curr_sem[4] == sem:
                        uniqueSem = False
            if uniqueSem:
                curr_sems.append(curr_sem[4])
            uniqueSem = True
            for sem in curr_sems:
                if prev_sem[4] == sem:
                    uniqueSem = False
            if uniqueSem:
                curr_sems.append(prev_sem[4])
                
#Alternate code if we reach end of matrix but there are no previous semesters
#with grades of "in progress"
if len(curr_sems) == 0:
    if curr_sem[2] == "In progress":
        curr_sems.append(curr_sem[4])

print("------------------------------")

#remove failed courses
#course, name, grade, credits, semester    
i = 0
while i < len(courses):
    c = courses[i]
    is_curr_sem = False
    #ignore current and future semesters
    for sem in curr_sems:
        #checks to see if course is in a current/future semester
        if c[4] == sem:
            is_curr_sem = True
    if is_curr_sem:
        #courses from current semesters that end with roman numeral 'I'
        #will trigger a false positive, this will catch that
        if c[2] == 'I':
            c[-3] = 'In progress'
            course = c[1] + ' I'
            c[1] = course
        #Courses from current semester that have been dropped will not
        #be removed unless we add a check for it
        if c[2] == 'W':
            print(c)
            courses.pop(i)
            i -= 1
        i += 1
        continue
    #For certain courses, D is a failing grade
    elif c[0][:-4] == 'CSCI' or c[0][:-4] == "MATH" or \
         c[0] == "ENGL 110" or c[0] == "ENGL 111":
        if c[2] == 'F' or c[2] == 'U' or c[2] == 'N' or \
        c[2] == 'D' or c[2] == 'W' or c[2] == 'I' or \
        c[2] == 'In progress':
            print(c)
            courses.pop(i)
            #decrement or else loop will skip
            #elements after popping an element
            i -= 1
    #For all other courses, D is not a failing grade
    elif c[2] == 'F' or c[2] == 'U' or c[2] == 'N' or \
         c[2] == 'W' or c[2] == 'I' or c[2] == 'In progress':
        print(c)
        courses.pop(i)
        i -= 1
    i += 1
