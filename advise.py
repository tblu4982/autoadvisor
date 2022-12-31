from preprocess import courses
import re

#checks csci courses for passing grade in terms eligibility
def csci_course_check (grade, sem, curr_sem):
    if grade <= "C" or grade == "S" or grade == "TR" or grade == "In progress" \
       or grade == "SP":
        #add a check for future courses, which should be omitted
        #from course recommendations
        #If course is from this semester, include it
        if grade == "In progress":
            if sem == curr_sem:
                return True
            #If course is from future semester, omit it
            else:
                return False
        #If course is not from this semester, then it has already
        #been taken
        else:
            return True
    #If we reached this, then we reached a failing course somehow
    else:
        return False

#checks other courses for passing grade in terms of eligibility
def course_check (grade, sem, curr_sem):
    if grade <= "D" or grade == "S" or grade == "TR" or grade == "In progress" \
       or grade == "SP":
        if grade == "In progress":
            if sem == curr_sem:
                return True
            else:
                return False
        else:
            return True
    else:
        return False

#array to store the current/future semesters
curr_sems = []

#print courses
#course, name, grade, credits, semester
print("------------------------------")
csci = 0
arts = 0
in_prog = 0
pending = 0
credit_hours = 0

for c in courses:
    credit_hours += float(c[3])
    if re.search("ARTS", c[0]):
        arts += 1
    elif re.search("CSCI", c[0]):
        csci += 1
        
    if c[2] == "In progress":
        if len(curr_sems) == 0:
            curr_sems.append(c[4])
        if c[4] == curr_sems[0]:
            in_prog += 1
        else:
            pending += 1

    print(c[0],c[1],c[2],c[3],c[4])


print("------------------------------")    
print("ARTS Classes: ", str(arts))
print("CSCI Classes: ", str(csci))
print("Classes In Progress: ", str(in_prog))
print("Future Classes: ", str(pending))
print("------------------------------")

#check course stacks
#course, name, grade, credits, semester
#boolean values that hold whether or not a course was passed
m120 = m121 = m260 = m261 = m280 = s340 = False
cs101 = cs150 = cs151 = cs250 = cs251 = False
cs281 = cs287 = cs296 = cs303 = cs356 = False
cs358 = cs392 = cs400 = cs445 = cs471 = False
cs485 = cs487 = cs489 = cs493 = cs494 = False
e110 = e111 = e342 = phil = False
phil_n = ""

#loop through matrix to find courses matching requirements
for c in courses:
    if c[0] == "MATH 260" or m260 and not s340:
        if course_check (c[2], c[4], curr_sems[0]):
            m260 = True
            if c[0] == "MATH 261" or m261:
                if course_check (c[2], c[4], curr_sems[0]):
                    m261 = True
                    if c[0] == "STAT 340" and cs281:
                        if course_check (c[2], c[4], curr_sems[0]):
                            s340 = True

    if c[0] == "MATH 120" and not m120 and not m260:
        if course_check (c[2], c[4], curr_sems[0]):
            m120 = True
            
    if c[0] == "MATH 121" and not m280 and not m260:
        if course_check (c[2], c[4], curr_sems[0]):
            m121 = True
            
    if c[0] == "MATH 280" and m121 or m280 or not m121 and m260 and not cs392:
        if course_check (c[2], c[4], curr_sems[0]):
            m280 = True
            if c[0] == "CSCI 281" or cs281:
                if csci_course_check (c[2], c[4], curr_sems[0]):
                    cs281 = True
                    if c[0] == "CSCI 392" and cs287:
                        if csci_course_check (c[2], c[4], curr_sems[0]):
                            cs392 = True

    
                            
    if c[0] == "CSCI 101" and not cs101:
        if csci_course_check (c[2], c[4], curr_sems[0]):
            cs101 = True

    if c[0] == "CSCI 150" and not cs150:
        if csci_course_check (c[2], c[4], curr_sems[0]):
            cs150 = True

    if c[0] == "CSCI 151" and not cs151:
        if csci_course_check (c[2], c[4], curr_sems[0]):
            cs151 = True

    if c[0] == "CSCI 250" and cs150 and cs151 and not cs250:
        if csci_course_check (c[2], c[4], curr_sems[0]):
            cs250 = True

    if c[0] == "CSCI 251" and cs150 and cs151 and not cs251:
        if csci_course_check (c[2], c[4], curr_sems[0]):
            cs251 = True

    if c[0] == "CSCI 296" and cs250 and cs251 and not cs296:
        if csci_course_check (c[2], c[4], curr_sems[0]):
            cs296 = True

    if c[0] == "CSCI 356" and cs250 and cs251 and not cs356:
        if csci_course_check (c[2], c[4], curr_sems[0]):
            cs356 = True

    if c[0] == "CSCI 358" and cs250 and cs251 and not cs358:
        if csci_course_check (c[2], c[4], curr_sems[0]):
            cs358 = True

    if c[0] == "CSCI 287" or cs287 and cs250 and cs251:
        if csci_course_check (c[2], c[4], curr_sems[0]):
            cs287 = True
            if c[0] == "CSCI 487":
                if csci_course_check (c[2], c[4], curr_sems[0]):
                    cs487 = True
            if c[0] == "CSCI 471":
                if csci_course_check (c[2], c[4], curr_sems[0]):
                    cs471 = True
            if c[0] == "CSCI 485":
                if csci_course_check (c[2], c[4], curr_sems[0]):
                    cs485 = True

    if c[0] == "CSCI 303" or cs303 and cs250 and 251:
        if csci_course_check (c[2], c[4], curr_sems[0]):
            cs303 = True
            if c[0] == "CSCI 489":
                if csci_course_check (c[2], c[4], curr_sems[0]):
                    cs489 = True
            if c[0] == "CSCI 445":
                if csci_course_check (c[2], c[4], curr_sems[0]):
                    cs445 = True

    if c[0] == "CSCI 400" or cs400 and credit_hours > 75 and not cs494:
        if csci_course_check (c[2], c[4], curr_sems[0]):
            cs400 = True
            if c[0] == "CSCI 493" or cs493:
                if csci_course_check (c[2], c[4], curr_sems[0]):
                    cs493 = True
                    if c[0] == "CSCI 494":
                        cs494 = True

    if c[0] == "ENGL 110" or e110 and not e342:
        if course_check (c[2], c[4], curr_sems[0]):
            e110 = True
            if c[0] == "ENGL 111" or e111:
                if course_check (c[2], c[4], curr_sems[0]):
                    e111 = True
                    if c[0] == "ENGL 342":
                        if course_check (c[2], c[4], curr_sems[0]):
                            e342 = True

    if c[0] == "PHIL 450" or c[0] == "PHIL 275" and not phil:
        if course_check (c[2], c[4], curr_sems[0]):
            phil_n = c[0]
            phil = True
#list to hold course recommendations
recommendations = []

#Math courses stack
print("MATH:", end= " ")
if m260:
    print("MATH 260", end = " ")
    if m261:
        print(", MATH 261", end = " ")
        if s340 and cs281:
            print(", STAT 340", end = " ")
        elif cs281:
            recommendations.append("Recommend STAT 340")
        else:
            print("\nNeed CSCI 281 before STAT 340")
    else:       
        recommendations.append("Recommend MATH 261")
else:
    recommendations.append("Recommend MATH 260")

if m280:
    print("MATH 280", end = " ")

print("\n------------------------------")

#CSCI courses stack
print("CSCI:", end= " ")
if cs101:
    print("CSCI 101", end= " ")
else:
    recommendations.append("Recommend CSCI 101")

if cs150:
    print("CSCI 150", end= " ")
else:
    recommendations.append("Recommend CSCI 150")

if cs151:
    print("CSCI 151", end= " ")
else:
    recommendations.append("Recommend CSCI 151")

if m280:
    if cs281:
        print("CSCI 281", end= " ")
    else:
        recommendations.append("Recommend CSCI 281")
        if cs392 and cs287:
            print("CSCI 392", end= " ")
        elif cs287:
            recommendations.append("Recommend CSCI 392")
        else:
            print("\nNeed CSCI 287 before CSCI 392")

if cs150 and cs151:
    if cs250:
        print("CSCI 250", end= " ")
    else:
        recommendations.append("Recommend CSCI 250")

    if cs251:
        print("CSCI 251", end= " ")
    else:
        recommendations.append("Recommend CSCI 251")

    if cs250 and cs251:
        if cs296:
            print("CSCI 296", end= " ")
        else:
            recommendations.append("Recommend CSCI 296")

        if cs356:
            print("CSCI 356", end= " ")
        else:
            recommendations.append("Recommend CSCI 356")

        if cs358:
            print("CSCI 358", end= " ")
        else:
            recommendations.append("Recommend CSCI 358")

        if cs287:
            print("CSCI 287", end= " ")
            if cs487:
                print("CSCI 487", end= " ")
            else:
                recommendations.append("Recommend CSCI 487")
            if cs471:
                print("CSCI 471", end= " ")
            else:
                recommendations.append("Recommend CSCI 471")
            if cs485:
                print("CSCI 485", end= " ")
            else:
                recommendations.append("Recommend CSCI 485")
        else:
            recommendations.append("Recommend CSCI 287")

        if cs303:
            print("CSCI 303", end= " ")
            if cs489:
                print("CSCI 489", end= " ")
            else:
                recommendations.append("Recommend CSCI 489")
            if cs445:
                print("CSCI 445", end= " ")
            else:
                recommendations.append("Recommend CSCI 445")
        else:
            recommendations.append("Recommend CSCI 303")

if cs400:
    print("CSCI 400", end= " ")
    if cs493:
        print("CSCI 493", end= " ")
        if cs494:
            print("CSCI 494", end= " ")
        else:
            recommendations.append("Recommend CSCI 494")
    else:
        recommendations.append("Recommend CSCI 493")
elif credit_hours > 75:
    recommendations.append("Recommend CSCI 400")

print("\n------------------------------")

#English courses stack
print("ENGL:", end= " ")

if e110:
    print("ENGL 110", end= " ")
    if e111:
        print("ENGL 111", end= " ")
        if e342:
            print("ENGL 342", end= " ")
        else:
            recommendations.append("Recommend ENGL 342")
    else:
        recommendations.append("Recommend ENGL 111")
else:
    recommendations.append("Recommend ENGL 110")

print("\n------------------------------")

#Miscellaneous courses stack
print("MISC:", end = " ")

if phil:
    print(phil_n, end= " ")
else:
    recommendations.append("Recommend PHIL 450 or PHIL 275")

#print course recommendations
print('\n------------------------------')
for n in recommendations:
    print(n)
