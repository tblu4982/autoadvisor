from preprocess import courses
import re
from datetime import datetime

#90 total credits and 72 core class credits before csci 400 and past csci 287
#students become eligible for 400 level courses if they have 75+ credits

#method to check if a course has a passing grade where a passing grade is 'C'
def passing_grade(grade, sem, curr_sems):
    if grade <= "C" or grade == "S" or grade == "TR" or grade == "In progress" \
       or grade == "SP":
        #add a check for future courses, which should be omitted
        #from course recommendations
        #If course is from this semester, include it
        if grade == "In progress":
            if sem == curr_sems:
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

#method to check if a course has a passing grade where a passing grade is 'D'
def passing_grade_alt(grade, sem, curr_sems):
    if grade <= "D" or grade == "S" or grade == "TR" or grade == "In progress" \
       or grade == "SP":
        if grade == "In progress":
            if sem == curr_sems:
                return True
            else:
                return False
        else:
            return True
    else:
        return False

def csci_free_elec(course):
    pass

def gs_elec(course):
    pass

def ss_elec(course):
    pass

def cms_elec(course):
    pass

def lit_elec(course):
    pass

def hist_elec(course):
    pass

def bcp_elec(course):
    pass

def math_elec(course):
    pass
    

def get_curr_sem():
    month = datetime.now().month
    year = datetime.now().year

    if month >= 1:
        if month >= 5:
            if month <= 8:
                if month == 12:
                    return "Winter " + str(year)
                else:
                    return "Fall " + str(year)
            else:
                return "Summer " + str(year)
        else:
            return "Spring " + str(year)

#Dictionary to hold core course stack, as we traverse through course matrix,
#we'll set dict values to true as we find them
core_course_stack = {"MATH 260": False, "CSCI 101": False, "CSCI 150": False, "CSCI 151": False,
                     "ENGL 110": False, "MATH 261": False, "CSCI 250": False, "CSCI 251": False,
                     "ENGL 111": False, "MATH 280": False, "CSCI 287": False, "CSCI 303": False,
                     "ENGL 342": False, "CSCI 281": False, "CSCI 356": False, "CSCI 296": False,
                     "PHIL 450": False, "PHIL 275": False, "STAT 340": False, "CSCI 392": False,
                     "CSCI 487": False, "CSCI 445": False, "CSCI 358": False, "CSCI 489": False,
                     "CSCI 400": False, "CSCI 471": False, "CSCI 493": False, "CSCI 485": False,
                     "CSCI 494": False}    

#array to store the current/future semesters
curr_sems = []

#print courses
#course, name, grade, credits, semester
print("------------------------------")
csci = 0
arts = 0
in_prog = 0
pending = 0
core_credits = 0
total_credits = 0

for c in courses:
    total_credits += float(c[3])
    if c[0] in core_course_stack:
        core_credits += float(c[3])
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

if len(curr_sems) == 0:
    curr_sems.append(get_curr_sem())


print("------------------------------")    
print("ARTS Classes: ", str(arts))
print("CSCI Classes: ", str(csci))
print("Classes In Progress: ", str(in_prog))
print("Future Classes: ", str(pending))
print("Core Credits: ", str(core_credits))
print("Total Credits: ", str(total_credits))
print("------------------------------")

#check course stacks

#loop through matrix to find courses matching requirements
for c in courses:
    passed_course = False
    if c[0] in core_course_stack:
        if c[0] == "PHIL 450" or c[0] == "PHIL 275" or c[0] == "ENGL 342":
            passed_course = passing_grade_alt(c[2], c[4], curr_sems)
        else:
            passed_course = passing_grade(c[2], c[4], curr_sems)
        if passed_course:
            core_course_stack[c[0]] = True
            
#list to hold course recommendations
recommendations = []

#Math courses stack
print("MATH:", end= " ")
if core_course_stack["MATH 260"]:
    print("MATH 260", end = " ")
    if core_course_stack["MATH 261"]:
        print(", MATH 261", end = " ")
        if core_course_stack["STAT 340"] and core_course_stack["CSCI 281"]:
            print(", STAT 340", end = " ")
        elif core_course_stack["CSCI 281"]:
            recommendations.append("Recommend STAT 340")
        else:
            print("\nNeed CSCI 281 before STAT 340")
    else:       
        recommendations.append("Recommend MATH 261")
else:
    recommendations.append("Recommend MATH 260")

if core_course_stack["MATH 280"]:
    print("MATH 280", end = " ")

print("\n------------------------------")

#CSCI courses stack
print("CSCI:", end= " ")
if core_course_stack["CSCI 101"]:
    print("CSCI 101", end= " ")
else:
    recommendations.append("Recommend CSCI 101")

if core_course_stack["CSCI 150"]:
    print("CSCI 150", end= " ")
else:
    recommendations.append("Recommend CSCI 150")

if core_course_stack["CSCI 151"]:
    print("CSCI 151", end= " ")
else:
    recommendations.append("Recommend CSCI 151")

if core_course_stack["MATH 280"]:
    if core_course_stack["CSCI 281"]:
        print("CSCI 281", end= " ")
    else:
        recommendations.append("Recommend CSCI 281")
        if core_course_stack["CSCI 392"] and core_course_stack["CSCI 287"]:
            print("CSCI 392", end= " ")
        elif core_course_stack["CSCI 287"]:
            recommendations.append("Recommend CSCI 392")
        else:
            print("\nNeed CSCI 287 before CSCI 392")

if core_course_stack["CSCI 150"] and core_course_stack["CSCI 151"]:
    if core_course_stack["CSCI 250"]:
        print("CSCI 250", end= " ")
    else:
        recommendations.append("Recommend CSCI 250")

    if core_course_stack["CSCI 251"]:
        print("CSCI 251", end= " ")
    else:
        recommendations.append("Recommend CSCI 251")

    if core_course_stack["CSCI 250"] and core_course_stack["CSCI 251"]:
        if core_course_stack["CSCI 296"]:
            print("CSCI 296", end= " ")
        else:
            recommendations.append("Recommend CSCI 296")

        if core_course_stack["CSCI 356"]:
            print("CSCI 356", end= " ")
        else:
            recommendations.append("Recommend CSCI 356")

        if core_course_stack["CSCI 358"]:
            print("CSCI 358", end= " ")
        else:
            recommendations.append("Recommend CSCI 358")

        if core_course_stack["CSCI 287"]:
            print("CSCI 287", end= " ")
            if core_course_stack["CSCI 487"]:
                print("CSCI 487", end= " ")
            else:
                recommendations.append("Recommend CSCI 487")
            if core_course_stack["CSCI 471"]:
                print("CSCI 471", end= " ")
            else:
                recommendations.append("Recommend CSCI 471")
            if core_course_stack["CSCI 485"]:
                print("CSCI 485", end= " ")
            else:
                recommendations.append("Recommend CSCI 485")
        else:
            recommendations.append("Recommend CSCI 287")

        if core_course_stack["CSCI 303"]:
            print("CSCI 303", end= " ")
            if core_course_stack["CSCI 489"]:
                print("CSCI 489", end= " ")
            else:
                recommendations.append("Recommend CSCI 489")
            if core_course_stack["CSCI 445"]:
                print("CSCI 445", end= " ")
            else:
                recommendations.append("Recommend CSCI 445")
        else:
            recommendations.append("Recommend CSCI 303")

if core_course_stack["CSCI 400"]:
    print("CSCI 400", end= " ")
    if core_course_stack["CSCI 493"]:
        print("CSCI 493", end= " ")
        if core_course_stack["CSCI 494"]:
            print("CSCI 494", end= " ")
        else:
            recommendations.append("Recommend CSCI 494")
    else:
        recommendations.append("Recommend CSCI 493")
elif total_credits > 75:
    recommendations.append("Recommend CSCI 400")

print("\n------------------------------")

#English courses stack
print("ENGL:", end= " ")

if core_course_stack["ENGL 110"]:
    print("ENGL 110", end= " ")
    if core_course_stack["ENGL 111"]:
        print("ENGL 111", end= " ")
        if core_course_stack["ENGL 342"]:
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

if core_course_stack["PHIL 450"] or core_course_stack["PHIL 275"]:
    if core_course_stack["PHIL 450"]:
        print("PHIL 450", end= " ")
    else:
        print("PHIL 275", end= " ")
else:
    recommendations.append("Recommend PHIL 450 or PHIL 275")

#print course recommendations
print('\n------------------------------')
for n in recommendations:
    print(n)

