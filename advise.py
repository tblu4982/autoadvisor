from preprocess import courses, fullname
import os
import re
from datetime import datetime

#at least 8 csci courses, 60 credits, and csci 287 before 400-level csci courses
#method to check if a course has a passing grade where a passing grade is 'C'
def passing_grade(grade, sem, curr_sem):
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

#method to check if a course has a passing grade where a passing grade is 'D'
def passing_grade_alt(grade, sem, curr_sem):
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

#Check for CSCI Elective
def csci_elec(course):
    #'CSCI 312', 'CSCI 361', 'CSCI 389', 'CSCI 396', 'CSCI 402', 'CSCI 450'
    #'CSCI 451', 'CSCI 453', 'CSCI 495', 'CSCI 496'
    elec = ['CSCI 312', 'CSCI 361', 'CSCI 389', 'CSCI 396', 'CSCI 402', \
            'CSCI 450', 'CSCI 451', 'CSCI 453', 'CSCI 495', 'CSCI 496']
    if course in elec:
        return True
    else:
        return False
    
#Check for Global Studies Elective
def gs_elec(course):
    #'SPAN', 'FREN', 'GERM', 'HIST 114', 'HIST 115', 'ARTS 199'
    elec = ['SPAN', 'FREN', 'GERM', 'HIST 114', 'HIST 115', 'ARTS 199']

    for i in elec:
        if re.search(i, course):
            return True
    return False

#Check for Social Sciences Elective
def ss_elec(course):
    #'SOCI', 'PSYC', 'CJUS 116', 'POLI 150', 'PHIL 314', GEOG 210
    #'MUSI 199', 'DRAM 199', 'SOCI 345', 'AGRI 295', 'ENGL 214'
    elec = ['SOCI', 'PSYC', 'CJUS 116', 'POLI 150', 'PHIL 314', \
            'GEOG 210', 'MUSI 199', 'DRAM 199', 'SOCI 345', 'AGRI 295', \
            'ENGL 214']
    for i in elec:
        if re.search(i, course):
            return True
    return False

#Check for CSCI/MATH/STAT Elective
def cms_elec(course):
    #any course that satisfies csci_elec or math_elec
    #'STAT 380'
    if csci_elec(course):
        return True
    if math_elec(course):
        return True
    if course == 'STAT 380':
        return True
    return False

#Check for Literature Elective
def lit_elec(course):
    #'ENGL 201', 'ENGL 202', 'ENGL 210', 'ENGL 211', 'ENGL 212', 'ENGL 213'
    #'ENGL 214', 'ENGL 215'
    elec = ['ENGL 201', 'ENGL 202', 'ENGL 210', 'ENGL 211', 'ENGL 212', \
            'ENGL 213', 'ENGL 214', 'ENGL 215']
    if course in elec:
        return True
    else:
        return False
    
#Check for History Elective
def hist_elec(course):
    #'HIST'
    if re.search('HIST', course):
        return True
    else:
        return False
    
#Check for Science Elective
def sci_elec(course):
    #'BIOL 120 + LAB', 'BIOL 121 + LAB', 'CHEM 151 + CHEM 153'
    #'PHYS 105 + LAB', 'PHYS 106 + LAB', 'PHYS 112 + LAB'
    #'PHYS 113 + LAB'
    elec = ['BIOL 120', 'BIOL 121', 'CHEM 151', 'CHEM 153', 'PHYS 105', \
            'PHYS 106', 'PHYS 112', 'PHYS 113']
    if course in elec:
        return True
    else:
        return False

def sci_check(lab, course):
    #All labs have the same ID as their course coutnerpart
    if lab == course:
        return True
    #special case for 'CHEM 151', whose lab has a different ID
    elif lab == 'CHEM 153' and course == 'CHEM 151':
        return True
    else:
        return False
#Check for Math Elective
def math_elec(course):
    #'MATH 325', 'MATH 340', 'MATH 348', 'MATH 350', 'MATH 360', 'MATH 425'
    elec = ['MATH 325', 'MATH 340', 'MATH 348', 'MATH 350', 'MATH 360', 'MATH 425']

    if course in elec:
        return True
    else:
        return False
    
#Finds the current semester
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

def print_to_file(path):
    with open(path, 'w') as file:
        file.write("------------------------------\n")
        file.write("ARTS Classes: " + str(arts) + "\n")
        file.write("CSCI Classes: " + str(csci) + "\n")
        file.write("Classes In Progress: " + str(in_prog) + "\n")
        file.write("Future Classes: " + str(pending) + "\n")
        file.write("Core Credits: " + str(core_credits) + "\n")
        file.write("Total Credits: " + str(total_credits) + "\n")
        file.write("------------------------------\n")
        file.write(math_stack + "\n------------------------------\n")
        file.write(csci_stack + "\n------------------------------\n")
        file.write(engl_stack + "\n------------------------------\n")
        file.write(misc_stack + "\n------------------------------\n")
        file.write(gs_stack + "\n------------------------------\n")
        file.write(lit_stack + "\n------------------------------\n")
        file.write(ss_stack + "\n------------------------------\n")
        file.write(csci_elec_stack + "\n------------------------------\n")
        file.write(math_elec_stack + "\n------------------------------\n")
        file.write(hist_stack + "\n------------------------------\n")
        file.write(cms_stack + "\n------------------------------\n")
        file.write(free_stack + "\n------------------------------\n")
        file.write(unused + "\n------------------------------\n")
        for n in recommendations:
            file.write(n + "\n")

#Dictionary to hold core course stack, as we traverse through course matrix,
#we'll set dict values to true as we find them
core_course_stack = {"MATH 260": False, "CSCI 101": False, "CSCI 150": False, "CSCI 151": False,
                     "ENGL 110": False, "MATH 261": False, "CSCI 250": False, "CSCI 251": False,
                     "ENGL 111": False, "MATH 280": False, "CSCI 287": False, "CSCI 303": False,
                     "ENGL 342": False, "CSCI 281": False, "CSCI 356": False, "CSCI 296": False,
                     "PHIL 450": False, "PHIL 275": False, "STAT 340": False, "CSCI 392": False,
                     "CSCI 487": False, "CSCI 445": False, "CSCI 358": False, "CSCI 489": False,
                     "CSCI 400": False, "CSCI 471": False, "CSCI 493": False, "CSCI 485": False,
                     "CSCI 494": False, "GEEN 310": False, "HPER 170": False}

elec_stack = {"SS": False, "HIST": False, "LIT": False, "SCI 1": False, "SCI 2": False,
              "MATH": False, "CMS": False, "CSCI 1": False, "CSCI 2": False, "GS": False,
              "FREE 1": False, "FREE 2": False}

#variable to hold the current semester
curr_sem = "Term: " + get_curr_sem()

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
        if c[4] == curr_sem:
            in_prog += 1
        else:
            pending += 1

    print(c[0],c[1],c[2],c[3],c[4])

#check course stacks
math_stack = "MATH: "
csci_stack = "CSCI: "
engl_stack = "ENGL: "
misc_stack = "MISC: "
gs_stack = "Global Studies: "
lit_stack = "Lit Elective: "
ss_stack = "Social Science:  "
csci_elec_stack = "CSCI Electives: "
math_elec_stack = "MATH Elective: "
sci_elec_stack = "Science Electives: "
hist_stack = "HIST Elective: "
cms_stack = "CSCI/MATH/STAT Elective: "
free_stack = "Free Electives: "
unused = "Unused Courses: "
#holds names of science classes so that we can compare to labs
sci = []
sci_lab = []

#loop through matrix to find courses matching requirements
for c in courses:
    passed_course = False
    course = c[0]
    grade = c[2]
    credit = float(c[3])
    term = c[4]
    abbr = course[:5]
    #search for core class requirements first
    if course in core_course_stack:
        #for certain courses, 'D' is a passing grade
        if course == "PHIL 450" or course == "PHIL 275" or course == "ENGL 342":
            passed_course = passing_grade_alt(grade, term, curr_sem)
        #for all other courses, 'C' is the minimum
        else:
            passed_course = passing_grade(grade, term, curr_sem)
        if passed_course:
            core_course_stack[course] = True
            if abbr == "MATH ":
                math_stack += course + " "
            if abbr == "CSCI ":
                csci_stack += course + " "
            if abbr == "ENGL ":
                engl_stack += course + " "
            if abbr == "PHIL ":
                misc_stack += course + " "
    #for non-core courses
    else:
        #check global studies elective
        if gs_elec(course) and not elec_stack["GS"]:
            elec_stack["GS"] = True
            gs_stack += course
        #check literature elective
        elif lit_elec(course) and not elec_stack["LIT"]:
            elec_stack["LIT"] = True
            lit_stack += course
        #check social science elective
        elif ss_elec(course) and not elec_stack["SS"]:
            elec_stack["SS"] = True
            ss_stack += course
        #check csci electives
        elif csci_elec(course) and not elec_stack["CSCI 2"]:
            if elec_stack["CSCI 1"]:
                elec_stack["CSCI 2"] = True
                csci_elec_stack += course
            else:
                elec_stack["CSCI 1"] = True
                csci_elec_stack += course + " "
        #check math elective
        elif math_elec(course) and not elec_stack["MATH"]:
            elec_stack["MATH"] = True
            math_elec_stack += course
        #check history elective
        elif hist_elec(course) and not elec_stack["HIST"]:
            elec_stack["HIST"] = True
            hist_stack += course
        #check science electives
        elif sci_elec(course):
            #check if course is lab class
            if credit == 1:
                #store lab course to an array to compare later
                if not course in sci_lab:
                    sci_lab.append(course)
                #search for main course
                for itr in sci:
                    #if main course matches lab, then science is satisfied
                    if sci_check(course, itr):
                        if not elec_stack["SCI 1"]:
                            sci_elec_stack += course + " "
                            elec_stack["SCI 1"] = True
                        elif not elec_stack["SCI 2"]:
                            sci_elec_stack += course
                            elec_stack["SCI 2"] = True
            #If not a lab class, then course must be main course
            else:
                #set main course to array to compare later
                if not course in sci:
                    sci.append(course)
                #search for matching lab course
                for itr in sci_lab:
                    if sci_check(itr, course):
                        if not elec_stack["SCI 1"]:
                            sci_elec_stack += course + " "
                            elec_stack["SCI 1"] = True
                        elif not elec_stack["SCI 2"]:
                            sci_elec_stack += course
                            elec_stack["SCI 2"] = True
        #csci/math/stat elective
        elif cms_elec(course) and not elec_stack["CMS"]:
            elec_stack["CMS"] = True
            cms_stack += course
        #unused courses
        else:
            #exclude courses with 0 credits
            if not elec_stack["FREE 1"] and credit > 0:
                elec_stack["FREE 1"] = True
                free_stack += course + " "
            elif not elec_stack["FREE 2"] and credit > 0:
                elec_stack["FREE 2"] = True
                free_stack += course
            else:
                unused += course + " "
            
#list to hold course recommendations
recommendations = []

#Math courses stack
if core_course_stack["MATH 260"]:
    if core_course_stack["MATH 261"]:
        if core_course_stack["CSCI 281"] and not core_course_stack["STAT 340"]:
            recommendations.append("Recommend STAT 340")
    else:       
        recommendations.append("Recommend MATH 261")
else:
    recommendations.append("Recommend MATH 260")

if not core_course_stack["MATH 280"]:
    recommendations.append("Recommend MATH 280")

#CSCI courses stack
if not core_course_stack["CSCI 101"]:
    recommendations.append("Recommend CSCI 101")

if not core_course_stack["CSCI 150"]:
    recommendations.append("Recommend CSCI 150")

if not core_course_stack["CSCI 151"]:
    recommendations.append("Recommend CSCI 151")

if core_course_stack["MATH 280"]:
    if core_course_stack["CSCI 281"]:
        if core_course_stack["CSCI 287"] and not core_course_stack["CSCI 392"]:
            recommendations.append("Recommend CSCI 392")
    else:
        recommendations.append("Recommend CSCI 281")

if core_course_stack["CSCI 150"] and core_course_stack["CSCI 151"]:
    if not core_course_stack["CSCI 250"]:
        recommendations.append("Recommend CSCI 250")

    if not core_course_stack["CSCI 251"]:
        recommendations.append("Recommend CSCI 251")

    if core_course_stack["CSCI 250"] and core_course_stack["CSCI 251"]:
        if not core_course_stack["CSCI 296"]:
            recommendations.append("Recommend CSCI 296")

        if not core_course_stack["CSCI 356"]:
            recommendations.append("Recommend CSCI 356")

        if not core_course_stack["CSCI 358"]:
            recommendations.append("Recommend CSCI 358")

        if core_course_stack["CSCI 287"]:
            if not core_course_stack["CSCI 487"]:
                recommendations.append("Recommend CSCI 487")
            if not core_course_stack["CSCI 471"]:
                recommendations.append("Recommend CSCI 471")
            if not core_course_stack["CSCI 485"]:
                recommendations.append("Recommend CSCI 485")
        else:
            recommendations.append("Recommend CSCI 287")

        if core_course_stack["CSCI 303"]:
            if not core_course_stack["CSCI 489"]:
                recommendations.append("Recommend CSCI 489")
            if not core_course_stack["CSCI 445"]:
                recommendations.append("Recommend CSCI 445")
        else:
            recommendations.append("Recommend CSCI 303")

if core_course_stack["CSCI 400"]:
    if core_course_stack["CSCI 493"]:
        if not core_course_stack["CSCI 494"]:
            recommendations.append("Recommend CSCI 494")
    else:
        recommendations.append("Recommend CSCI 493")
elif total_credits > 75:
    recommendations.append("Recommend CSCI 400")

#English courses stack
if core_course_stack["ENGL 110"]:
    if core_course_stack["ENGL 111"]:
        if not core_course_stack["ENGL 342"]:
            recommendations.append("Recommend ENGL 342")
    else:
        recommendations.append("Recommend ENGL 111")
else:
    recommendations.append("Recommend ENGL 110")

#Miscellaneous courses stack
if not core_course_stack["PHIL 450"] or core_course_stack["PHIL 275"]:
    recommendations.append("Recommend PHIL 450 or PHIL 275")

#write advisory report to file
path = "students\\" + fullname + "\\advise.txt"
#check if file already exists
if os.path.exists(path):
    #if it exists, update it
    os.remove(path)
    print_to_file(path)
#if not, create it
else:
    print_to_file(path)
