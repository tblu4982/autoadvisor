import advise

def main(fullname, config_file):
    for name in fullname:    
        f1 = open("students\\" + name + "\\" + config_file.split('/')[-1].split('.')[0] + "\courses.txt", "r")
        f2 = open("students\\" + name + "\\" + config_file.split('/')[-1].split('.')[0] + "\semesters.txt", "r")

        print("Formatting Transcript for " + name + "...")

        #Read semesters from file and insert it into list semesters
        semesters = []
        for line in f2:
            if len(line) > 2:
                line=line.strip()
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

        #Adjust courses to have course title as one element in list
        #For example ['CHEM', '152', 'General', 'Chemistry', 'II', 'S', '0.000'] becomes
        #['CHEM', '152', 'General Chemistry II ', 'S', '0.000']
        j=0
        in_prog = False
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
                courses[i].append(s)
                i += 1
            i += 1

        #Then remove dashes
        for c in courses:
            if c[0] == "-":
                courses.remove(c)

        #join course acronym and number together
        for c in courses:
            course = c[0]+ " " + c[1]
            c[0] = course
            c.pop(1)

        #Set semester and Replace 'inprog' with 'In progress',
        for c in courses:
            if c[2] == 'inprog':
                c[2] = 'In progress'

        #remove failed courses
        #course, name, grade, credits, semester    
        i = 0
        while i < len(courses):
            c = courses[i]
            course = c[0]
            abbr = course[:-4]
            grade = c[2]
            if grade == 'F' or grade == 'U' or grade == 'N' or \
                 grade == 'W' or grade == 'I':
                courses.pop(i)
                i -= 1
            i += 1

        #remove duplicate courses
        i = 0
        while i < len(courses):
            j = i + 1
            c1 = courses[i][0]
            abbr = c1[:-4]
            if not abbr == 'BIOL':
                if not abbr == 'PHYS':
                    if not abbr == 'CHEM':
                        while j < len(courses):
                            c2 = courses[j][0]
                            if c1 == c2:
                                courses.pop(i)
                                break
                            j += 1
            i += 1
                    
        #Pass name(string) and courses(list) to advise.py
        #advise.main(courses, name)
        advise.main(courses, name, config_file)
