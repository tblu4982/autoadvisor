from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import tkinter #GUI
from tkinter import *
from tkinter import messagebox
import os #used for read/write to files
import re #used for regular expressions
import sys #used to stop execution under certain circumstances

#Class to hold user login info for session
class credentials():
    #Opens a window and prompts user to select if they are student or advisor
    def greeting_window (self, greeting):
        user = Tk()
        user.geometry("200x70")
        user.title("Auto Advisor")
        welcome = Label(user, text = greeting)
        prompt = Label(user, text='Are you an Advisor or a Student?')
        #-----------FIX ERROR HERE: AUTOMATICALLY TRIGGERS ADVISOR ON LAUNCH----------
        advisor = Button(user, text="Advisor", command = self.click_advisor())
        #-----------FIX ERROR HERE: AUTOMATICALLY TRIGGERS ADVISOR ON LAUNCH----------
        student = Button(user, text="Student", command = lambda:[self.click_student(), user.destroy()])

        welcome.pack(side = TOP)
        prompt.pack(side = TOP)
        advisor.pack(side = LEFT, expand = True, fill = BOTH)
        student.pack(side = LEFT, expand = True, fill = BOTH)
        user.mainloop()

    #method to grab student login info
    def click_student(self):
        print("test")
        self.auth = "student"
        #opens a window to grab student login info
        s_portal = Tk()
        s_portal.geometry("200x70")
        s_portal.title("Auto Advisor")
        Label(s_portal, text='User ID').grid(row=0)
        Label(s_portal, text='Password').grid(row=1)
        uid = Entry(s_portal)
        pwd = Entry(s_portal)
        #when submit button is clicked, it sends credentials to Banner Portal
        submit = Button(s_portal, text='Submit', command = lambda:[self.set_credentials(uid, pwd), s_portal.destroy()])
        uid.grid(row = 0, column = 1)
        pwd.grid(row = 1, column = 1)
        submit.grid(row = 3, column = 1)
        
    #sets user's login info
    def set_credentials(self, uid, pwd):
        self.username = uid.get()
        self.password = pwd.get()
        
    #returns login info
    def get_credentials(self):
        return self.auth, self.username, self.password
    
    #method to grab advisor login info
    def click_advisor(self):
        auth = "advisor"
        print("Advisor", "You clicked 'Advisor'")

    #method that prompts user to correct invalid login info
    def invalid_creds(self):
        self.greeting = 'Invalid login credentials!'
        self.greeting_window(self.greeting)

    #instantiates class
    def __init__(self):
        username = " "
        password = " "
        auth = " "
        greeting = 'Welcome to Auto-Advisor!'
        self.greeting_window(greeting)

#method to return person's name
def get_name(name):
    for i in name:
        if i.text[0:7] == "Welcome":
            fullname = i.text.split(",")[1].replace(" ", "").replace(".", "")
            return fullname
        
#method to create file to store data
def create_file_path(fullname ,path, sp_path, file_type, data):
    #check if file directory exists
    if os.path.exists(path):
        #if folder exists, check if file already exists
        file_path = path + sp_path
        #if file exists, overwrite it
        if os.path.exists(file_path):
            #remove old file, then create new file
            print("Overwriting " + file_type + " file for " + fullname + "...")
            os.remove(file_path)
            with open(file_path, 'w') as target_file:
                if file_type == "courses":
                    for line in data:
                        for index in line:
                            target_file.write(index + " ")
                        target_file.write("\n")
                else:
                    for line in data:
                        target_file.write(line + "\n")
        else:
            #create new file
            print("Creating " + file_type + " file for " + fullname + "...")
            with open(file_path, 'w') as target_file:
                if file_type == "courses":
                    for line in data:
                        for index in line:
                            target_file.write(index + " ")
                        target_file.write("\n")
                else:
                    for line in data:
                        target_file.write(line + "\n")
    else:
        print("Error! Folder path for " + fullname + " does not exist!")

#Gets user login credentials
user = credentials()
try:
    auth, username, password = user.get_credentials()
#Exception Handling that closes program if tkinter box is closed prematurely
except AttributeError:
    sys.exit("Program Terminated")

#---------------USE TKINTER TO GET DESTINATION FOLDER----------------
    
driver = webdriver.Edge(r'C:\Users\terre\Documents\edgedriver_win64\msedgedriver.exe')
driver.get('https://ssb-prod.ec.vsu.edu/BNPROD/twbkwbis.P_WWWLogin')

#enter credentials into Banner
uid = driver.find_element_by_id("UserID")
pwd = driver.find_element_by_id("PIN")

uid.send_keys(username)
pwd.send_keys(password)

driver.find_element_by_xpath("//form").submit()

#--------ADD LOGIC THAT CHECKS TO SEE IF USER IS AT LANDING PAGE------
#INVALID LOGIN URL: https://ssb-prod.ec.vsu.edu/BNPROD/twbkwbis.P_ValLogin
#try:
    #auth, username, password = user.invalid_creds()
#Exception Handling that closes program if tkinter box is closed prematurely
#except AttributeError:
    #sys.exit("Program Terminated")
#--------ADD LOGIC THAT CHECKS TO SEE IF USER IS AT LANDING PAGE------

#---------IMPLEMENT UNIQUE CRAWLING LOGIC FOR STUDENT/ADVISOR----------
if auth == "advisor":
    pass

if auth == "student":
    pass
#---------IMPLEMENT UNIQUE CRAWLING LOGIC FOR STUDENT/ADVISOR----------

#grab person's name to use for file name
name = driver.find_elements_by_xpath("//td[@class='pldefault']")
fullname = get_name(name)

#crawl webpage to academic transcript
driver.find_element_by_link_text("Student").click()
driver.find_element_by_link_text("Student Records").click()
driver.find_element_by_link_text("Academic Transcript").click()

levl_id = Select(driver.find_element_by_id("levl_id"))
type_id = Select(driver.find_element_by_id("type_id"))

#create folder using student name
#search to see if student folder already exists
#-----------------USE DESTINATION FOLDER FROM ABOVE TO SET PATH---------------
path = "C:/Users/terre/Documents/Senior Project/AutoAdvisor/students/" + fullname
#-----------------USE DESTINATION FOLDER FROM ABOVE TO SET PATH---------------
if not os.path.exists(path):
    os.makedirs(path)
    if os.path.exists(path):
        print("Path successfully created!")
    else:
        print("Failed to create path!")
else:
    print("Path already exists!")

driver.find_element_by_xpath("//table[contains(@class, 'dataentrytable')]").submit()

#scrape transcript for semester headers
output = driver.find_elements_by_xpath("//span[contains(@class, 'fieldOrangetextbold')]")

semesters = []
for i in output:
    #check if semester is properly structured
    if i.text[:6] == "Term: ":
        semesters.append(i.text)
        semesters.append(" ")
    #structure semester to add missing information
    else:
        semesters.append("Term: " + i.text)
        semesters.append(" ")

#print semesters to file
create_file_path(fullname , path, "/semesters.txt", "semesters", semesters)

#scrape transcript for courses
output = driver.find_elements_by_xpath("//td[contains(@class, 'dddefault')]")
#stores courses as a course matrix
courses = []
#used to construct course line from web data
proto = []
#used to determine if we reached the end of a semester
counter = 0
#used to determine if we reached current/future semesters
is_curr = False

for i in output:
    #1st index in course should be course abbreviation ID
    if len(i.text) == 4:
        proto.append(i.text)
    #if we reach this, then we are in between courses
    elif len(proto) == 0:
        #If there are no courses in matrix, then no need to increment
        if len(courses) == 0:
            continue
        else:
            counter += 1
            continue
    #if we reach this, then we have reached the end of a course
    elif re.search("[0-9].000", i.text):
        #if counter is at least 6, then we have likely passed the end
        #of a semester
        if counter >= 6 and len(courses) > 0:
            #At least 18 elements are popped before we reach current/future semesters
            if counter > 18:
                is_curr = True
            courses.append("-")
        counter = 0
        #add a marker for current/future courses in place of missing grade
        if is_curr:
            proto.append("inprog")
        #append course to course matrix
        proto.append(i.text)
        course = proto.copy()
        #remove "U" from course
        if course[2] == "U":
            course.pop(2)
        courses.append(course)
        proto.clear()
    #if we reach this, then we are not at the end of a course
    else:
        proto.append(i.text)

#append dash for final course structure
courses.append("-")
    
#print courses to file
create_file_path(fullname ,path, "/courses.txt", "courses", courses)

driver.quit()
