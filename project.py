from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import tkinter #GUI
from tkinter import *
from tkinter import messagebox
from datetime import datetime
import os #used for read/write to files
import re #used for regular expressions
import sys #used to stop execution under certain circumstances

#Class to hold user login info for session
class credentials():
    #Opens a window and prompts user to select if they are student or advisor
    def greeting_window (self, greeting):
        portal = Tk()
        portal.geometry("200x70")
        portal.title("Auto Advisor")
        welcome = Label(portal, text = greeting)
        prompt = Label(portal, text='Are you an Advisor or a Student?')
        advisor = Button(portal, text="Advisor", command = lambda:[portal.destroy(), self.click_advisor()])
        student = Button(portal, text="Student", command = lambda:[portal.destroy(), self.click_student()])

        welcome.pack(side = TOP)
        prompt.pack(side = TOP)
        advisor.pack(side = LEFT, expand = True, fill = BOTH)
        student.pack(side = LEFT, expand = True, fill = BOTH)
        portal.mainloop()

    #method to grab student login info
    def click_student(self):
        print("Clicked 'Student'")
        self.auth = "student"
        #opens a window to grab student login info
        login = Tk()
        login.geometry("200x90")
        login.title("Auto Advisor - Student")
        welcome = Label(login, text = "Loggin in as: Student")
        Label(login, text='User ID').grid(row=1)
        Label(login, text='Password').grid(row=2)
        uid = Entry(login)
        pwd = Entry(login, show ="*")
        #when submit button is clicked, it sends credentials to Banner Portal
        submit = Button(login, text='Submit', command = lambda:[self.set_credentials(uid, pwd), login.destroy()])
        back = Button(login, text='Return', command = lambda:[login.destroy(), self.greeting_window(self.greeting)])
        welcome.grid(row = 0, columnspan = 3)
        uid.grid(row = 1, column = 1, columnspan = 2)
        pwd.grid(row = 2, column = 1, columnspan = 2)
        submit.grid(row = 3, column = 1)
        back.grid(row = 3, column = 2)
        
    #sets user's login info
    def set_credentials(self, uid, pwd):
        self.username = uid.get()
        self.password = pwd.get()
        
    #returns login info
    def get_credentials(self):
        return self.auth, self.username, self.password
    
    #method to grab advisor login info
    def click_advisor(self):
        print("Clicked 'Advisor'")
        self.auth = "advisor"
        #opens a window to grab student login info
        login = Tk()
        login.geometry("200x90")
        login.title("Auto Advisor - Advisor")
        welcome = Label(login, text = "Logging in as: Advisor")
        Label(login, text='User ID').grid(row=1)
        Label(login, text='Password').grid(row=2)
        uid = Entry(login)
        pwd = Entry(login, show ="*")
        #when submit button is clicked, it sends credentials to Banner Portal
        submit = Button(login, text='Submit', command = lambda:[self.set_credentials(uid, pwd), login.destroy()])
        back = Button(login, text='Return', command = lambda:[login.destroy(), self.greeting_window(self.greeting)])
        welcome.grid(row = 0, columnspan = 3)
        uid.grid(row = 1, column = 1, columnspan = 2)
        pwd.grid(row = 2, column = 1, columnspan = 2)
        submit.grid(row = 3, column = 1)
        back.grid(row = 3, column = 2)

    #grabs student id if advisor is logged in
    def get_student_id(self):
        get_sid = Tk()
        get_sid.geometry("230x80")
        get_sid.title("Auto Advisor")
        prompt = Label(get_sid, text = "Enter Student V-Number")
        Label(get_sid, text='V-Number:').grid(row = 1, column = 0)
        sid = Entry(get_sid)
        submit = Button(get_sid, text='Submit', command = lambda:[self.set_sid(sid), get_sid.destroy()])
        cancel = Button(get_sid, text='Cancel', command = get_sid.destroy)
        prompt.grid(row = 0, column = 1, columnspan = 2)
        sid.grid(row = 1, column = 1, columnspan = 2)
        submit.grid(row = 2, column = 1)
        cancel.grid(row = 2, column = 2)
        get_sid.mainloop()
        
    #returns student id
    def set_sid(self, vnum):
        self.sid = vnum.get()
        
    def return_sid(self):
        return self.sid

    #method that prompts user to correct invalid login info
    def invalid_creds(self):
        self.greeting = 'Invalid login credentials!'
        self.greeting_window(self.greeting)

    #instantiates class
    def __init__(self):
        self.username = ""
        self.password = ""
        self.auth = ""
        self.sid = ""
        self.greeting = 'Welcome to Auto-Advisor!'
        self.greeting_window(self.greeting)

#method to return appropriate value for term dropdown
def get_timecode():
    #get current month and year
    month = datetime.now().month
    year = str(datetime.now().year)
    #change month to string
    if month >= 1:
        #set month to "01" if month is between 1 and 4
        if month >= 5:
            #set month to "05" if month is between 5 and 7
            if month >= 8:
                #set month to "08" if month is between 8 and 11
                if month == 12:
                    #change month to string if month is 12
                    month = "12"
                else:
                    month = "08"
            else:
                month = "05"
        else:
            month = "01"
    #return year and month to use as value for dropdown selector
    return year + month
        
#method to create file to store data
def create_file_path(fullname ,path, filename, file_type, array):
    #check if file directory exists
    if os.path.exists(path):
        #if folder exists, check if file already exists
        file_path = path + filename
        #if file exists, overwrite it
        if os.path.exists(file_path):
            #remove old file, then create new file
            print("Overwriting " + file_type + " file for " + fullname + "...")
            os.remove(file_path)
            with open(file_path, 'w') as target_file:
                if file_type == "courses":
                    for line in array:
                        for index in line:
                            target_file.write(index + " ")
                        target_file.write("\n")
                else:
                    for line in array:
                        target_file.write(line + "\n")
        else:
            #create new file
            print("Creating " + file_type + " file for " + fullname + "...")
            with open(file_path, 'w') as target_file:
                if file_type == "courses":
                    for line in array:
                        for index in line:
                            target_file.write(index + " ")
                        target_file.write("\n")
                else:
                    for line in array:
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

if len(auth) == 0 or len(username) == 0 or len(password) == 0:
    sys.exit("Credentials not set")

#--------ADD A CHECK TO SEE IF WEBDRIVER IS CURRENT VERSION--------
#May be relevant: SessionNotCreatedException
driver = webdriver.Edge(r'C:\Users\terre\Documents\edgedriver_win64\msedgedriver.exe')
driver.get('https://ssb-prod.ec.vsu.edu/BNPROD/twbkwbis.P_WWWLogin')
#--------ADD A CHECK TO SEE IF WEBDRIVER IS CURRENT VERSION--------

#enter credentials into Banner
uid = driver.find_element_by_id("UserID")
pwd = driver.find_element_by_id("PIN")

uid.send_keys(username)
pwd.send_keys(password)

driver.find_element_by_xpath("//form").submit()

#--------ADD LOGIC THAT CHECKS TO SEE IF USER IS AT LANDING PAGE------
#LANDING PAGE URL: https://ssb-prod.ec.vsu.edu/BNPROD/twbkwbis.P_GenMenu?name=bmenu.P_MainMnu
#use regular expressions since end of url can vary between users
home_url = "https://ssb-prod.ec.vsu.edu/BNPROD/twbkwbis.P_GenMenu?name=bmenu.P_MainMnu"
url = driver.current_url.split("&")[0]
print(url)
#try:
    #auth, username, password = user.invalid_creds()
#Exception Handling that closes program if tkinter box is closed prematurely
#except AttributeError:
    #sys.exit("Program Terminated")
#--------ADD LOGIC THAT CHECKS TO SEE IF USER IS AT LANDING PAGE------

if auth == "advisor":
    #crawl through banner until we get to student id
    driver.find_element_by_link_text("Faculty and Advisors").click()
    driver.find_element_by_link_text("Student Information Menu").click()
    driver.find_element_by_link_text("ID Selection").click()
    #The following dropdown menu lacks an id
    #so to select for current semester, I noticed html code stored
    #dropdown values as dates, I used a method to grab current
    #month and year, and store it to a variable to use as value
    date = get_timecode()
    term = Select(driver.find_element_by_xpath("//select[@name='term']"))
    term.select_by_value(date)
    driver.find_element_by_xpath("//td[@class='dedefault']").submit()
    #requests student id and uses it to get to transcript
    user.get_student_id()
    sid = user.return_sid()
    vnum = driver.find_element_by_id("Stu_ID")
    vnum.send_keys(sid)
    driver.find_element_by_xpath("//table[contains(@class, 'dataentrytable')]").submit()
    #gets student name
    name = driver.find_element_by_tag_name('b')
    fullname = name.text.replace(" ", "").replace(".", "")
    driver.find_element_by_xpath("//input[@type='submit' and @value='Submit']").submit()

    #crawl webpage to academic transcript
    driver.find_element_by_link_text("Academic Transcript").click()

    levl_id = Select(driver.find_element_by_id("levl_id"))
    type_id = Select(driver.find_element_by_id("tprt_id"))

if auth == "student":
    #grab person's name to use for file name
    name = driver.find_elements_by_xpath("//td[@class='pldefault']")
    fullname = ""
    for i in name:
        if i.text[0:7] == "Welcome":
            fullname = i.text.split(",")[1].replace(" ", "").replace(".", "")

    #crawl webpage to academic transcript
    driver.find_element_by_link_text("Student").click()
    driver.find_element_by_link_text("Student Records").click()

    #crawl webpage to academic transcript
    driver.find_element_by_link_text("Academic Transcript").click()

    levl_id = Select(driver.find_element_by_id("levl_id"))
    type_id = Select(driver.find_element_by_id("type_id"))

#create folder using student name
#search to see if student folder already exists
path = "students/" + fullname
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
#constants used to improve readability
END_OF_SEM = 6
END_OF_GRADES = 18

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
        if counter >= END_OF_SEM and len(courses) > 0:
            #At least 18 elements are popped before we reach current/future semesters
            if counter > END_OF_GRADES:
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

