from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import NoSuchElementException, SessionNotCreatedException
import tkinter #GUI
from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
from datetime import datetime
import os #used for read/write to files
import re #used for regular expressions
import sys #used to stop execution under certain circumstances
import preprocess

#Class to hold user login info for session
class credentials():
    #Opens a window and prompts user to select config file and to log in
    def greeting_window (self, greeting, filename):
        portal = Tk()
        portal.geometry("180x60")
        portal.title("Auto Advisor: Home")
        is_anon = BooleanVar(portal, value = self.anonymous)
        is_current_sem = BooleanVar(portal, value = self.current_sem)
        welcome = Label(portal, text = greeting)
        portal.after(1, lambda: portal.focus_force())
        #if no configuration file is selected, prompt user for config file
        if len(filename) == 0 or not re.search('.xlsx$', filename):
            config = Button(portal, text="Set Configuration File", command = lambda:[portal.destroy(), self.set_config()])
            portal.bind('<Return>', lambda x:[portal.destroy(), self.set_config()])
            welcome.pack(side = TOP, expand = True)
            config.pack(side = BOTTOM, expand = True)
        #if a config file has be selected, prompt user to log in
        else:
            portal.geometry("260x120")
            login = Button(portal, text="Log In", command = lambda:[portal.destroy(), self.click_login()])
            config = Button(portal, text="Change Configuration File", command = lambda:[portal.destroy(), self.set_config()])
            portal.bind('<Return>', lambda x:[portal.destroy(), self.click_login()])
            anon_btn = Checkbutton(portal, text="Hide Student Names", variable = is_anon, command = lambda:[self.toggle_anonymous(is_anon.get())])
            sem_btn = Checkbutton(portal, text="Toggle on if planning for Current Semester", variable = is_current_sem, command = lambda:[self.toggle_sem(is_current_sem.get())])
            config.pack(side = BOTTOM, expand = True)
            welcome.pack(side = TOP, expand = True)
            login.pack(side = BOTTOM, expand = True)
            anon_btn.pack(side = TOP, expand = True)
            sem_btn.pack(side = TOP, expand = True)
        portal.mainloop()

    def toggle_anonymous (self, is_anon):
        if is_anon:
            self.anonymous = True
        else:
            self.anonymous = False

    def is_anonymous(self):
        return self.anonymous

    def toggle_sem (self, is_current_sem):
        if is_current_sem:
            self.current_sem = True
        else:
            self.current_sem = False

    def get_sem_flag(self):
        return self.current_sem

    #function that grabs file path of config file
    def set_config(self):
        file_name = askopenfilename(title = 'Select Config File', filetypes = [('Excel Files','*.xlsx')])
        self.filename = file_name
        self.greeting_window(self.greeting, self.filename)               

    #function that returns config file path
    def get_config(self):
        return self.filename

    #method to grab student login info
    def click_login(self):
        #opens a window to grab student login info
        login = Tk()
        login.geometry("200x103")
        login.title("Auto Advisor: Login")
        welcome = Label(login, text = "Enter your V-Number and PIN:")
        login.after(1, lambda: login.focus_force())
        Label(login, text='V-Number').grid(row=1, pady = 1)
        Label(login, text='PIN').grid(row=2, pady = 1)
        uid = Entry(login)
        pwd = Entry(login, show ="*")
        #when submit button is clicked, it sends credentials to Banner Portal
        submit = Button(login, text='Submit', command = lambda:[self.set_credentials(uid, pwd), login.destroy()])
        back = Button(login, text='Return', command = lambda:[login.destroy(), self.greeting_window(self.greeting, self.filename)])
        login.bind('<Return>', lambda x:[self.set_credentials(uid, pwd), login.destroy()])
        welcome.grid(row = 0, columnspan = 3, padx = 10,pady = 2)
        uid.grid(row = 1, column = 1, columnspan = 3)
        pwd.grid(row = 2, column = 1, columnspan = 3)
        submit.grid(row = 3, column = 0, columnspan = 2, ipadx = 20, pady = 3)
        back.grid(row = 3, column = 2, columnspan = 2, ipadx = 20, pady = 3)
        
    #sets user's login info
    def set_credentials(self, uid, pwd):
        self.username = uid.get()
        self.password = pwd.get()
        
    #returns login info
    def get_credentials(self):
        return self.username, self.password

    #warns the user that they have failed to log in 3 times
    def login_warning(self):
        warn = Tk()
        warn.geometry("200x90")
        warn.title("Warning!")
        warning = Label(text = "Warning! You have failed to login 3 times now. Please ensure that you enter your information correctly.")
        warn.after(1, lambda: warn.focus_force())
        back = Button(warn, text='Return', command = lambda:[warn.destroy(), self.greeting_window(self.greeting, self.filename)])
        warn.bind('<Return>', lambda x:[warn.destroy(), self.greeting_window(self.greeting, self.filename)])
        warning.pack(side = TOP)
        back.pack(side = TOP)
        warn.mainloop()

    #terminates the program if user fails to log in 4 times
    def login_timeout(self):
        terminate = Tk()
        terminate.geometry("200x90")
        terminate.title("Too Many Login Attempts")
        terminate.after(1, lambda: terminate.focus_force())
        message = Label(text = "You have attempted too many failed login attempts. To prevent account lockout, please try again later.")
        end = Button(terminate, text='Exit', command = lambda:[terminate.destroy(), driver.quit(), sys.exit("Program Terminated: Too Many Login Attempts!")])
        terminate.bind('<Return>', lambda x:[terminate.destroy(), driver.quit(), sys.exit("Program Terminated: Too Many Login Attempts!")])
        message.pack(side = TOP)
        end.pack(side = TOP)
        terminate.mainloop()

    #method that prompts user to correct invalid login info
    def invalid_creds(self):
        self.greeting = 'Invalid login credentials!'
        self.login_count += 1
        if self.login_count < 3:
            self.greeting_window(self.greeting, self.filename)
        elif self.login_count == 3:
            self.login_warning()
        else:
            self.login_timeout()

    #method that warns the user when student login credentials are entered
    def student_detected(self):
        window = Tk()
        window.geometry("200x50")
        window.title("Student Detected!")
        window.after(1, lambda: window.focus_force())
        msg = Label(window, text="Only Advisors can use AutoAdvisor!")
        ok_btn = Button(window, text="OK", command = window.destroy)
        window.bind('<Return>', window.destroy)
        msg.pack(side = TOP)
        ok_btn.pack(side = BOTTOM)
        window.mainloop()

    #function that grabs list of v-numbers from a text file
    def get_vnums(self):
        vnums = []
        while True:
            #get path of text file
            file_name = askopenfilename(title = 'Select vnums File', filetypes = [('Text Files', '*.txt')])
            #if no file is selected, prompt the user to select a file
            if len(file_name) == 0:
                    window = Tk()
                    window.geometry("200x70")
                    window.title("Invalid File!")
                    window.after(1, lambda: window.focus_force())
                    msg = Label(window, text="Warning: Need vnums text file!")
                    ok_btn = Button(window, text="OK", command = window.destroy)
                    cancel_btn = Button(window, text='Cancel', command = lambda:[window.destroy(), driver.quit(), sys.exit('Program Terminated')])
                    window.bind('<Return>', window.destroy)
                    msg.pack(side = TOP)
                    ok_btn.pack(side = BOTTOM)
                    cancel_btn.pack(side = BOTTOM)
                    window.mainloop()
            #if a text file is selected, grab v-numbers from it
            else:
                self.filename = file_name
                f1 = open(file_name)
                for vnum in f1:
                    vnums.append(vnum)
                f1.close()
                #return list of v-numbers
                print(vnums)
                return vnums

    #UNTESTED CODE!
    #Warns the user if Edge webdriver is not installed or up to date
    def update_webdriver(self):
        update = Tk()
        update.geometry('200x90')
        update.title('Update Edge Webdriver')
        update.after(1, lambda: update.focus_force())
        msg = Label(update, text="Please update Edge webdriver!")
        ok_btn = Button(update, test="OK", command = lambda:[update.destroy, sys.exit('Program Terminated')])
        update.bind('<Return>', update.destroy)
        msg.pack(side = TOP)
        ok_btn.pack(side = BOTTOM)
        update.mainloop()

    #instantiates class
    def __init__(self):
        self.username = ""
        self.password = ""
        self.filename = ""
        self.greeting = 'Welcome to Auto-Advisor!'
        self.login_count = 0
        self.anonymous = False
        self.current_sem = False
        self.greeting_window(self.greeting, self.filename)

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

def build_path(path, fullname):
    #create folder using student name
    #check if name has been established
    if bool(fullname):
        #search to see if student folder already exists
        if not os.path.exists(path):
            os.makedirs(path)
            if os.path.exists(path):
                print("Path successfully created!")
            else:
                print("Failed to create path!")
        else:
            print("Path already exists!")
    else:
        sys.exit("An unexpected error has occurred: Unable to locate student's name!")

def build_files(path, driver, fullname):
    driver.find_element(By.XPATH, "//table[contains(@class, 'dataentrytable')]").submit()

    #scrape transcript for courses and semesters
    data = driver.find_elements(By.XPATH, "//table[@class = 'datadisplaytable']")

    courses = []
    semesters = []
    #boolean flags to distinguish semesters
    sem_start = sem_end = False
    #used to determine if we reached current/future semesters
    is_curr = False
    #used to hold course as we build it from data
    proto = ""

    for i in data:
        output = i.text.split('\n')

    for line in output:
        #Find each semester as we iterate through scraped data
        #signifies the start of courses in progress
        if re.search("COURSES IN PROGRESS", line) and not is_curr:
            is_curr = True
        if re.search("Spring", line) or re.search("Summer", line) or \
           re.search("Fall", line) or re.search("Winter", line) or \
           re.search("1:", line):
            if sem_start:
                #lets us know that we've reached the end of a previous semester
                sem_end = True
            #signifies the start of a new semester
            sem_start = True
            #store the current semester into semesters array
            if re.search("Spring [0-9]{4}|Summer [0-9]{4}|Fall [0-9]{4}|Winter [0-9]{4}", line):
                semesters.append(re.search("Spring [0-9]{4}|Summer [0-9]{4}|Fall [0-9]{4}|Winter [0-9]{4}", line).group())
            #for semesters that are missing a term name
            else:
                semesters.append("UNKNOWN")
            semesters.append("-")
        #if we are in a semester, find the courses
        elif sem_start:
            #add marker to course array to signify semesters
            if sem_end:
                courses.append("-")
                sem_end = False
            #append credits to course
            if bool(proto):
                proto.append(line)
                course = proto.copy()
                #pop empty indices at beginning of course
                while not bool(course[0]):
                    course.pop(0)
                #remove 'U' from courses
                if course[2] == 'U':
                    course.pop(2)
                #append course to array and reset proto
                courses.append(course)
                proto.clear()
            #find a course and store it to the holder array
            elif re.search("[A-Z]{4} ", line) and not re.search("-Top-", line):
                proto = line.replace("\n", " ").split(" ")
                #check if we reached current/future semesters
                if is_curr:
                    #add 'inprog' to current/future courses
                    proto.append("inprog")

    #Append separator for final course structure
    courses.append("-")

    #print courses to file
    create_file_path(fullname , path, "/courses.txt", "courses", courses)
    #print semesters to file
    create_file_path(fullname , path, "/semesters.txt", "semesters", semesters)
        
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
    username, password = user.get_credentials()
#Exception Handling that closes program if tkinter box is closed prematurely
    if not bool(username) or not bool(password):
        sys.exit("Program Terminated!")
except AttributeError:
    sys.exit('Program Terminated!')

config_file = user.get_config()

#--------ADD A CHECK TO SEE IF WEBDRIVER IS CURRENT VERSION--------
#May be relevant: SessionNotCreatedException
options = Options()
options.add_argument("headless")
try:
    #driver = webdriver.Edge(options = options)
    driver = webdriver.Edge()
    driver.get('https://ssb-prod.ec.vsu.edu/BNPROD/twbkwbis.P_WWWLogin')
except SessionNotCreatedException:
    user.update_webdriver()
#--------ADD A CHECK TO SEE IF WEBDRIVER IS CURRENT VERSION--------

#Track if user is at landing page
home_url = "https://ssb-prod.ec.vsu.edu/BNPROD/twbkwbis.P_GenMenu?name=bmenu.P_MainMnu"
url = driver.current_url.split("&")[0]

vnums = []
is_logged = False

#Create loop to allow user to login with multiple attempts
i = 0
while not is_logged:
    #set new login credentials
    username, password = user.get_credentials()
    
    if not bool(username) or not bool(password):
        continue
        
    #attempt to log in
    uid = driver.find_element(By.ID, "UserID")
    pwd = driver.find_element(By.ID, "PIN")

    uid.send_keys(username)
    pwd.send_keys(password)

    driver.find_element(By.XPATH, "//form").submit()

    i += 1

    #detect whether student or advisor is logged in by available links
    try:
        #if this link is available, then advisor is logged in
        driver.find_element(By.LINK_TEXT, "Faculty and Advisors")
        is_logged = True
    #if above link can't be found, check if student is logged in
    except NoSuchElementException:
        try:
            driver.find_element(By.LINK_TEXT, "Student")
            user.student_detected()
            driver.quit()
            sys.exit("Program Terminated!")
        #if no successful login detected, prompt user to try again
        except NoSuchElementException:
            old_username = username
            old_password = password
            user.invalid_creds()

            if old_username == username and old_password == password:
                driver.quit()
                sys.exit("No input change for login credentials detected! Terminating Porgram...")

vnums = user.get_vnums()

#Used to hold student names
fullname = []
names = []
error_vnums = []

is_anonymous = user.is_anonymous()
sem_flag = user.get_sem_flag()

itr = 0
#crawl through banner until we get to student id
driver.find_element(By.LINK_TEXT, "Faculty and Advisors").click()
index = 0
first = True
while index < len(vnums):
    driver.find_element(By.LINK_TEXT, "Student Information Menu").click()
    driver.find_element(By.LINK_TEXT, "ID Selection").click()
    #The following dropdown menu lacks an id
    #so to select for current semester, I noticed html code stored
    #dropdown values as dates, I used a method to grab current
    #month and year, and store it to a variable to use as value
    date = get_timecode()
    #We visit the 'Select Term' page only once, during the first time we crawl
    if first:
        term = Select(driver.find_element(By.XPATH, "//select[@name='term']"))
        term.select_by_value(date)
        driver.find_element(By.XPATH, "//td[@class='dedefault']").submit()
        first = False
    #requests student id and uses it to get to transcript
    sid = driver.find_element(By.ID, "Stu_ID")
    sid.send_keys(vnums[index])
    #Additional crawling logic if we've reached the last vnum in list
    if index == len(vnums):
        driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
        #gets student name
        try:
            name = driver.find_element(By.TAG_NAME, 'b')
            names.append(name.text.split(" "))
            fullname.append(name.text.replace(" ", "").replace(".", ""))
            driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
        #if this error occurs, then v-number is not associated with a student
        except NoSuchElementException:
            print("Error occurred at " + vnums[index])
            error_vnums.append(vnums.pop(index))
            driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
            driver.find_element(By.LINK_TEXT, "Faculty Services").click()
            index += 1
            continue
    #For all other vnums that are not last in list
    else:
        try:
            name = driver.find_element(By.TAG_NAME, 'b')
            names.append(name.text.split(" "))
            fullname.append(name.text.replace(" ", "").replace(".", ""))
            driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
        except NoSuchElementException:
            print("Error occurred at " + vnums[index])
            error_vnums.append(vnums.pop(index))
            index -= 1
            driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
            driver.find_element(By.LINK_TEXT, "Faculty Services").click()
            index += 1
            continue
        

    #crawl webpage to academic transcript
    driver.find_element(By.LINK_TEXT, "Academic Transcript").click()

    levl_id = Select(driver.find_element(By.ID, "levl_id"))
    type_id = Select(driver.find_element(By.ID, "tprt_id"))

    #set path for student using their name
    if is_anonymous:
        path = "students/" + vnums[index].strip() + '/' + config_file.split('/')[-1].split('.')[0]
        build_path(path, vnums[index].strip())
        build_files(path, driver, vnums[index].strip())
    else:
        path = "students/" + fullname[index].strip() + '/' + config_file.split('/')[-1].split('.')[0]
        build_path(path, fullname[index].strip())
        build_files(path, driver, fullname[index].strip())

    driver.find_element(By.LINK_TEXT, "Faculty Services").click()
    itr += 1
    index += 1

#Web driver is no longer needed
driver.quit()

if is_anonymous:
    preprocess.main(vnums, config_file, vnums, names, sem_flag)
else:
    preprocess.main(fullname, config_file, vnums, names, sem_flag)
print('Program complete! Check files for advisory report(s).')
if len(error_vnums) > 0:
    print('Could not generate transcripts for the following V-Numbers:')
    for vnum in error_vnums:
        print(vnum)
