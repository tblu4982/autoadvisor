from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import NoSuchElementException, SessionNotCreatedException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import tkinter #GUI
from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
from datetime import datetime
import os #used for read/write to files
import re #used for regular expressions
import sys #used to stop execution under certain circumstances
import preprocess

VERSION = "2.4.5"

#Class to hold user login info for session
class credentials():
    #Opens a window and prompts user to select config file and to log in
    def greeting_window (self, greeting, filename):
        portal = Tk()
        #portal.geometry("200x90")
        portal.title("Auto Advisor: Home")
        is_anon = BooleanVar(portal, value = self.anonymous)
        is_current_sem = BooleanVar(portal, value = self.current_sem)
        welcome = Label(portal, text = greeting)
        portal.after(1, lambda: portal.focus_force())
        #if no configuration file is selected, prompt user for config file
        if len(filename) == 0 or not re.search('.xlsx$', filename):
            config = Button(portal, text="Set Configuration File", command = lambda:[portal.destroy(), self.set_config()])
            portal.bind('<Return>', lambda x:[portal.destroy(), self.set_config()])
            welcome.pack(side = TOP)
            config.pack(side = BOTTOM)
        #if a config file has be selected, prompt user to log in
        else:
            login = Button(portal, text="Log In", command = lambda:[portal.destroy(), self.click_login()])
            config = Button(portal, text="Change Configuration File", command = lambda:[portal.destroy(), self.set_config()])
            sem_btn = Checkbutton(portal, text="Toggle on if planning for Current Semester", variable = is_current_sem, command = lambda:[self.toggle_sem(is_current_sem.get())])
            portal.bind('<Return>', lambda x:[portal.destroy(), self.click_login()])
            config.pack(side = BOTTOM)
            welcome.pack(side = TOP)
            login.pack(side = BOTTOM)
            sem_btn.pack(side = TOP, expand = True)
        anon_btn = Checkbutton(portal, text="Anonymous Mode", variable = is_anon, command = lambda:[self.toggle_anonymous(is_anon.get())])
        anon_btn.pack(side = TOP)
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
        def on_option_value_change(*args):
            self.set_verify_method(option.get())
        #opens a window to grab student login info
        login = Tk()
        #login.geometry("200x90")
        login.title("Auto Advisor: Login")
        auth_options = ["Google Authenticator", "Okta 2FA Code", "Okta Push Notification"]
        welcome = Label(login, text = "Enter your V-Number and PIN:")
        login.after(1, lambda: login.focus_force())
        Label(login, text='E-mail').grid(row=1)
        Label(login, text='Password').grid(row=2)
        Label(login, text="@vsu.edu").grid(row=1, column = 2)
        Label(login, text="Auth Method:").grid(row = 3)
        uid = Entry(login, width = 15)
        pwd = Entry(login, show ="*", width = 25)
        option = StringVar(value=self.get_verify_method())
        option.trace_add("write", on_option_value_change)
        print(self.get_verify_method())
        auth_menu = OptionMenu(login, option, self.get_verify_method(), *auth_options)
        #when submit button is clicked, it sends credentials to Banner Portal
        submit = Button(login, text='Submit', command = lambda:[self.set_credentials(uid, pwd), self.set_verify_method(option.get()), login.destroy()])
        back = Button(login, text='Return', command = lambda:[login.destroy(), self.greeting_window(self.greeting, self.filename)])
        login.bind('<Return>', lambda x:[self.set_credentials(uid, pwd), login.destroy()])
        welcome.grid(row = 0, columnspan = 3)
        uid.grid(row = 1, column = 1)
        pwd.grid(row = 2, column = 1, columnspan = 2)
        auth_menu.grid(row = 3, column = 1, columnspan = 2)
        submit.grid(row = 4, column = 1)
        back.grid(row = 4, column = 2)
        
    #sets user's login info
    def set_credentials(self, uid, pwd):
        self.username = uid.get() + "@vsu.edu"
        self.password = pwd.get()
        
    #returns login info
    def get_credentials(self):
        return self.username, self.password

    #---LEGACY CODE! Considering removal!!!
    #warns the user that they have failed to log in 3 times
    def login_warning(self):
        warn = Tk()
        #warn.geometry("200x90")
        warn.title("Warning!")
        warning = Label(text = "Warning! You have failed to login 3 times now. Please ensure that you enter your information correctly.")
        warn.after(1, lambda: warn.focus_force())
        back = Button(warn, text='Return', command = lambda:[warn.destroy(), self.greeting_window(self.greeting, self.filename)])
        warn.bind('<Return>', lambda x:[warn.destroy(), self.greeting_window(self.greeting, self.filename)])
        warning.pack(side = TOP)
        back.pack(side = TOP)
        warn.mainloop()

    #---LEGACY CODE! Considering removal!!!
    #terminates the program if user fails to log in 4 times
    def login_timeout(self):
        terminate = Tk()
        #terminate.geometry("200x90")
        terminate.title("Too Many Login Attempts")
        terminate.after(1, lambda: terminate.focus_force())
        message = Label(text = "You have attempted too many failed login attempts. To prevent account lockout, please try again later.")
        end = Button(terminate, text='Exit', command = lambda:[terminate.destroy(), driver.quit(), sys.exit("Program Terminated: Too Many Login Attempts!")])
        terminate.bind('<Return>', lambda x:[terminate.destroy(), driver.quit(), sys.exit("Program Terminated: Too Many Login Attempts!")])
        message.pack(side = TOP)
        end.pack(side = TOP)
        terminate.mainloop()

    #---LEGACY CODE! Consider modifying!!!
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

    #---LEGACY CODE, no longer functional. Consider Removing!!!
    #method that warns the user when student login credentials are entered
    def student_detected(self):
        window = Tk()
        #window.geometry("200x50")
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
                for line in f1:
                    vnum = line.strip()
                    if bool(vnum):
                        vnums.append(vnum)
                f1.close()
                #return list of v-numbers
                return vnums

    #UNTESTED CODE!
    #Warns the user if Edge webdriver is not installed or up to date
    def update_webdriver(self):
        update = Tk()
        #update.geometry('200x90')
        update.title('Update Edge Webdriver')
        update.after(1, lambda: update.focus_force())
        msg = Label(update, text="Please update Edge webdriver!")
        ok_btn = Button(update, text="OK", command = lambda:[update.destroy(), sys.exit('Program Terminated')])
        update.bind('<Return>', update.destroy)
        msg.pack(side = TOP)
        ok_btn.pack(side = BOTTOM)
        update.mainloop()
    
    #Gets 2FA input from user    
    def get_pin(self):
        pinget = Tk()
        auth_method = self.get_verify_method()
        match auth_method:
            case "Google Authenticator":
                pinget.title('Google Auth 2FA Code')
                pinget.after(1, lambda: pinget.focus_force())
                msg = Label(pinget, text='Enter 2FA Code:')
                code = Entry(pinget)
                submit = Button(pinget, text='Submit', command = lambda:[self.set_pin(code), pinget.destroy()])
                pinget.bind('<Return>', lambda x:[self.set_pin(code), pinget.destroy()])
                msg.pack(side = TOP)
                code.pack()
                submit.pack(side = BOTTOM)
                pinget.mainloop()
            case "Okta 2FA Code":
                pinget.title('Okta 2FA Code')
                pinget.after(1, lambda: pinget.focus_force())
                msg = Label(pinget, text='Enter 2FA Code:')
                code = Entry(pinget)
                submit = Button(pinget, text='Submit', command = lambda:[self.set_pin(code), pinget.destroy()])
                pinget.bind('<Return>', lambda x:[self.set_pin(code), pinget.destroy()])
                msg.pack(side = TOP)
                code.pack()
                submit.pack(side = BOTTOM)
                pinget.mainloop()
            case "Okta Push Notification":
                pinget.destroy()
       
    #Sets 2FA code for retrieval    
    def set_pin(self, pin):
        self.pin = pin.get()
    
    #Fetches stored 2FA code
    def return_pin(self):
        return self.pin
    
    def set_verify_method(self, method):
        self.verify_method = method
        
    def get_verify_method(self):
        return self.verify_method

    #instantiates class
    def __init__(self):
        self.username = ""
        self.password = ""
        self.filename = ""
        self.pin = ""
        self.greeting = 'Welcome to Auto-Advisor!'
        self.login_count = 0
        self.anonymous = False
        self.current_sem = False
        self.verify_method = "Okta Push Notification"
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
    #scrape transcript for courses and semesters
    try:
        wait.until(EC.visibility_of_all_elements_located((By.XPATH, "//table")))
        data = driver.find_elements(By.XPATH, "//table")

        courses = []
        semesters = []
        #boolean flags to distinguish semesters
        sem_start = sem_end = False
        #used to determine if we reached current/future semesters
        is_curr = False
        #used to hold course as we build it from data
        proto = ""
        course_marker = False
        
        output = []

        for i in data:
            output += i.text.split('\n')

        for line in output:
            #Find each semester as we iterate through scraped data
            #signifies the start of courses in progress
            if re.search("Course\(s\) in progress", line) and not is_curr:
                is_curr = True
            if re.search("Subject Course Level Title Grade Credit Hours Quality Points Start and End Dates R|Subject Course Title Grade Credit hours Quality points R|Subject Course Level Title Credit Hours Start and End Dates|Subject Course Campus Level Title Credit Hours Start and End Dates", line):
                if sem_start:
                    #lets us know that we've reached the end of a previous semester
                    sem_end = True
                #signifies the start of a new semester
                sem_start = True
            #if we are in a semester, find the courses
            elif sem_start:
                #add marker to course array to signify semesters
                if sem_end:
                    courses.append("-")
                    sem_end = False
                #append credits to course
                if course_marker:
                    if re.search("[0-9]+.[0-9]{3}", line):
                        course_marker = False
                        for i in line.replace("\n", " ").split(" "):
                            proto.append(i)
                        course = proto.copy()
                        #pop empty indices at beginning of course
                        while not bool(course[0]):
                            course.pop(0)
                        #remove 'U' from courses
                        if course[2] == 'U':
                            course.pop(2)
                        #remove unnecessary indices at end of course
                        try:
                            while re.search("[0-9]+.[0-9]{3}" ,course[-2]):
                                course.pop()
                        except IndexError as e:
                            print(e)
                            print(course)
                            print(line)
                            sys.exit()
                        #append course to array and reset proto
                        courses.append(course)
                        #print(proto)
                        proto.clear()
                    #check if we reached current/future semesters
                    else:
                        for i in line.replace("\n", " ").split(" "):
                            proto.append(i)
                #find a course and store it to the holder array
                elif re.search("[A-Z]{4}", line):
                    proto = line.replace("\n", " ").split(" ")
                    #Extra code to catch transfer courses and current semester courses, since their contents are stored in one line
                    if re.search("0.000$", line) or is_curr:
                        course = proto.copy()
                        #Add 'inprog' to current semester courses as they have no letter grade
                        if is_curr:
                            course.insert(-1, "inprog")
                        #remove unnecessary indices at end of course
                        else:
                            course.pop()
                        courses.append(course)
                        proto.clear()
                    else:
                        course_marker = True

        #Append separator for final course structure
        courses.append("-")
        
        data = driver.find_elements(By.CSS_SELECTOR, ".sub-heading.period-padding.ng-binding")
        
        output = []

        for i in data:
            output += i.text.split('\n')
            
        for line in output:
            semesters.append(line)
            semesters.append('-')

        #print courses to file
        create_file_path(fullname , path, "/courses.txt", "courses", courses)
        #print semesters to file
        create_file_path(fullname , path, "/semesters.txt", "semesters", semesters)
        
        return True
    except TimeoutException:
        return False
        
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
try:
    #driver = webdriver.Edge(options = options)
    driver = webdriver.Edge()
    driver.get('https://login.vsu.edu')
except SessionNotCreatedException:
    user.update_webdriver()
#--------ADD A CHECK TO SEE IF WEBDRIVER IS CURRENT VERSION--------

wait = WebDriverWait(driver, 10)

vnums = []
is_logged = False

#Create loop to allow user to login with multiple attempts
username, password = user.get_credentials()
    
#attempt to log in
wait.until(EC.visibility_of_element_located((By.ID, "input28")))
uid = driver.find_element(By.ID, "input28")
uid.send_keys(username)
driver.find_element(By.CSS_SELECTOR, ".button").click()

wait = WebDriverWait(driver, 5)

try:
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Verify with something else")))
    driver.find_element(By.LINK_TEXT, "Verify with something else").click()
except TimeoutException:
    pass

wait = WebDriverWait(driver, 600)

# Get authentication method from user
auth_type = user.get_verify_method()

match auth_type:
    # Get 2FA code from Google Authenticator
    case "Google Authenticator":
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".authenticator-row:nth-child(1) .button")))
        driver.find_element(By.CSS_SELECTOR, ".authenticator-row:nth-child(1) .button").click()
        
        user.get_pin()
        pin = user.return_pin()
        wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@type = 'text']")))
        pin_entry = driver.find_element(By.XPATH, "//input[@type = 'text']")
        pin_entry.send_keys(pin)
        driver.find_element(By.CSS_SELECTOR, ".button").click()
        wait = WebDriverWait(driver, 10)
    # 2FA Code from Okta Verify
    case "Okta 2FA Code":
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".authenticator-row:nth-child(2) .button")))
        driver.find_element(By.CSS_SELECTOR, ".authenticator-row:nth-child(2) .button").click()
        
        user.get_pin()
        pin = user.return_pin()
        wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@type = 'text']")))
        pin_entry = driver.find_element(By.XPATH, "//input[@type = 'text']")
        pin_entry.send_keys(pin)
        driver.find_element(By.CSS_SELECTOR, ".button").click()
        wait = WebDriverWait(driver, 10)
    # Push Notification method from Okta Verify
    case "Okta Push Notification":
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".authenticator-row:nth-child(3) .button")))
        driver.find_element(By.CSS_SELECTOR, ".authenticator-row:nth-child(3) .button").click()
        user.get_pin()
    case _:
        raise Exception("Error! authentication method does not match known values! (" + auth_type + ")")

try:
    wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@type = 'password']")))    
    pwd = driver.find_element(By.XPATH, "//input[@type = 'password']")
    pwd.send_keys(password)
    driver.find_element(By.CSS_SELECTOR, ".button").click()
except TimeoutException:
    pass

wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@aria-label='launch app Banner Faculty Self Service 9']")))
driver.find_element(By.XPATH, "//a[@aria-label='launch app Banner Faculty Self Service 9']").click()

wait = WebDriverWait(driver, 10)

original_window = driver.current_window_handle

for window_handle in driver.window_handles:
    if window_handle != original_window:
        driver.switch_to.window(window_handle)
        
second_window = driver.current_window_handle

vnums = user.get_vnums()

wait.until(EC.title_is("Faculty Services Dashboard"))
driver.find_element(By.LINK_TEXT, "Advising Student Profile").click()

#Used to hold student names
fullname = []
names = []
error_vnums = []

is_anonymous = user.is_anonymous()
sem_flag = user.get_sem_flag()

dt = datetime.now()
timestamp = dt.strftime("%b") + "-" + str(dt.day) + "-" + str(dt.year) + "-" + str(dt.hour) + "-" + str(dt.minute) + "-" + str(dt.second)

status = Tk()
status_txt = StringVar()
status_txt.set(f"(0/{len(vnums)})")
status_msg = Label(status, textvariable = status_txt).pack()

progress = Progressbar(status, orient = HORIZONTAL, length = 100, maximum = len(vnums), mode = 'determinate')

def status_update(i):
    progress['value'] = i
    status_txt.set(f"{int(progress['value'])}/{len(vnums)} Students Processed")
    #status_msg = Label(status, text = f"{int(progress['value'])}/{len(vnums)} Students Processed").pack()
    progress.update_idletasks()
    progress.update()

progress.pack(pady = 10)
#status.mainloop()

itr = 0
#crawl through banner until we get to student id
index = 0
first = True
while index < len(vnums):
    status_update(index)
    advisor = ""
    #requests student id and uses it to get to transcript
    wait.until(EC.visibility_of_element_located((By.ID, "s2id_select2-term")))
    sid = driver.find_element(By.ID, "idSearchInput")
    action = ActionChains(driver).send_keys_to_element(sid, vnums[index]).perform()
    wait = WebDriverWait(driver, 1)
    #Additional crawling logic if we've reached the last vnum in list
    #if index == len(vnums):
        #driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
        #gets student name
    #try:
    try:
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '.search-result.name')))
    except TimeoutException:
        action = ActionChains(driver).send_keys_to_element(sid, vnums[index]).send_keys_to_element(sid, Keys.ENTER).perform()
        try:
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '.search-result.name')))
        except TimeoutException:
            print(f"Error! {vnums[index]} could not be found!")
            error_vnums.append([vnums.pop(index), "Invalid V-Number"])
            continue
    wait = WebDriverWait(driver, 10)
    name = driver.find_element(By.CSS_SELECTOR, '.search-result.name')
    if name != "No match found.":
        names.append(name.text.split(" "))
        if names[0] == "Mr":
            names.pop(0)
        fullname.append(name.text.replace(" ", "").replace(".", ""))
        wait.until(EC.element_to_be_clickable((By.ID, "term-go"))).click()
    else:
        print(f"Error! {vnums[index]} could not be found!")
        error_vnums.append(vnums.pop(index))
        continue

    #crawl webpage to academic transcript
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Academic Transcript")))
    try:
        advisor_listing = driver.find_element(By.CSS_SELECTOR, ".facultyLinkClass:nth-child(2)").text.split(" ")
        advisor = advisor_listing[-1]
    except NoSuchElementException:
        pass
    driver.find_element(By.LINK_TEXT, "Academic Transcript").click()
    
    for window_handle in driver.window_handles:
        if window_handle != original_window and window_handle != second_window:
            driver.switch_to.window(window_handle)
            
    wait.until(EC.visibility_of_element_located((By.ID, "transcriptLevelSelection")))
    driver.find_element(By.ID, "transcriptLevelSelection").click()
    wait.until(EC.visibility_of_element_located((By.XPATH, "//li[@id='ui-select-choices-row-1-']/div/div")))
    driver.find_element(By.XPATH, "//li[@id='ui-select-choices-row-1-']/div/div").click()
    wait.until(EC.visibility_of_element_located((By.ID, "transcriptTypeSelection")))
    driver.find_element(By.ID, "transcriptTypeSelection").click()
    wait.until(EC.visibility_of_element_located((By.XPATH, "//li[@id='ui-select-choices-row-2-']/div/div")))
    driver.find_element(By.XPATH, "//li[@id='ui-select-choices-row-2-']/div/div").click()
    wait.until(EC.visibility_of_element_located((By.XPATH, "//button[contains(.,'Submit')]")))
    driver.find_element(By.XPATH, "//button[contains(.,'Submit')]").click()

    if not bool(advisor):
        advisor = "No Advisor Listed"
    """if not os.path.exists(os.path.join("advisors", advisor)):
        os.makedirs(os.path.join("advisors", advisor))
        print("Created new advisor directory for " + advisor)"""
        
    print("Adding " + fullname[index] + " to advisor " + advisor)
    print("(student " + str(index + 1) + " of " + str(len(vnums)) + ")")

    #set path for student using their name
    if is_anonymous:
        path = "advisors/" + timestamp + "/" + advisor + "/" + vnums[index].strip() + '/' + config_file.split('/')[-1].split('.')[0]
        build_path(path, vnums[index].strip())
        if (not build_files(path, driver, vnums[index].strip())):
            error_vnums.append([vnums.pop(index), "Transcript missing or unable to parse"])
            driver.close()
            driver.switch_to.window(second_window)
            driver.find_element(By.LINK_TEXT, "Advisee Search").click()
            fullname.pop(index)
            continue
    else:
        path = "advisors/" + timestamp + "/" + advisor + "/" + fullname[index].strip() + '/' + config_file.split('/')[-1].split('.')[0]
        build_path(path, fullname[index].strip())
        if (not build_files(path, driver, fullname[index].strip())):
            error_vnums.append([vnums.pop(index), "Transcript missing or unable to parse"])
            driver.close()
            driver.switch_to.window(second_window)
            driver.find_element(By.LINK_TEXT, "Advisee Search").click()
            fullname.pop(index)
            continue

    driver.close()
    driver.switch_to.window(second_window)
    driver.find_element(By.LINK_TEXT, "Advisee Search").click()
    itr += 1
    index += 1

#Web driver is no longer needed
driver.quit()
status.quit()

#--------------CHANGE FUNCTION CALLS TO OTHER SCRIPTS SO IT READS ADVISOR FROM DIRECTORY------------

if is_anonymous:
    preprocess.main(vnums, config_file, vnums, names, sem_flag, timestamp)
else:
    preprocess.main(fullname, config_file, vnums, names, sem_flag, timestamp)
print('Program complete! Check files for advisory report(s).')
if len(error_vnums) > 0:
    print('AutoAdvisor has encountered an issue parsing some V-Numbers. Please check the error file')
    f2 = "advisors/" + timestamp + "/error_vnums.txt"
    with open(f2, 'w') as outfile:
        outfile.write("AutoAdvisor has encountered an issue parsing the following V-Number(s):\n")
        for vnum in error_vnums:
            outfile.write(f"{vnum}\n")
