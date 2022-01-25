# This code opens up the BridgeEDU site, logs in as Brenton Kreiger,
# and scrapes for the latest gradebook from the desired course, then
# downloads the gradebook csv to the downloads folder.

# import select packages for selenium, more info at https://selenium-python.readthedocs.io/getting-started.html
from typing import Type
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.touch_actions import TouchActions
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
import pandas as pd

import GradeScrape

# Grab the course number and name from the main script "GradeScrape.py" as strings
c = GradeScrape.course
n = GradeScrape.name
lab = GradeScrape.program
sem = GradeScrape.semester
course_label = 'BP' + c

# Get some time data
year = datetime.utcnow().strftime('%Y')
date = datetime.utcnow().strftime('%Y-%m-%d-')

# Css path for the course
css_path = 'a[href$="' + lab + '+' + c + '+' + sem + '/course/"]'

# Set up paths (make sure available offline on drive stream)
path_grading_master = "/Volumes/GoogleDrive/Shared drives/Ambassador Program/Committee_BridgeEDU/Bridge Program Coordinator/BEDU Grading 2021-2022.xlsx"

# # Get enrollment for the course ## Commented out August 30, 2021 because this is for certs code
# sheet0 = pd.read_excel(path_grading_master, sheet_name=course_label) #from google stream
# enroll = sheet0[['Username','Grade','Certificate Delivered']]
#
# # Drop all the honors kids
# # enroll.drop(enroll.index[enroll['Enrollment Track'] == 'honor'], inplace = True)
#
# # Drop all the certified kids
# enroll.drop(enroll.index[enroll['Certificate Delivered'] == 'Y'], inplace = True)
#
# # Drop all those who aren't passing
# enroll.drop(enroll.index[enroll['Grade'] < 0.7], inplace = True)

# set the webdriver to pre-supported "Chrome" and make sure it is in the project bin
from selenium.webdriver.remote.webelement import WebElement

# driver = webdriver.Chrome()
driver = webdriver.Chrome(ChromeDriverManager().install())

# use the get function to open the webpage
driver.get("https://eiaeducation.org")
# time.sleep(3)

# find the sign in element
sign_in = WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((By.LINK_TEXT, 'Sign in'))
)
sign_in.click()
#sign_in = driver.find_element_by_link_text('Sign in')

# Use action chain to click on sign in
# https://www.selenium.dev/selenium/docs/api/py/webdriver/selenium.webdriver.common.action_chains.html#module-selenium.webdriver.common.action_chains
# actions = ActionChains(driver)
# actions.click(sign_in)
# actions.perform()
# time.sleep(3)

# Find input fields
time.sleep(1) #need elements to load
username = driver.find_element_by_css_selector('input#login-email.input-block')
password = driver.find_element_by_id('login-password')

# Send keys to these fields
username.send_keys("brenton.kreiger@eiabridges.org")
password.send_keys("Sweetdreams1!")

# Use action chain to click sign in
sign_in_again = WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'login-button'))
)
sign_in_again.click()
# sign_in_again = driver.find_element_by_class_name('login-button')
# actions = ActionChains(driver)
# actions.click(sign_in_again)
# actions.perform()
# time.sleep(3)
# Enter the Course
    # n = name of the course but we need to do this by href because there are 2 SBD courses...
BP_course = WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, css_path))
)
BP_course.click()
# BP_course = driver.find_element_by_css_selector(css_path)
# actions = ActionChains(driver)
# actions.click(BP_course)
# actions.perform()
# time.sleep(3)
    # Go to Instructor tab
instruct = WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((By.LINK_TEXT, 'Instructor'))
)
instruct.click()
# instruct = driver.find_element_by_link_text('Instructor')
# actions = ActionChains(driver)
# actions.click(instruct)
# actions.perform()
# time.sleep(3)
    # Click the Data Download tab
data = WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'button.btn-link.data_download'))
)
data.click()
# data = driver.find_element_by_css_selector('button.btn-link.data_download')
# actions = ActionChains(driver)
# actions.click(data)
# actions.perform()
# time.sleep(3)
    # Click the Grade Download
grade = WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((By.NAME, 'calculate-grades-csv'))
)
grade.click()
#grade = driver.find_element_by_name('calculate-grades-csv')
#actions = ActionChains(driver)
#actions.click(grade)
#actions.perform()
BP_name = lab + '_' + c + '_' + sem + '_grade_report_' + date
print(BP_name)
    # Wait for it to appear
csvfile = WebDriverWait(driver, 1200).until(
        EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, BP_name))
    )
# time.sleep(120) # wait 120 seconds for the file to show up
    # Download the CSV using the course number "c"
#csvfile = driver.find_element_by_partial_link_text(BP_name)
# actions = ActionChains(driver)
# actions.click(csvfile)
# actions.perform()
csvfile.click()
time.sleep(3)

######### COMMENT OUT MARCH 3, 2021 ############

# # either fix certs or stop doing this
#
# # Only if enroll is not empty
# if enroll.empty == False:
#     # Click the Certificate tab
#     cert = WebDriverWait(driver, 120).until(
#             EC.presence_of_element_located((By.CSS_SELECTOR, 'button.btn-link.certificates'))
#     )
#     cert.click()
#
#         # Find exception input and put csv in it
#     # bulk = WebDriverWait(driver, 120).until(
#     #        EC.presence_of_element_located((By.ID, 'bulk-exception-upload'))
#     # )
#     # driver.find_element_by_id('browseBtn-bulk-csv').send_keys("/Users/brentonkreiger/Desktop/enroll.csv")
#     j = len(enroll)
#     for j in range(0, j):
#         student = enroll.iloc[j]['Username']
#         driver.find_element_by_id('certificate-exception').send_keys(student)
#
#         # Click the add exception button
#     # add_exceptions = WebDriverWait(driver, 120).until(
#     #         EC.presence_of_element_located((By.CLASS_NAME, 'btn-blue.upload-csv-button'))
#     # )
#     # add_exceptions.click()
#         add_exceptions = WebDriverWait(driver, 120).until(
#                 EC.presence_of_element_located((By.ID, 'add-exception'))
#         )
#         add_exceptions.click()
#         driver.find_element_by_id('certificate-exception').clear()
#
#     gen_exceptions = WebDriverWait(driver, 120).until(
#             EC.presence_of_element_located((By.ID, 'generate-exception-certificates'))
#     )
#     # throws error if no certs to make so use if statement
#     gen_exceptions.click()

########## STOP BLOCK COMMENT

# Close window
driver.close()

