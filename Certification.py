import pandas as pd
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
#import GradeScrape

# Grab the course number and name from the main script "GradeScrape.py" as strings
c = '011'
course_label = 'BP' + c

# Css path for the course
# css_path = 'a[href$="' + lab + '+' + c + '+' + sem + '/course/"]'

# Set up paths (make sure available offline on drive stream)
path_grading_master = "/Volumes/GoogleDrive/Shared drives/Ambassador Program/Committee_BridgeEDU/Bridge Program Coordinator/BEDU Grading 2020-2021.xlsx"

# Read in enrollment data for the course of interest

sheet0 = pd.read_excel(path_grading_master, sheet_name=course_label) #from google stream
enroll = sheet0[['Username','Grade','Enrollment Track']]

# Drop all the honors kids
enroll.drop(enroll.index[enroll['Enrollment Track'] == 'honor'], inplace = True)

# Drop all those who aren't passing
enroll.drop(enroll.index[enroll['Grade'] < 0.7], inplace = True)

# export the username list as a csv
enroll.to_csv("/Users/brentonkreiger/Desktop/enroll.csv", columns = ['Username'], header=None, index=None)

if enroll.empty == False:
    print('yayyyy')