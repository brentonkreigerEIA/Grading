import pandas as pd
import shutil as shit

# This is the master code for running the following functions, the procedure is as follows:
# As you loop through the course list, run Installation.py to download the csv and then run
# Grading.py to replace the grades in the google sheet.

# Issues

# Set up lists
course_program = ['EIA','EIA','EIA','EIA','EIA','EIA','EIA','EIA','EIA','EIA','EIA','EIA','EIA','EIA','EIA','EIA','EIA','EIA_WV']
course_sem = ['2021_F','2021_F','2020_F','2021_F','2021_F','2021_F','2021_F','2021_F','2022_S','2021_F','2021_F','2022_S','2022_S','2022_S','2021_F','2021_F','2021_F','2021_F']
course_list = ['001','011','021','031','BP041','BP101','111','121','BP131','201','BP211','301','311','401','411','441','501','001']
course_names = ['Bridge Program Introduction','Bridge Corps Training','Team Ambassador Training',
                'Design Engineer in Charge (DEIC)','New Member Crash Course','How To Start A Chapter',
                'Fundraising','Chapter Operations','Travel Logistics','Suspended Bridge Design',
                'Advanced Suspended Bridge Design','Construction Management','Suspended Bridge Construction',
                'Cross Cultural Competency','Project Management','Marketing','Post-Travel Reflection',
                'West Virginia Vehicular Bridge Program']
i = len(course_names)
lost_boyz = pd.DataFrame(columns=['Student ID', 'University Name', 'Email', 'Username'])

# Create back-up
path_grading_master = "/Volumes/GoogleDrive/Shared drives/Ambassador Program/Committee_BridgeEDU/Bridge Program Coordinator/BEDU Grading 2021-2022.xlsx"
path_backup = "/Volumes/GoogleDrive/Shared drives/Ambassador Program/Committee_BridgeEDU/Bridge Program Coordinator/BEDU Grading 2021-2022 - backup.xlsx"
#data_backup = pd.read_excel(path_grading_master, index_col=None)
#data_backup = pd.read_excel(path_grading_master)
#data_backup = pd.concat(pd.read_excel(path_grading_master, sheet_name=None), ignore_index=True)
shit.copy(path_grading_master, path_backup)
#data_backup.to_excel(path_backup)

for i in range(0, i):
    # Define course number and name (loop in future)
    course = course_list[i]
    name = course_names[i]
    program = course_program[i]
    semester = course_sem[i]

    # Run installation.py to get the csv files in download folder
    stream = open("Installation.py")
    read_file = stream.read()
    exec(read_file)
    print('CSV file for ' + name + ' has been downloaded')
    # print('Certificates for audit students in ' + name + ' have been generated')

    # Run grading.py
    stream = open("Grading.py")
    read_file = stream.read()
    exec(read_file)
    print('Grading for ' + name + ' has been complete')

    # Print a statement for a course finished
    print('Grade scraping done for ' + name + ' now complete, moving on to next course')
else:
    print('Grade scraping is complete')
    print('Remember to check and update the team matching spreadsheet!')
    exit(0)
