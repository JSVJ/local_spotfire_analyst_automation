import subprocess, psutil
import time
import ait


SOURCE_FILE_PATH = r"" # GIVE THE FILE path here
ESTIMATE_APP_OPEN_TIME =  # Seconds to open application
ESTIMATE_REPORT_OPEN_TIME =  # Seconds to open report 


# File is opened as a child process
file = subprocess.Popen([SOURCE_FILE_PATH],shell=True) 
time.sleep(ESTIMATE_APP_OPEN_TIME) # Give sometime for the application to open

# Option to WORK OFFLINE

# MOVING MOUSE POINTER
# x = 546
# y = 783
# pyautogui.moveTo(x, y)
# time.sleep(1)
# pyautogui.click(x, y)

# pressing O will simulate the click of work offline mode. It's a keyboard shortcut. 
ait.press('o')

time.sleep(ESTIMATE_REPORT_OPEN_TIME)

# Closing the file
parent = psutil.Process(file.pid)
children = parent.children(recursive=True)
print(children)
child_pid = children[0].pid
print(child_pid)

subprocess.check_output("Taskkill /PID %d /F" % child_pid)