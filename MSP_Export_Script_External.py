####################################################################################################################################################################
#Author: Tom Diederen
#Release: 3.1, May 2024
#Summary:
#   This Python script automates repetitive operations in Microsoft Project.
#   It applies keyboard shortcuts to show different Gantt charts and takes screenshots.
#   The screenshots are of type .png and are saved in the same folder as this script.
#Setup:
#   1. Make sure alt+tab switches from this script to MSFT Project.
#   2. The values in the dictionaries "scenario1_dict" and "scenario2_dict" must match the row numbers of the tasks that they refer to in Project.
#   3. The switch view functions below (e.g. Switch to Gantt view: alt + 3) need to match the keyboard shortcut of recorded macros in project:
#       a.  Record a macro in Project that switches to each view (Gantt, task, etc.)..
#       b.  Add that macro to the quick access toolbar at the top of project
#       c.  Make sure the shortcut (alt + number) in project matches what's in the script 
#               For example: if alt + 3 is mapped to "View1" and that keyboard shortcut changes, the script needs to be updated accordingly)
####################################################################################################################################################################

'''
Packages
'''
import pyautogui
import time
from datetime import date
import datetime

'''
Line IDs
'''
scenario1_dict = {
    "scen1"            : 3,
    "scen1_product1"   : 20,
    #"scen1_product2"  : 108,
    "scen1_product3"   : 108,
    "scen1_product4"   : 260,
    "scen1_product5"   : 180,
    "scen1_product6"   : 615,
    "scen1_product7"   : 349,
    "scen1_product8"   : 514,
    #"scen1_product9a" : 780,
    "scen1_product10"  : 431,
    "scen1_product9"   : 697,
    "scen1_overview1"  : 774,
    "scen1_overview2"  : 823,
    "scen1_overview3"  : 869,
}

scenario2_dict = {
    "scen2"            : 883,
    "scen2_product1"   : 901,
    #"scen2_product2"  : 986,
    "scen2_product3"   : 986,
    "scen2_product4"   : 1133,
    "scen2_product5"   : 1055,
    "scen2_product6"   : 1471,
    "scen2_product7"   : 1218,
    "scen2_product8"   : 1376,
    #"scen2_product9a" : 1858,
    "scen2_product10"  : 1298,
    "scen2_product9"   : 1553,
    "scen2_overview1"  : 1624,
    "scen2_overview2"  : 1667,
    "scen2_overview3"  : 1705,
}

'''
Utility Methods
'''
#Switch to MS Project (alt+tab)
def alt_tab():
    pyautogui.keyDown("alt")
    pyautogui.press("tab")
    pyautogui.keyUp("alt")

#Switch to Gantt view: alt + 3. 
def gantt_view():
    time.sleep(0.15)
    pyautogui.keyDown("alt")
    pyautogui.press("3")
    pyautogui.keyUp("alt")
    time.sleep(0.15)

#Switch to task view: alt + 4.  
def task_view():
    time.sleep(0.15)
    pyautogui.keyDown("alt")
    pyautogui.press("4")
    pyautogui.keyUp("alt")
    time.sleep(0.15)

#Switch to overview1: alt + 5. 
def overview1():
    time.sleep(0.15)
    pyautogui.keyDown("alt")
    pyautogui.press("5")
    pyautogui.keyUp("alt")
    time.sleep(0.15)

#Switch to overview2: alt + 6. 
def overview2():
    time.sleep(0.15)
    pyautogui.keyDown("alt")
    pyautogui.press("6")
    pyautogui.keyUp("alt")
    time.sleep(0.15)

#Switch to overview3: alt + 7.  
def overview3():
    time.sleep(0.15)
    pyautogui.keyDown("alt")
    pyautogui.press("7")
    pyautogui.keyUp("alt")
    time.sleep(0.15)

#Switch to Si overview: alt + 8. 
def si_view():
    time.sleep(0.15)
    pyautogui.keyDown("alt")
    pyautogui.press("8")
    pyautogui.keyUp("alt")
    time.sleep(0.15)

#Switch to MB overview: alt + 9. 
def mb_view():
    time.sleep(0.15)
    pyautogui.keyDown("alt")
    pyautogui.press("9")
    pyautogui.keyUp("alt")
    time.sleep(0.15)

#Switch to DT overview: alt + 09. 
def scen2_view():
    time.sleep(0.15)
    pyautogui.keyDown("alt")
    pyautogui.press("0")
    pyautogui.press("9")
    pyautogui.keyUp("alt")
    time.sleep(0.15)    

#Switch to TPT overview
def tpt_view():
    time.sleep(0.15)
    pyautogui.keyDown("alt")
    pyautogui.press("0")
    pyautogui.press("8")
    pyautogui.keyUp("alt")
    time.sleep(0.15)    

def goto_task(line_number):
    #Open goto prompt
    pyautogui.keyDown("ctrl")
    pyautogui.press("g")
    pyautogui.keyUp("ctrl")
    #Enter line number, a small delay is needed for data input because Project sometimes doesn't capture input this fast
    time.sleep(0.15)
    pyautogui.write(f"{line_number}")
    time.sleep(0.15)
    pyautogui.press("enter")

def expand_task():
    pyautogui.hotkey("alt", "shift", "+")
    time.sleep(0.15)

def collapse_task():
    pyautogui.hotkey("alt", "shift", "-")
    time.sleep(0.15)

#Take screen shot of ROI and save.
def take_screenshot(line_id, name, view_type):
    #Go to task view and expand task, then show Gantt.
    task_view()
    goto_task(line_id)
    expand_task()
    match view_type:
        case "product":
            gantt_view()
        case "view1":
            overview1()
        case "view2":
            overview2()
        case "view3":
            overview3()
        case "Si" | "MB" | "DT" | "TPT":
            if view_type == "TPT":
                tpt_view()
            else:
                #Nested ternary to pick which view to select.
                si_view() if (view_type == "Si") else (mb_view() if view_type == "MB" else scen2_view())
            
            #Sort tasks by priority (in Project)
            #Go to view tab
            pyautogui.press("alt")
            time.sleep(0.15)
            pyautogui.press("w")
            time.sleep(0.15)
            #Sort
            pyautogui.press("r")
            time.sleep(0.15)
            pyautogui.press("t")
            time.sleep(0.15)
            pyautogui.press("s")
            time.sleep(3.00)
            pyautogui.press("enter")

        case _:
            print("ERROR INCORRECT TYPE for take_screenshot()")
    
    #Ctrl+home in case Project view is scrolled down
    pyautogui.keyDown("ctrl")
    pyautogui.press("home")
    pyautogui.keyUp("ctrl")
    
    #Project needs some time to load the Gantt, without delay, the screen capture happens too fast and the Gantt doesn't show
    time.sleep(1) 
    
    #Take screenshot of ROI containing Gantt, add timestamp to file name. MB and DT overviews are larger so they have a bigger ROI.
    if (view_type in ["product", "view1", "view2", "view3"]):
        im = pyautogui.screenshot(f"MSP_{name}_{datetime.datetime.now().strftime('%Y-%m-%d')}.png", region=(35,200, 1850, 575))
    elif(view_type in ["MB", "DT", "TPT", "Si"]):
        im = pyautogui.screenshot(f"MSP_{name}_{datetime.datetime.now().strftime('%Y-%m-%d')}.png", region=(35,200, 1850, 795))    
    
    #Go back to task view and close task again so it doesn't show in the next screen shot.
    task_view()
    if view_type == "product":
        goto_task(line_id)
        collapse_task()

#Expand all scenario 2 tasks.
def expand_all_scen2():
    task_view()
    for key in scenario2_dict:
        goto_task(scenario2_dict[key])
        expand_task()

#Expand all scenario 1 tasks.
def expand_all_scen1():
    task_view()
    for key in scenario1_dict:
        goto_task(scenario1_dict[key])
        expand_task()

#Collapse all scenario 2 tasks
def collapse_all_scen2():
    task_view()
    for key in reversed(scenario2_dict):
        goto_task(scenario2_dict[key])
        collapse_task()

#Collapse all scen1 tasks
def collapse_all_scen1():
    for key in reversed(scenario1_dict):
        goto_task(scenario1_dict[key])
        collapse_task()

#Take scenario 1 schedule screenshots
def take_scen1_screenshots():
    take_screenshot(scenario1_dict["scen1_product1"], "Scenario1_product1", "product")
    #take_screenshot(scenario1_dict["scen1_product2"], "Scenario1_product2", "product")
    take_screenshot(scenario1_dict['scen1_product3'], "Scenario1_product3", "product")
    take_screenshot(scenario1_dict['scen1_product4'], "Scenario1_product4", "product")
    take_screenshot(scenario1_dict["scen1_product5"], "Scenario1_product5", "product")
    take_screenshot(scenario1_dict["scen1_product6"], "Scenario1_product6", "product")
    take_screenshot(scenario1_dict["scen1_product7"], "Scenario1_product7", "product")
    take_screenshot(scenario1_dict["scen1_product8"], "Scenario1_product8", "product")
    #take_screenshot(scenario1_dict["scen1_product9a"], "Scenario1_product9a, "product") 
    take_screenshot(scenario1_dict["scen1_product9"], "Scenario1_product9", "product") 
    take_screenshot(scenario1_dict["scen1_product10"], "Scenario1_product10", "product")
    take_screenshot(scenario1_dict["scen1_overview1"], "Scenario1_overview1", "view1")
    take_screenshot(scenario1_dict["scen1_overview2"], "Scenario1_overview2", "view2")
    take_screenshot(scenario1_dict["scen1_overview3"], "Scenario1_overview3", "view3")

#Take scenario 2 schedule screenshots
def take_scen2_screenshots():
    take_screenshot(scenario2_dict["scen2_product1"], "Scenario2_product1", "product")
    #take_screenshot(scenario2_dict["scen2_product2"], "Scenario2_product2", "product")
    take_screenshot(scenario2_dict["scen2_product3"], "Scenario2_product3", "product")
    take_screenshot(scenario2_dict["scen2_product4"], "Scenario2_product4", "product")
    take_screenshot(scenario2_dict["scen2_product5"], "Scenario2_product5", "product")
    take_screenshot(scenario2_dict["scen2_product6"], "Scenario2_product6", "product")
    take_screenshot(scenario2_dict["scen2_product7"], "Scenario2_product7", "product")
    take_screenshot(scenario2_dict["scen2_product8"], "Scenario2_product8", "product")
    #take_screenshot(scenario2_dict["scen2_product9a"], "Scenario2_product9a", "product") 
    take_screenshot(scenario2_dict["scen2_product9"], "Scenario2_product9", "product") 
    take_screenshot(scenario2_dict["scen2_product10"], "Scenario2_product10", "product")
    take_screenshot(scenario2_dict["scen2_overview1"], "Scenario2_overview1", "view1")
    take_screenshot(scenario2_dict["scen2_overview2"], "Scenario2_overview2", "view2")
    take_screenshot(scenario2_dict["scen2_overview3"], "Scenario2_overview3", "view3")


def take_composite_screenshots():
    expand_all_scen2()
    take_screenshot(1, "Scenario2_Si_Overview", "Si")
    take_screenshot(1, "Scenario2_MB_Overview", "MB")
    take_screenshot(1, "Scenario2_DT_Overview", "DT")
    take_screenshot(1, "Scenario2_TPT_Overview", "TPT")
    collapse_all_scen2()
    expand_all_scen1()
    take_screenshot(1, "Scenario1_Si_Overview", "Si")
    take_screenshot(1, "Scenario1_MB_Overview", "MB")
    take_screenshot(1, "Scenario1_DT_Overview", "DT")
    take_screenshot(1, "Scenario1_TPT_Overview", "TPT")
    collapse_all_scen1()

'''
Script
'''
#Switch to Project
alt_tab()
#Open task view and expand scenario 1 schedule
task_view()
goto_task(scenario1_dict["scen1"])
expand_task()
#Take product screenshots and close scenario 1 schedule again
take_scen1_screenshots()
goto_task(scenario1_dict["scen1"])
collapse_task()
#Expand scen2 schedule
goto_task(scenario2_dict["scen2"])
expand_task()
#Take product screenshots and close scenario 2 schedule again
take_scen2_screenshots()
goto_task(scenario2_dict["scen2"])
collapse_task()
#Take composite screenshots
take_composite_screenshots()
print('Gantt exports completed. Have a nice day!')