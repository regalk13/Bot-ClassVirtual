import subprocess
import pyautogui
import pyautogui as p, webbrowser
import time
from win10toast import ToastNotifier
import pandas as pd
from datetime import datetime
from os import startfile
from time import sleep
import keyboard

def zoom_meeting(meetingid, pswd):
    subprocess.call("C:/Users/usuario/AppData/Roaming/Zoom/bin/Zoom.exe")

    time.sleep(10)
    
    #clicks the join button
    join_btn = pyautogui.locateCenterOnScreen('data/join_button.png')
    pyautogui.moveTo(join_btn)
    pyautogui.click()

    # Type the meeting ID
    meeting_id_btn =  pyautogui.locateCenterOnScreen('data/meeting_id_button.png')
    pyautogui.moveTo(meeting_id_btn)
    pyautogui.write(meetingid)

    # Disables both the camera and the mic
    media_btn = pyautogui.locateAllOnScreen('data/media_btn.png')
    for btn in media_btn:
        pyautogui.moveTo(btn)
        pyautogui.click()
        time.sleep(2)

    # Hits the join button
    join_btn = pyautogui.locateCenterOnScreen('data/join_btn.png')
    pyautogui.moveTo(join_btn)
    pyautogui.click()
    
    time.sleep(5)
    #Types the password and hits enter
    meeting_pswd_btn = pyautogui.locateCenterOnScreen('data/meeting_pswd.png')
    pyautogui.moveTo(meeting_pswd_btn)
    pyautogui.click()
    pyautogui.write(pswd)
    pyautogui.press('enter')

def google_meet(meeting_link:str):
    wb = webbrowser.Chrome(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
    wb.open_new_tab(meeting_link)
    sleep(10)
    keyboard.send('tab', do_press=True, do_release=True)
    keyboard.send('tab', do_press=True, do_release=True)
    keyboard.send('tab', do_press=True, do_release=True)
    keyboard.send('tab', do_press=True, do_release=True)
    keyboard.send('enter', do_press=True, do_release=True)
    keyboard.send('tab', do_press=True, do_release=True)
    keyboard.send('enter', do_press=True, do_release=True)
    sleep(5)
    keyboard.send('tab', do_press=True, do_release=True)
    keyboard.send('tab', do_press=True, do_release=True)
    keyboard.send('tab', do_press=True, do_release=True)
    keyboard.send('enter', do_press=True, do_release=True)


def alert(lecture:str):
    toaster = ToastNotifier()
    toaster.show_toast("Class Notification", f"{lecture} class right now...")

def join_class():
    """
    Checks if there is any class on a particular date by extracting data from timetable.xlsx.
    """
    timetable = pd.read_excel(r"classtime.xlsx", sheet_name=datetime.now().strftime("%A"))
    current_time = datetime.now().strftime("%H:%M")
    current_hour = int(datetime.now().strftime("%H"))
    current_minute = int(datetime.now().strftime("%M"))

    for _, item in timetable.iterrows():
        class_hour = int(item["Class Time"].split(":")[0])
        class_minute = int(item["Class Time"].split(":")[-1])

        if class_hour == current_hour and current_minute == class_minute:
            class_name = item["Class Name"]
            joining_mode = item["Mode"].capitalize()

            if joining_mode not in ["Zoom", "Meet"]:
                raise Exception("Improper data filled in timetable.xlsx! Mode can either be 'Zoom' or 'Meet'.")

            if joining_mode == "Zoom":
                meeting_id = item["Meeting ID/Link"]
                meeting_password = item["Meeting Password"]
                alert(class_name)
                zoom_meeting(str(meeting_id), str(meeting_password))
                print(f"In {class_name} class, good luck...")

            elif joining_mode == "Meet":
                meeting_link = item["Meeting ID/Link"]
                alert(class_name)
                google_meet(str(meeting_link))
                print(f"In {class_name} class, good luck...")

    else:
        print(f"You don't have class right now {current_time}...")


if __name__ == "__main__":
    while 1:
        join_class()
        sleep(30)

