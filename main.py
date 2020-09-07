#Importing packages
import subprocess as sb
import pyautogui as pag
import time
import datetime as dt
import pandas as pd
import calendar

#Every PyAutoGUI function call will wait one second after performing its action. Non-PyAutoGUI instructions will not have this pause.
pag.PAUSE = 1
#The fail-safe feature will stop the program if you quickly move the mouse as far up and left as you can.
pag.FAILSAFE = True
#Creating a flag for the driver loop
flag=True

#Getting the data
schedule = pd.read_excel('Schedule.xlsx', sheet_name='Sheet1')

#Converting days to a numerical format
weekdays={'Mo': 0, 'Tu': 1, 'We': 2, 'Th': 3, 'Fr': 4 ,'Sa': 5 , 'Su': 6}
def days(input):
    data=[]
    for x in input.split(","):
        data.append(weekdays[x])
    return data
schedule["Day (DD,DD)"]=schedule["Day (DD,DD)"].apply(days)

#Converting times to a smaller format and 10 minutes behind
def times(input):
    data=dt.datetime.combine(dt.date.min, input) - dt.datetime.combine(dt.date.min, dt.time(hour = 0, minute = 10, second = 0)) #Subtracting the time by 10 minutes using .combine that creates a timedelta data type
    return (dt.datetime.min + data).time().strftime('%H:%M') #Converting timedelta data type to time data type and formatting
schedule["Time (HH:MM)"]=schedule["Time (HH:MM)"].apply(times)

#Function to launch a meeting
def launch(meetingID):
    print("Launching Zoom")

    #Adding a delay of 1 second
    time.sleep(1)
    #Launching the app
    sb.call('C:\\Users\\rishi\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe')
    time.sleep(1)

    try:
        print("Authenticating")

        #Going to the home tab if currently at another tab
        home = pag.locateOnScreen('Button Images\\Home.png')
        if home != None:
            pos = pag.center(home)
            pag.click(pos)

        #Joining the meeting
        join1 = pag.locateOnScreen('Button Images\\Join1.png')
        pos = pag.center(join1)
        pag.click(join1)

        #Entering the ID
        pag.typewrite(meetingID)

        #Disabling audio
        audio = pag.locateOnScreen('Button Images\\Audio.png')
        pos = pag.center(audio)
        pag.click(pos)

        #Disabling video
        video = pag.locateOnScreen('Button Images\\Video.png')
        pos = pag.center(video)
        pag.click(pos)

        #Joining the meeting
        join2 = pag.locateOnScreen('Button Images\\Join2.png')
        pos = pag.center(join2)
        pag.click(pos)

        print("Authenticated")
        print("Meeting Launched")
        return True

    #Exiting the program
    except pag.FailSafeException:
        print("Program Manually Exited Through FailSafe")
        return False

while (flag==True):
    current_day = dt.datetime.today().weekday()
    current_time = dt.datetime.now().strftime('%H:%M')

    print("\nCurrent Day: ",calendar.day_name[current_day])
    print("Current Time: ", current_time)

    if current_day in schedule["Day (DD,DD)"]:
        if current_time in schedule["Time (HH:MM)"].values:
            meetingID=schedule[current_time == schedule["Time (HH:MM)"]]["ID"]
            print("Meeting Found")
            flag = launch(str(meetingID).split()[1]) #Passing only the meetingID and as a string
            time.sleep(60)

    print("No Meeting Found")
    time.sleep(60) #Check every minute







