from Download_Attendance_Sheet import name
import string
import random
from openpyxl import load_workbook
import facebook
import datetime
from datetime import date
import time
import smtplib
from facepy import GraphAPI
from string import ascii_uppercase

def make_cars(Saturday, Sunday):
    #global randmess
    #randmess = "Hello World"
    global drivers

    largebooty = []
    smallbooty = []
    dissapointments = []
    
    wb = load_workbook(name + '.xlsx')
    ws = wb.active

    alphabet = list(string.ascii_uppercase)
    for cell in alphabet[1:]:
        date = str(ws[cell+'1'].value)[0:10] # get just date portion of string in this format '2019-01-25'
        # print(date)
        if Saturday == date:
            satPractice = ws[cell + '1']
            satIndex = cell
        if Sunday == date:
            sunPractice = ws[cell + '1']
            sunIndex = cell

    attendanceList = list()
    # index 0 is satGoingList
    # index 1 is satEarlyList
    # index 2 is satKGList
    # index 3 is satKGEarlyList
    # index 4 is sunGoingList
    # index 5 is sunEarlyList
    # index 6 is sunKGList
    # index 7 is sunKGEarlyList
    for i in range(8):
        attendanceList.append(list())

    for row in range(2, 29):
        satMemberResponse = satIndex + '{}'.format(row) # response of member for Saturday practice
        sunMemberReponse = sunIndex + '{}'.format(row) # response of member for Sunday practice
        member = ws['A' + str(row)].value
        if ws[satMemberResponse].value == 1:
            attendanceList[0].append(member)
        elif ws[satMemberResponse].value == 'E':
            attendanceList[1].append(member)
        elif ws[satMemberResponse].value == 'K':
            attendanceList[2].append(member)
        elif ws[satMemberResponse].value == 'KE':
            attendanceList[3].append(member)
        if ws[sunMemberReponse].value == 1:
            attendanceList[4].append(member)
        elif ws[sunMemberReponse].value == 'E':
            attendanceList[5].append(member)
        elif ws[sunMemberReponse].value == 'K':
            attendanceList[6].append(member)
        elif ws[sunMemberReponse].value == 'KE':
            attendanceList[7].append(member)
    for i in range(8):
        random.shuffle(attendanceList[i])
    print("\nSaturday Going List: ", attendanceList[0])
    print("Saturday Early Car List:", attendanceList[1])
    print("Saturday KG List:", attendanceList[2])
    print("Saturday KG and Early Car List:", attendanceList[3])
    print("Number of People Going Saturday: " + str(len(attendanceList[0]) + len(attendanceList[1]) +
                                                    len(attendanceList[2]) + len(attendanceList[3])))
    print("Number of people that don't need early car Saturday: " + str(len(attendanceList[0]) + len(attendanceList[2])))
    print("Number of people going early Saturday: " + str(len(attendanceList[1]) + len(attendanceList[3])))
    print("\nSunday Going List: ", attendanceList[4])
    print("Sunday Early Car List:", attendanceList[5])
    print("Sunday KG List:", attendanceList[6])
    print("Sunday KG and Early Car List:", attendanceList[7])
    print("Number of People Going Sunday: " + str(len(attendanceList[4]) + len(attendanceList[5]) +
                                                  len(attendanceList[6]) + len(attendanceList[7])))
    print("Number of people that don't need early car Sunday: " + str(len(attendanceList[4]) + len(attendanceList[6])))
    print("Number of people going early Sunday: " + str(len(attendanceList[5]) + len(attendanceList[7])))

    drivers = ['Amanda', 'Jackie', 'Serena', 'Rain', 'Camila']
    i = 0
    #print(len(drivers))
    #if('Amanda' in attendanceList[0]):
    #    print('True')
    #else:
    #    print('False')
    while(i < len(drivers)):
        count = 0
        if(drivers[i] in attendanceList[0]):
            print("Saturday Driver: ", drivers[i])
            while(count < len(attendanceList[0])):
                if(attendanceList[0][count] not in drivers):
                    print(" " * len(drivers[i] + '           '), attendanceList[0][count])
                    count+=1
                else:
                    count+=1
        elif(drivers[i] in attendanceList[1]):
            print("Saturday Early Driver: ", drivers[i])
            while(count < len(attendanceList[1])):
                if(attendanceList[1][count] not in drivers):
                    print(" " * len(drivers[i] + '  '), attendanceList[1][count])
                    count+=1
                else:
                    count+=1
        elif(drivers[i] in attendanceList[2]):
            print("Saturday KG Driver: ", drivers[i])
            while(count < len(attendanceList[2])):
                if(attendanceList[2][count] not in drivers):
                    print(" " * len(drivers[i] + '  '), attendanceList[2][count])
                    count+=1
                else:
                    count+=1
        elif(drivers[i] in attendanceList[3]):
            print("Saturday KG Early Driver: ", drivers[i])
            while(count < len(attendanceList[3])):
                if(attendanceList[3][count] not in drivers):
                    print(" " * len(drivers[i] + '  '), attendanceList[3][count])
                    count+=1
                else:
                    count+=1
        elif(drivers[i] in attendanceList[4]):
            print("Sunday Driver: ", drivers[i])
            while(count < len(attendanceList[4])):
                if(attendanceList[4][count] not in drivers):
                    print(" " * len(drivers[i] + '  '), attendanceList[4][count])
                    count+=1
                else:
                    count+=1
        elif(drivers[i] in attendanceList[5]):
            print("Sunday Early Driver: ", drivers[i])
            while(count < len(attendanceList[5])):
                if(attendanceList[5][count] not in drivers):
                    print(" " * len(drivers[i] + '  '), attendanceList[5][count])
                    count+=1
                else:
                    count+=1
        elif(drivers[i] in attendanceList[6]):
            print("Sunday KG Driver: ", drivers[i])
            while(count < len(attendanceList[6])):
                if(attendanceList[6][count] not in drivers):
                    print(" " * len(drivers[i] + '  '), attendanceList[6][count])
                    count+=1
                else:
                    count+=1
        elif(drivers[i] in attendanceList[7]):
            print("Sunday KG Early Driver: ", drivers[i])
            while(count < len(attendanceList[7])):
                if(attendanceList[7][count] not in drivers):
                    print(" " * len(drivers[i] + '  '), attendanceList[7][count])
                    count+=1
                else:
                    count+=1
        i+=1
#page_access_token = "EAAUZAj0nVhxQBAK4Liir0QIDrq3HdFpCTq89u8K8mHfZAXEapwLNdHZASi0ZCZCq9iqiAHfH7LXsa9JFut4W5WaAAskE0hZB8IgQb6ZCh0SyGwSUW2b87nBAKQT8VRMYctFwxPaY93bV05oDRlRvL4xasmjkFpKZCz9Eu9ZBuf9BKOA7Vqwdmo8Tr9L6vMMPEL3s6L5d6nwBjsKQrg2woISZAIDCj6uFmATJiWZBnG2LgEp0QZDZD"
#graph = facebook.GraphAPI(page_access_token)
#facebook_page_id = "521974088432565"
#graph.put_object(facebook_page_id, "feed",  message='Hello world')
global today
global day
global nextsat
global nextsun
def what_fucking_day_is_it():
    today = date.today()
    week_days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    day = datetime.datetime.today().weekday()
    print(week_days[day])
    print("Today's date:", today)
    print(day)
    if(day == 5 or day == 6):
        print("Today is Weekend")
    else:
        print("Today is Weekday")
        nextsat = today + datetime.timedelta(5-day)
        nextsun = today + datetime.timedelta(6-day)
        print("Next Saturday is: ", nextsat)
        print("Next Sunday is: ", nextsun)
def judgement_time():
    current_time = time.strftime("%H:%M:%S")
    post_time = "00:00:00"
    nextfri = today + datetime.timedelta(4-day)
    if(today == nextfri and current_time == post_time):
        Brianna_Wants_To_Be_Emailed_Instead()
def Brianna_Wants_To_Be_Emailed_Instead():
    what_fucking_day_is_it()
    make_cars(nextsat, nextsun)
    # creates SMTP session
    s = smtplib.SMTP('smtp.gmail.com', 587)

    # start TLS for security
    s.ehlo()
    s.starttls()
    s.ehlo()

    # Authentication
    s.login("usc.dragonboat@gmail.com", "fightforgold")

    # message to be sent
    #message = "Hello Brianna, \n    This is a test"
    counter = 0
    message = []
    message.append("List of Drivers: ")
    while(counter < len(drivers)):
        #people = ""
        #for driver in drivers:
        #    people += driver + " "
        #message.append("List of Drivers: " + people)
        message.append(drivers[counter])
        counter+=1
    print(message)
    # sending the mail
    #s.sendmail("edisonsiu8899@gmail.com", "sunb@usc.edu", message)
    #s.sendmail("edisonsiu8899@gmail.com", "edisonsiu8899@gmail.com", message)
    s.sendmail("usc.dragonboat@gmail.com", "sunb@usc.edu", message)

    # terminating the session
    s.quit()

def main():
    #make_cars('2020-02-08', '2020-02-09')
    #what_fucking_day_is_it()
    #Brianna_Wants_To_Be_Emailed_Instead()
    judgement_time()
main()
