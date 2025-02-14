import re, openpyxl, os
import pyinputplus as pyip

#Importing datetime module to use strftime method to format time data from the excel sheet
from datetime import datetime


meetingSchedule = openpyxl.load_workbook("meetingSchedule.xlsx")
meeting = []

#This function will read through the excel file storing meeting schedule and look if the room at that specific time is already booked or not
def checkAvailability(meeting, message):
    sheet = meetingSchedule.active
    nextEmptyRow = 0
    for row in range(1,51):
        buildingCellContent = sheet["A" + str(row)].value
        roomCellContent = str(sheet["B" + str(row)].value)
        dateCellContent = str(sheet["C" + str(row)].value)
        timeCellContent = str(sheet["D" + str(row)].value)

        if (buildingCellContent == meeting[0] and roomCellContent == meeting[1] and dateCellContent == meeting[2] and timeCellContent == meeting[3]):
            response = "Declined"
            print(response)

            #Text file to audit the request and response when the reservation is unsuccessful
            meetingAudit = open("meetingAudit.txt", "a")
            meetingAudit.write("Request Time: " + str(datetime.now()) + "\n" + "Request: " + message + "\n" + "Request Status: " + response + "\n\n")
            meetingAudit.close()
            return
          
    #To find the next empty row
    for row in range(1,51):
        if (sheet["A" + str(row)].value == None):
            nextEmptyRow = row
            break

    #If the room is available, it will be saved into the excel
    sheet["A" + str(nextEmptyRow)] = meeting[0]
    sheet["B" + str(nextEmptyRow)] = meeting[1]
    sheet["C" + str(nextEmptyRow)] = meeting[2]
    sheet["D" + str(nextEmptyRow)] = meeting[3]
    response = "Approved"
    print(response)

    #Text file to audit the request and response when the reservation is successful
    meetingAudit = open("meetingAudit.txt", "a")
    meetingAudit.write("Request Time: " + str(datetime.now()) + "\n" + "Request: " + message + "\n" + "Request Status: " + response + "\n\n")
    meetingAudit.close()
        


#The function will take a whole message as a string argument and parse it to find the details required for booking a meeting like:
        #The building
        #The room
        #The day
        #The date
        #The time
def messageParser(message):

    #The string is transformed to uppercase to accept for uppercase and lower case building names
    message = message.upper()

    buildingMatch, roomMatch, dateMatch, timeMatch = regexPatternMatch(message)

    building = buildingMatch.group()
    room = roomMatch.group()
    date = dateMatch.group()
    time = timeMatch.group()


    try:
        #Parsing the time string into a datetime object to change its format
        timeObject = datetime.strptime(time, "%H:%M")

        #Formatting the time into HH:MM:SS format with leading zeros
        formattedTime = timeObject.strftime("%H:%M:%S")
    except ValueError:
        print("Error formatting time")

    print(f'''Building: {building}
Room: {room}
Date: {date}
Time: {formattedTime}''')


    #All the data from the message will be stored into a list later used to check for room availability
    meeting = [building, room, date, formattedTime]
    checkAvailability(meeting, message)

#Message from the customer
#This can also be automated using a bot that goes through emails looking specificly for the ones that serve the purpose of reservations


#Defined a function for regex pattern match so that it can be called during message parsing and vlidating input from the user
def regexPatternMatch(message):
    #This will check for single uppercase alphabets that denote the building name
    buildingRegexObject = re.compile(r'\b[ABCDEFHJKMN]\b')
    roomRegexObject = re.compile(r'\d\d\d')

    #Using the most used date formats:
            #YYYY-MM-DD
            #DD-MM-YYYY
            #MM/DD/YYYY
    dateRegexObject = re.compile(r'(\d{4}-\d{2}-\d{2}|\d{2}-\d{2}-\d{4}|\d{2}/\d{2}/\d{4})')
    timeRegexObject = re.compile(r'\d+:\d\d')

    buildingMatch = buildingRegexObject.search(message)
    roomMatch = roomRegexObject.search(message)
    dateMatch = dateRegexObject.search(message)
    timeMatch = timeRegexObject.search(message)

    #Returinh the math object to callers in main() and messageParsing() to check whether they are of None type or not
    return buildingMatch, roomMatch, dateMatch, timeMatch

def main():
    userMessage = input("Request: ")
    userEmail = pyip.inputEmail(prompt = "Please enter your email to receive a copy of your confirmation: ")
    buildingMatch, roomMatch, dateMatch, timeMatch = regexPatternMatch(userMessage)

    #Using negative logic with while loop to make sure that the user has given all the information necessary to make a reservation
    while not (buildingMatch and roomMatch and dateMatch and timeMatch):
        print("\nYour message did not had all the information needed for reserving a room. Kindly check for: Building, Room, Date and Time.\n")
        userMessage = input("Request: ")
        userEmail = pyip.inputEmail(prompt = "Please enter your email to receive a copy of your confirmation: ")
        buildingMatch, roomMatch, dateMatch, timeMatch = regexPatternMatch(userMessage)

    messageParser(userMessage)

main()

#Saving the data into excel file and closing it
meetingSchedule.save("meetingSchedule.xlsx")
meetingSchedule.close()



