import win32com.client
import pythoncom
import os
import sys
import csv
from threading import Timer


class CreateRallyUserFormFetcher:
    def __init__(self):
        super().__init__()

        self.documentsSaved = []
        self.csvToBe = [['FirstName', 'LastName','EmailAddress', 'Project']]
        self.foundASubmission = False
        self.shouldAppendToCsv = False
        self.count = 1
        self.submissionDownloadPath = "C:/RallyUserCSV/Submissions"

    def createSubmissionFolder(self):
        try:
            os.mkdir("C:/RallyUserCSV/")
            os.mkdir("C:/RallyUserCSV/Submissions")
        except OSError:
            print("Creation of the directory %s failed" % self.submissionDownloadPath)
        else:
            print("Successfully created the directory %s " % self.submissionDownloadPath)

    def createFinishedFolderInOutlook(self):
        # CoInitializer to make Win32Com find Outlook
        pythoncom.CoInitialize()
        # Get the outlook instance from windows
        outlook = win32com.client.Dispatch(
            "Outlook.Application").GetNamespace("MAPI")
        try:
            donebox = outlook.GetDefaultFolder(6).Folders("FetchedRallyUsers")
        except:
            outlook.Folders(1).Folders("Inbox").Folders.Add("FetchedRallyUsers")

    """ Method to execute a 'task' every 'timer' second interval """
    def setFetchingInterval(self, timer, task):
        isStop = task()
        if not isStop:
            Timer(timer, self.setFetchingInterval, [timer, task]).start()

    # Method to fetch the emails containing New Rally User Submission Forms
    def fetchNewRallyUsersEmails(self):
        # CoInitializer to make Win32Com find Outlook
        pythoncom.CoInitialize()
        # Get the outlook instance from windows
        outlook = win32com.client.Dispatch(
            "Outlook.Application").GetNamespace("MAPI")

        # Get the outlook accounts tied to the client
        accounts = win32com.client.Dispatch(
            "Outlook.Application").Session.Accounts

        # Start for on all accounts
        for account in accounts:
            global inbox
            # Get the inbox folder
            inbox = outlook.GetDefaultFolder(6)
            # Folder to move the  messages too once they have been processed in this method
            donebox = outlook.GetDefaultFolder(6).Folders("FetchedRallyUsers")
            # Get all the messages in inbox
            messages = inbox.Items
            # Start for on messages 
            for message in reversed(messages):

               # Check for subject and attachment criteria
                if("New Rally Users" in message.Subject and len(message.Attachments) >= 1):
                    print("Found an email. . . ")
                    # Let the program know it found a submission
                    self.foundASubmission = True
                    # Assign attachment to a local variables
                    attachment = message.Attachments.Item(1)
                    # Create a anew path to put this file
                    path = str('C:\\RallyUserCSV\\submissions\\rally-users'+str(self.count)+'.docx')
                    # Save the attachment as a file to the path 
                    attachment.SaveASFile(path)
                    # Add the path to the list documentsSaved to know where all the files are loated for further processing
                    self.documentsSaved.append(path)
                    # Increment the class counter for file name 
                    self.count += 1
                    # Move the message to the donebox so it does not get processed again
                    message.Move(donebox)
            
            # IF a submission was found
            if self.foundASubmission:
                # CAll local method to retrieve form data
                self.retrieveFormsData()
            
        # Reset foundASubmission variable to false
        self.foundASubmission = False

    # Method to sift through the fetched files and ingest the data
    def retrieveFormsData(self):

        # Get an instance of Microsoft Word
        word = win32com.client.Dispatch("Word.Application")
        # Hide it from the user
        word.visible = False
        # For each path in the documentsSaved List
        for path in self.documentsSaved:
            # Open the word doc at this path
            wb = word.Documents.Open(path)
            # Get the first and only instance of a table
            table = wb.Tables(1)
            # Start for on each row
            for j in range(1,table.Rows.Count+1):
                # Instantiate a newRow array that will hold the data for FirstName, LastName, EMail and Project
                newRow = []
                # As long as this row is not the header (1st row) 
                if j != 1:
                    # Start for on each column
                    for i in range(1 , table.Columns.Count+1):
                        # If the lengh of the content at this cell is not 1 meaning empty
                        if(len(table.Cell(Row = j,Column = i).Range.Text.strip()) != 1):
                            # print(table.Cell(Row = j,Column = i).Range.Text.strip())
                            # Grab the stripped string text in this column
                            column = table.Cell(Row = j,Column = i).Range.Text.strip()
                            # Append that column to the newRow Array
                            newRow.append(self.removeNonAscii(str(column)))
                    # Append that newRow array to the self.csvToBe Array
                    self.csvToBe.append(newRow)
                
            # Close the Workbook
            wb.Close(True)
        # Reset documents Saved array
        self.documentsSaved  = []
        # Call method writeDataToCSV
        self.writeDataToCSV()

    # Method to remove any formatting characters from a string
    def removeNonAscii(self,s): 
        return "".join(i for i in list(s) if ord(i)<126 and ord(i)>31)

    # Method to write the data ingested from the submission documents into a new csv for Java Rally Manager to consume
    def writeDataToCSV(self):
        # Default Write mode is W+ , create if not existing
        writeMode = "w+"
        # If the csv already was created, set the writeMode to a (Append)
        if(self.shouldAppendToCsv):
            writeMode = "a"
        # Open the file
        with open("C:/RallyUserCSV/csv.csv", writeMode, newline='') as users_csv:
            # Make sure from this point on the program appends to the csv file
            self.shouldAppendToCsv = True
            # Instantiate a writer
            userCsvWriter = csv.writer(users_csv, delimiter=',')
            # Start for on each row in the csvToBe
            for row in self.csvToBe:
                # As long as the array to be written to the csv has a length not equal to 0
                if len(row) != 0:
                    # Write the row to the csv
                    userCsvWriter.writerow(row)
        # Resert csvToBe for next batch of submissions
        self.csvToBe = []

manager = CreateRallyUserFormFetcher()
manager.createFinishedFolderInOutlook()
manager.createSubmissionFolder()
manager.setFetchingInterval(20, manager.fetchNewRallyUsersEmails)
