"""
******************************************************************************

DANY InPho: Import Tool v1.0.0

Robert Chavez
All rights reserved

******************************************************************************
"""


import pandas as pd
import Tkinter, tkFileDialog
import pyodbc
import sqlalchemy as sql
import getpass
import datetime
import tkMessageBox
import os, os.path
import shutil
import sys

class interface(Tkinter.Tk):
    def __init__(self, parent):
        Tkinter.Tk.__init__(self, parent)
        self.parent = parent
        self.initialize()
    
    
    def initialize(self):
        self.grid()

        # set label for Matter number box
        matterLabelVar = Tkinter.StringVar()
        matterLabel = Tkinter.Label(self, textvariable=matterLabelVar, anchor="w", 
                                    justify="left", fg="black", font="Constantia 10 bold")
        matterLabel.grid(column=0,columnspan=2,row=1, sticky='W')
        matterLabelVar.set(u"Enter DANY matter/case number:")

        #setup the matterNumber box        
        self.matterVariable = Tkinter.StringVar()
        self.matterVariable.set("Matter number")
        self.matterBox = Tkinter.Entry(self, textvariable = self.matterVariable)
        self.matterBox.grid(column=0,row=2,sticky='EW',padx=10)

        #setup the partial box        
        self.partialVar = Tkinter.StringVar()
        self.partialBox = Tkinter.Entry(self, textvariable = self.partialVar)
        self.partialBox.grid(column=1,row=2,sticky='EW',padx=10)

        # pressing Return does something
        self.matterBox.bind("<Return>", self.OnPressEnter)
        
        #configure the verifyMatterNumber button
        verifyMatterNumberButton = Tkinter.Button(self,text=u"Verify Matter Number", 
                                                  command=self.OnButtonClick)
        verifyMatterNumberButton.grid(column=2,row=2)
        verifyMatterNumberButton.bind("<Return>", self.OnPressEnter)


        # set label for compliance method
        complianceLabelVar = Tkinter.StringVar()
        complianceLabel = Tkinter.Label(self, textvariable=complianceLabelVar, 
                                        anchor="w", justify="left", fg="black", 
                                        font="Constantia 10 bold")
        complianceLabel.grid(column=0,row=11, sticky='W')
        complianceLabelVar.set(u"Compliance Type:")

        # configure the complianceMethod radio button
        self.complianceMethodVariable = Tkinter.StringVar()
        self.complianceMethodVariable.set("Grand Jury Subpoena")
        Tkinter.Radiobutton(self, text="Grand Jury Subpoena", 
                            variable=self.complianceMethodVariable, 
                            value="Grand Jury Subpoena").grid(column=0,row=12, sticky='W')
        Tkinter.Radiobutton(self, text="Trial Subpoena", 
                            variable=self.complianceMethodVariable, 
                            value="Trial Subpoena").grid(column=0,row=13, sticky='W')
        Tkinter.Radiobutton(self, text="Investigation Request", 
                            variable=self.complianceMethodVariable, 
                            value="Investigation Request").grid(column=0,row=14, sticky='W')
        
               
#        # set label for DOC Tracking Number
#        DOCLabelVar = Tkinter.StringVar()
#        DOCLabel = Tkinter.Label(self, textvariable=DOCLabelVar, anchor="w", 
#                                 justify="left", fg="black", font="Constantia 10 bold")
#        DOCLabel.grid(column=1,row=11, sticky='W')
#        DOCLabelVar.set(u"DOC Tracking Number:")
#        
#        # configure the DOCTrackingNumber text box
#        self.DOCTrackingVar = Tkinter.StringVar()
#        DOCTrackingBox = Tkinter.Entry(self, textvariable=self.DOCTrackingVar)
#        DOCTrackingBox.grid(column=1,row=12,sticky='W',padx=10)
        
        # set label for other options
        otherOptionsVar = Tkinter.StringVar()
        otherOptionsLabel = Tkinter.Label(self, textvariable=otherOptionsVar, anchor="w", 
                                 justify="left", fg="black", font="Constantia 10 bold")
        otherOptionsLabel.grid(column=0,columnspan=2,row=20, sticky='W')
        otherOptionsVar.set(u"Other import options:")        
        
        # add option to add to DANY SQL
        self.addToDANYSQLVar = Tkinter.IntVar()
        self.addToDANYSQLVar.set(1)

        
    
        # add option for multi disks
        self.multiDiskVar = Tkinter.IntVar()
        self.multiDiskVar.set(0)
        multiDiskCB = Tkinter.Checkbutton(text='This a multi-CD import ("Part 1 of 2" is written on this CD)', variable=self.multiDiskVar)
        multiDiskCB.grid(column=0,row=35, columnspan=3,sticky='W')
        
        # add option to add to x drive
        self.copyToXVar = Tkinter.IntVar()
        self.copyToXVar.set(1)

    

 # Uncomment this to turn on dev options        
#        addToDANYSQLCB = Tkinter.Checkbutton(self, text="Add to DANY database?", 
#                                             variable=self.addToDANYSQLVar)
#        addToDANYSQLCB.grid(column=0,row=30,sticky='W')
#        copyToXCB = Tkinter.Checkbutton(self, text="Copy audio files to X: drive?", 
#                                             variable=self.copyToXVar)    
#        copyToXCB.grid(column=1,row=30,sticky='W')


        
        # label bar on top of window
        self.labelVariable = Tkinter.StringVar()
        label = Tkinter.Label(self,textvariable=self.labelVariable, anchor="w",
                              fg="white",bg="#0154a0",font="Constantia 12 bold")
        label.grid(column=0,row=0,columnspan=3,sticky='EW')
        self.labelVariable.set(u"   Welcome to DANY InPho   ")

        # status bar on bottom of window
        self.statusVariable = Tkinter.StringVar()
        self.status = Tkinter.Label(self,textvariable=self.statusVariable, anchor="w",
                              fg="#f6731b",bg="#0154a0",font="Constantia 11")
        self.status.grid(column=0,row=52,columnspan=4,sticky='EW')
        
        # ADA label
        self.adaVariable = Tkinter.StringVar()
        self.adaVariable.set(u"Docket, Indictment, ArrestID, ICMS, TDI (F-number)")
        ada = Tkinter.Label(self, textvariable=self.adaVariable, anchor="w",
                            justify="left",fg="black")
        ada.grid(column=0,row=5,columnspan=3,sticky='w',padx=10)
        
        # Email label
        self.emailVariable = Tkinter.StringVar()
        email = Tkinter.Label(self, textvariable=self.emailVariable, anchor="w",
                              justify="left",fg="black")
        email.grid(column=0,row=7,columnspan=3,sticky='w',padx=10)
        
        # Bureau label
        self.bureauVariable = Tkinter.StringVar()
        bureau = Tkinter.Label(self, textvariable=self.bureauVariable, anchor="w",
                               justify="left",fg="black")
        bureau.grid(column=0,row=6,columnspan=3, sticky='w',padx=10)
        
        # Case name label
        self.caseNameVariable = Tkinter.StringVar()
        caseName = Tkinter.Label(self, textvariable=self.caseNameVariable, anchor="w",
                                 justify="left",fg="black")
        caseName.grid(column=0,row=8,columnspan=3,sticky='w',padx=10)
      
        # configure the CSV button
        CSVButton = Tkinter.Button(self,text=u"Import Calls", command=self.OnCSVImport,
                                   font="Constantia 12 bold")
        CSVButton.grid(column=0,row=50,columnspan=3)
        
        # empty bar above DOC and type radio buttons
        self.blankVar1 = Tkinter.StringVar()
        blank1 = Tkinter.Label(self,textvariable=self.blankVar1, anchor="w"
                              ,font="Constantia 10 bold")
        blank1.grid(column=0,row=9,columnspan=3,sticky='EW')

        # empty bar above other options
        self.blankVar1 = Tkinter.StringVar()
        blank1 = Tkinter.Label(self,textvariable=self.blankVar1, anchor="w"
                              ,font="Constantia 10 bold")
        blank1.grid(column=0,row=19,columnspan=3,sticky='EW')

        # empty bar on bottom of window
        self.blankVar2 = Tkinter.StringVar()
        blank2 = Tkinter.Label(self,textvariable=self.blankVar2, anchor="w"
                              ,font="Constantia 10 bold")
        blank2.grid(column=0,row=51,columnspan=3,sticky='EW')
        
        # empty bar above import button
        self.blankVar3 = Tkinter.StringVar()
        blank3 = Tkinter.Label(self,textvariable=self.blankVar3, anchor="w"
                              ,font="Constantia 10 bold")
        blank3.grid(column=0,row=49,columnspan=3,sticky='EW')

        
        # setting up window size and other options
        x = (self.winfo_screenwidth() - self.winfo_reqwidth()) / 2 - 210
        y = (self.winfo_screenheight() - self.winfo_reqheight()) / 2 - 219
                    
        self.grid_columnconfigure(0,weight=1)
        self.resizable(False,False)
        self.update()
        
        #dev option size
        #self.geometry("420x462+%d+%d" % (x,y))
        
        # regular size
        self.geometry("420x438+%d+%d" % (x,y))
        
        
        # Mac size
        # self.geometry("495x513+%d+%d" % (x,y))
        
        
        self.matterBox.focus_set()
        self.matterBox.selection_range(0, Tkinter.END)
                      
        
    def OnButtonClick(self):
        self.matterBox.focus_set()
        self.matterBox.selection_range(0, Tkinter.END)

        # Verifies case info for the matterVariable and displays the other case datapoints
        importTool.validateMatterNumber(self.matterVariable.get())

        if importTool.validatedToggle == True:
            self.status.configure(fg="green")
            self.statusVariable.set("Matter number OK!")

        
    def OnPressEnter(self,event):
        self.matterBox.focus_set()
        self.matterBox.selection_range(0, Tkinter.END)
        
        # prints out case info for the number
        importTool.validateMatterNumber(self.matterVariable.get())
      
        if importTool.validatedToggle == True:
            self.status.configure(fg="green")
            self.statusVariable.set("Matter number OK!")

    def OnCSVImport(self):
        if importTool.validatedToggle == False:
            tkMessageBox.showerror("InPho", "You did not validate your matter number. Please enter a valid case number before importing your calls.")
        else:    
            if self.multiDiskVar.get() == 1:
                firstQuestion = tkMessageBox.askokcancel("InPho", "You have selected a multi-disk import. This option is only used if the CD you are trying to import is labeled 'Part 1 of X'. If that's correct, please confirm disk 1 is in the drive and click 'OK'. If not, click 'Cancel'.")                    
                if firstQuestion is True:
                    self.MultiCDImport()
                else:
                    self.multiDiskVar.set(0)
                    self.addToDANYSQLVar.set(1)
                    self.copyToXVar.set(1)
            else:
                try:
                    if self.addToDANYSQLVar.get()==1:
                        importTool.ImportCSV()
                    if (self.copyToXVar.get() == 1) and (importTool.corruptCSV == False):
                        importTool.CopyFilesToX()
                        self.copyToXVar.set(1)
            
                except IOError,e:
                    tkMessageBox.showerror("InPho", "Please make sure a CD is in your computer\n\n" + str(e))
        
    def MultiCDImport(self):
        importTool.CopyFilesToX()
        self.multiDiskVar.set(0)
        self.addToDANYSQLVar.set(1)
        self.OnCSVImport()

class ImportTool:
    
    validatedToggle = False
    corruptCSV = False
    src = "D:\\"
    
    def CopyFilesToX(self):
        newMatter = importTool.gMatterNumber.replace("/", "-").upper()
                
        dest = r'\\DANY.NYCNET\DANYXDrive\Rikers Calls' + "\\" + str(importTool.gBureau) + '\\' + str(importTool.gADAUsername)  + "\\" + str(newMatter) + "\\"
        print "Calls copied here: " + dest 
        
        
        # create the file location on the X drive
        if not os.path.exists(dest):
            os.makedirs(dest)

        # copy the files from the CD to X drive
        src_files = os.listdir(self.src)
        
        numfiles = len([f for f in os.listdir(self.src) if f[0] != '.'])
        count = 1
        
        # copy loop with counter for status bar
        for file_name in src_files:
            full_file_name = os.path.join(self.src, file_name)
            if (os.path.isfile(full_file_name)):
                extension = os.path.splitext(full_file_name)[1]
                if (extension == ".flac"):
                    try:
                        shutil.copy2(full_file_name, dest)
                        self.status = str(count) + "/" + str(numfiles) + " calls have been copied"
                        app.statusVariable.set(self.status)
                        app.update()
                        count += 1
                    except IOError,e:
                        if (e.errno == 13):
                            self.status = str(count) + "/" + str(numfiles) + " calls have been copied"
                            app.statusVariable.set(self.status)
                            app.update()
                            count += 1
                            continue
                        else:
                            tkMessageBox.showerror("Copy error", str(e))
        
        app.statusVariable.set("Insert next CD to import additional calls for this matter number")
        
        if app.multiDiskVar.get() == 0:
        # notify user that calls were imported successfully
            message = str(count) + " calls successfully imported. Please open up DANY InPho in Excel to generate a call log."
            tkMessageBox.showinfo("InPho", message)
        else:
            tkMessageBox.showinfo("InPho", "Please insert disk 2 of 2 for this multi-disk import and click 'OK' to finish the import")

        
    def AskForCSVFile(self):
        # Request file location from user
        root = Tkinter.Tk()
        root.withdraw()
        root.overrideredirect(True)
        root.geometry('0x0+0+0')
        root.deiconify()
        root.lift()
        root.focus_force()
        
        #ask user for file location
        self.file_path = tkFileDialog.askopenfilename(parent=root)
        root.destroy()



    def ImportCSV(self):
        self.file_path = self.src + "RECORDING_FILE.CSV"        
        
        # Uncomment this to upload a repaired csv file
        # self.file_path = "C:/Users/ChavezR/Desktop/RECORDING_FILE.CSV"
    
        # Import the CSV file, set column names, fix date data, start on row 5
        df = pd.read_csv(str(self.file_path), names=['NYSID', 'BAC', 'Facility', 
                         'Extension', 'HousingArea', 'DateTime', 'NumberDialed'
                         , 'CallType', 'Duration', 'LastName', 'FirstName'], 
                         parse_dates=[5], skiprows=5, infer_datetime_format=True)
        
        if str(df.loc[0,'NYSID']) == "FROM ARCH.CALL_LOG CL":
            tkMessageBox.showerror("InPho", "CSV file is corrupt, please contact Robert Chavez in CSU.")
            self.corruptCSV = True
            
        else:
            self.corruptCSV = False
            try:
                # Create a new column that concatenates 'LastName' and 'FirstName'
                df['InmateName'] = df['LastName'] + ', ' + df['FirstName']
                # Delete 'LastName' and 'FirstName' columns
                df = df.drop(['LastName', 'FirstName'], 1)
                numberOfCalls = len(df.index) 
    
                # create and set the FileName column
                temp = pd.DatetimeIndex(df['DateTime'])
                df['Year'] = temp.year
                df['Day'] = temp.day
                df['Month'] = temp.month
                df['Hour'] = temp.hour
                df['Minute'] = temp.minute
                df['Second'] = temp.second
                
    
                dateOfChange = datetime.date(2014, 03, 28)
                print numberOfCalls 
    
                df.loc[df['DateTime'] < dateOfChange, 'FileName'] = df.BAC.map(str) + '_' + df.Year.map("{:04}".format, str) + df.Month.map("{:02}".format, str) + df.Day.map("{:02}".format, str) + df.Hour.map("{:02}".format, str) + df.Minute.map("{:02}".format, str) + df.Second.map("{:02}".format, str)
                df.loc[df['DateTime'] > dateOfChange, 'FileName'] = df.BAC.map(str) + '_' + df.Year.map("{:04}".format, str) + df.Month.map("{:02}".format, str) + df.Day.map("{:02}".format, str) + df.Hour.map("{:02}".format, str) + df.Minute.map("{:02}".format, str) + df.Second.map("{:02}".format, str) + "_" + df.NumberDialed.map(str)
                
                
                # create and set the MatterName and Bureau column
                if importTool.isDocket:
                    df['MatterName'] = "People v. " + str(importTool.defendant.title()).strip()
                    df['Bureau'] = str(importTool.bureau)
                elif importTool.isIndictment:
                    df['MatterName'] = "People v. " + str(importTool.defendant.title()).strip()
                    df['Bureau'] = str(importTool.bureau)
                elif importTool.isArrestID:
                    df['MatterName'] = "People v. " + str(importTool.defendant.title()).strip()
                    df['Bureau'] = str(importTool.bureau)
                elif importTool.isICMS:
                    df['MatterName'] = str(importTool.casename)
                    df['Bureau'] = str(importTool.Bureau)
                elif importTool.isFNumber:
                    df['MatterName'] = str(importTool.InvestigationName)
                    df['Bureau'] = str(importTool.BureauFullName)
                else:
                    df['MatterName'] = "Unknown matter name"
                    df['Bureau'] = "Unknown bureau name"
        
                # create and set the Contact, ContactUsername, and MatterNumber columns
                df['Contact'] = str(importTool.ADA)
                df['ContactUsername'] = str(importTool.ADAUsername)
                
                tmpMatter = app.matterVariable.get().upper()
                tmpPartial = app.partialVar.get().upper()

                if len(tmpPartial) > 0:
                    df['MatterNumber'] = tmpMatter + tmpPartial
                else:
                    df['MatterNumber'] = tmpMatter

                df['ComplianceMethod'] = app.complianceMethodVariable.get()
                df['NumberOfDisks'] = 1
                df['DOCTrackingNumber'] = ""
                df['ProcessedBy'] = getpass.getuser()
                df['DateTimeProcessed'] = datetime.datetime.now()

                #change column types to match DB                
                df['BAC'] = df['BAC'].astype(object)
                df['NumberDialed'] = df['NumberDialed'].astype(object)
                df['Duration'] = df['Duration'].astype(object) 
                df['Extension'] = df['Extension'].astype(object)    
                df['NumberOfDisks'] = df['NumberOfDisks'].astype(object)    
    
                # Reorder columns appropriately
                df = df[['InmateName', 'NYSID', 'BAC', 'DateTime', 'NumberDialed', 
                         'Duration', 'Facility', 'Extension', 'FileName', 'MatterName', 'MatterNumber', 'Contact', 'ContactUsername', 'Bureau', 
                         'ComplianceMethod', 'DOCTrackingNumber', 'NumberOfDisks', 'ProcessedBy', 'DateTimeProcessed']]
        
                # Save DataFrame to CSV
                # df.to_csv("recording_file_revised.csv", index=False)
        
                    
                # SQL direct insert
#               engine = sql.create_engine('mssql+pyodbc://DTC-SQL10/MasterCallDb')
#               df.to_sql('InmateCalls', engine, if_exists='append')

                    
                    # SQL using stored procedure

                x = 0                   
                    
                while x < numberOfCalls:
                        
#                       sqlInmateName = 'N' + df.loc[x, 'InmateName']
#                       sqlNYSID = 'N' + df.loc[x, 'NYSID']
#                       sqlBAC = 'N' + df.loc[x, 'BAC']
#                       sqlDateTime = 'N' + df.loc[x, 'DateTime']
#                       sqlNumberDialed = 'N' + df.loc[x, 'NumberDialed']
#                       sqlDuration = 'N' + df.loc[x, 'Duration']
#                       sqlFacility = 'N' + df.loc[x, 'Facility']
#                       sqlExtension = 'N' + df.loc[x, 'Extension']
#                       sqlFileName = 'N' + df.loc[x, 'FileName']
#                       sqlMatterName = 'N' + df.loc[x, 'MatterName']
#                       sqlMatterNumber = 'N' + df.loc[x, 'MatterNumber']
#                       sqlContact = 'N' + df.loc[x, 'Contact']
#                       sqlContactUsername = 'N' + df.loc[x, 'ContactUsername']
#                       sqlBureau = 'N' + df.loc[x, 'Bureau']
#                       sqlComplianceMethod = 'N' + df.loc[x, 'ComplianceMethod']
#                       sqlDOCTrackingNumber = 'N' + df.loc[x, 'DOCTrackingNumber']
#                       sqlNumberOfDisks = 'N' + df.loc[x, 'NumberOfDisks']
#                       sqlProcessedBy = 'N' + df.loc[x, 'ProcessedBy']
#                       sqlDateTimeProcessed = 'N' + df.loc[x, 'DateTimeProcessed']

                    sqlInmateName = str(df.loc[x, 'InmateName'])
                    sqlNYSID = df.loc[x, 'NYSID']
                    sqlBAC = df.loc[x, 'BAC']
                    sqlDateTime = df.loc[x, 'DateTime']
                    sqlNumberDialed = df.loc[x, 'NumberDialed']
                    sqlDuration = df.loc[x, 'Duration']
                    sqlFacility = df.loc[x, 'Facility']
                    sqlExtension = df.loc[x, 'Extension']
                    sqlFileName = df.loc[x, 'FileName']
                    sqlMatterName = str(df.loc[x, 'MatterName'])
                    sqlMatterNumber = df.loc[x, 'MatterNumber']
                    sqlContact = str(df.loc[x, 'Contact'])
                    sqlContactUsername = df.loc[x, 'ContactUsername']
                    sqlBureau = df.loc[x, 'Bureau']
                    sqlComplianceMethod = df.loc[x, 'ComplianceMethod']
                    sqlDOCTrackingNumber = df.loc[x, 'DOCTrackingNumber']
                    sqlNumberOfDisks = df.loc[x, 'NumberOfDisks']
                    sqlProcessedBy = df.loc[x, 'ProcessedBy']
                    sqlDateTimeProcessed = str(df.loc[x, 'DateTimeProcessed'])
                    sqlInmateName = sqlInmateName.replace("'","''")
                    sqlMatterName = sqlMatterName.replace("'","''")
                    sqlContact = sqlContact.replace("'","''")    
                    sqlDateTimeProcessed, temp = sqlDateTimeProcessed.split('.')

                    cnxn = pyodbc.connect('Driver={SQL Server};SERVER=dtc-sql10;DATABASE=MasterCallDb;Trusted_Connection=Yes;')
                    cursor = cnxn.cursor()

                    SQL = "EXEC [dbo].[usp_iInmateCalls] \n @InmateName = '" + str(sqlInmateName) + "', \n @NYSID = '" + str(sqlNYSID) + "', \n @BAC = '" + str(sqlBAC) + "', \n @DateTime = '" + str(sqlDateTime) + "', \n @NumberDialed = '" + str(sqlNumberDialed) + "', \n @Duration = " + str(sqlDuration) + ", \n @Facility = '" + str(sqlFacility) + "', \n @Extension = " + str(sqlExtension) + ", \n @FileName = '" + str(sqlFileName) + "', \n @MatterName = '" + str(sqlMatterName) + "', \n @MatterNumber = '" + str(sqlMatterNumber) + "', \n @Contact = '" + str(sqlContact) + "', \n @ContactUsername = '" + str(sqlContactUsername) + "', \n @Bureau = '" + str(sqlBureau) + "', \n @ComplianceMethod = '" + str(sqlComplianceMethod) + "', \n @DOCTrackingNumber = '" + str(sqlDOCTrackingNumber) + "', \n @NumberOfDisks = " + str(sqlNumberOfDisks) + ", \n @ProcessedBy = '" + str(sqlProcessedBy) + "', \n @DateTimeProcessed = '" + str(sqlDateTimeProcessed) + "'"
                        
                    cursor.execute(SQL)
                    cursor.commit()            

                    x += 1
       
       
            except IOError, e:
                tkMessageBox.showerror("InPho", "Error: " + str(e))
            

    # Validation for matterNumber
    def validateMatterNumber(self, matterNumber):

        #dev server
        #cnxn = pyodbc.connect('Driver={SQL Server};SERVER=dtc-planstgsql1;DATABASE=PLANDMS;Trusted_Connection=Yes;')
        
        cnxn = pyodbc.connect('Driver={SQL Server};SERVER=dtc-dmssql1-clu;DATABASE=DMS;Trusted_Connection=Yes;')
        cursor = cnxn.cursor()
            
        # clean the matterNumber of any extra spaces
        matterNumber = matterNumber.strip()
        firstChar = matterNumber[:1]
        secondChar = matterNumber[1:2]
        
        # identifying the matterNumber type (Indictment, Docket, Arrest ID, FNum, ICMS)
        self.isDocket = (("N" in matterNumber) or ("n" in matterNumber)) and (matterNumber.startswith("20") or matterNumber.startswith("19"))
        self.isIndictment = "/" in matterNumber
        self.isArrestID = (("M" in matterNumber) or ("m" in matterNumber)) and len(matterNumber) == 9
        self.isICMS = firstChar.isalpha() and (firstChar != "F" and firstChar != "f") and (secondChar.isalpha() == False)
        self.isFNumber = (("F" == firstChar) or ("f" == firstChar)) and len(matterNumber) == 12
        


        #####  DOCKET  #####
        # validate docket number using stored procedure        
        if self.isDocket:
            print "Its a docket number"
            docketResults = cursor.execute("""
                
                DECLARE @return_value int
                EXEC @return_value = usp_ValidateDocIndArr @docket = ?;           
                SELECT @return_value
                
               """
               , matterNumber)
            self.matterResults = docketResults

            # assign all the case info to row variable
            row = self.matterResults.fetchone()

            if row:
                # assign case info to variables
                self.docket = row.docket 
                self.latestindictment = row.latestindictment
                self.arrestID = row.arrestid
                self.defendant = row.defendant
                self.ADA = row.ADA
                self.ADAUsername = row.ADAUsername
                self.email = row.email
                self.bureau = row.Bureau
                
                # clear status variable
                app.statusVariable.set("")
                app.status.configure(fg="#f6731b")

                # matterNumber has been validated                
                self.validatedToggle = True
                
                # set some global variables
                self.gADA = row.ADA
                self.gADAUsername = row.ADAUsername
                self.gBureau = row.Bureau
                self.gMatterNumber = matterNumber
            
            #gracefully handle no SQL results for verification
            else:
                app.statusVariable.set("That matter number does not exist, please try again" )
                app.status.configure(fg="Red", font="Constantia 12 bold")
                print "No SQL data"
                self.docket = "N/A"        
                self.latestindictment = "N/A"
                self.arrestID = "N/A"
                self.defendant = "N/A"
                self.ADA = "N/A"
                self.ADAUsername = "N/A"
                self.email = "N/A"
                self.bureau = "N/A"
                
                self.validatedToggle = False

            # set labels for all the case info
            app.adaVariable.set("ADA: " + self.ADA)  
            app.emailVariable.set("Email: " + self.email)
            app.bureauVariable.set("Bureau: " + self.bureau)
            app.caseNameVariable.set("Case Name: People v. " + self.defendant)
                
            
        #####  INDICTMENT  #####       
        # validate indictment number using stored procedure        
        elif self.isIndictment:
            
            # fix leading zero issue            
            if(len(matterNumber)<10):
                if(len(matterNumber)==9):
                    matterNumber = '0' + matterNumber
                elif(len(matterNumber)==8):
                    matterNumber = '00' + matterNumber
                elif(len(matterNumber)==7):
                    matterNumber = '000' + matterNumber
                elif(len(matterNumber)==6):
                    matterNumber= '0000' + matterNumber
                else:
                    matterNumber = matterNumber
            
            print "Its an indictment number"
            indictmentResults = cursor.execute("""
                
                DECLARE @return_value int
                EXEC @return_value = usp_ValidateDocIndArr @indictment = ?;
                SELECT @return_value                
                
                """
                , matterNumber)
            self.matterResults = indictmentResults
            
            # assign all the case info to row variable
            row = self.matterResults.fetchone()
            
            if row:
            # assign case info to variables
                self.docket = row.docket        
                self.latestindictment = row.latestindictment
                self.arrestID = row.arrestid
                self.defendant = row.defendant
                self.ADA = row.ADA
                self.ADAUsername = row.ADAUsername
                self.email = row.email
                self.bureau = row.Bureau   

                # clear status variable
                app.statusVariable.set("")
                app.status.configure(fg="#f6731b")

                # matterNumber has been validated                
                self.validatedToggle = True

                # set some global variables
                self.gADA = row.ADA
                self.gADAUsername = row.ADAUsername
                self.gBureau = row.Bureau
                self.gMatterNumber = matterNumber

            # gracefully handle no SQL results for verification
            else:
                print "No SQL data"
                app.statusVariable.set("That matter number does not exist, please try again")
                app.status.configure(fg="Red", font="Constantia 12 bold")
                self.docket = "N/A"        
                self.latestindictment = "N/A"
                self.arrestID = "N/A"
                self.defendant = "N/A"
                self.ADA = "N/A"
                self.ADAUsername = "N/A"
                self.email = "N/A"
                self.bureau = "N/A"
                
                self.validatedToggle = False
        
            # set labels for all the case info
            app.adaVariable.set("ADA: " + self.ADA)
            app.emailVariable.set("Email: " + self.email)
            app.bureauVariable.set("Bureau: " + self.bureau)
            app.caseNameVariable.set("Case Name: People v. " + self.defendant)


        #####  ARRESTID  #####
        # validate ArrestID using stored procedure
        elif self.isArrestID:
            print "Its an Arrest ID"
            arrestIDResults = cursor.execute("""
                
                DECLARE @return_value int
                EXEC @return_value = usp_ValidateDocIndArr @arrestid = ?;
                SELECT @return_value                
                
                """
                , matterNumber)
            self.matterResults = arrestIDResults
            
            # assign all the case info to row variable
            row = self.matterResults.fetchone()

            if row:
                # assign case info to variables
                self.docket = row.docket        
                self.latestindictment = row.latestindictment
                self.arrestID = row.arrestid
                self.defendant = row.defendant.title()
                self.ADA = row.ADA
                self.ADAUsername = row.ADAUsername
                self.email = row.email
                self.bureau = row.Bureau   

                # clear status variable
                app.statusVariable.set("")
                app.status.configure(fg="#f6731b")

                # matterNumber has been validated                
                self.validatedToggle = True
            
                # set some global variables
                self.gADA = row.ADA
                self.gADAUsername = row.ADAUsername
                self.gBureau = row.Bureau
                self.gMatterNumber = matterNumber
            
            
            # gracefully handle no SQL results for verification
            else:
                print "No SQL data"
                app.statusVariable.set("That matter number does not exist, please try again")
                app.status.configure(fg="Red", font="Constantia 12 bold")
                self.docket = "N/A"        
                self.latestindictment = "N/A"
                self.arrestID = "N/A"
                self.defendant = "N/A"
                self.ADA = "N/A"
                self.ADAUsername = "N/A"
                self.email = "N/A"
                self.bureau = "N/A"
                
                self.validatedToggle = False
        
            # set labels for all the case info
            app.adaVariable.set("ADA: " + self.ADA)
            app.emailVariable.set("Email: " + self.email)
            app.bureauVariable.set("Bureau: " + self.bureau)
            app.caseNameVariable.set("Case Name: People v. " + self.defendant)


        #####  ICMS  #####
        # validate ICMS number using stored procedure
        elif self.isICMS:
            print "Its an ICMS number"
            matterNumber = matterNumber[(((matterNumber.find('-0'))*(-1))-1):]

            #cnxn = pyodbc.connect('Driver={SQL Server};Server=dtc-planstgsql1;Database=PLANKoin;Trusted_Connection=Yes;')

            cnxn = pyodbc.connect('Driver={SQL Server};Server=dtc-stgsql1;Database=PLANKoin;Trusted_Connection=Yes;')            
            cursor = cnxn.cursor()            

            ICMSResults = cursor.execute("""
                
                DECLARE @return_value int
                EXEC @return_value = usp_ValidateICMS @ICMSNumber = ?;
                SELECT @return_value                
                
                """
                , matterNumber)
            self.matterResults = ICMSResults
        
            # assign all the case info to row variable
            row = self.matterResults.fetchone()
            
            if row: 
                self.leadada = row.leadada
                self.investigationnumber = row.investigationnumber
                self.investigationid = row.investigationid
                self.status = row.status
                self.casename = row.casename
                self.Bureau = row.Bureau
                self.ADA = row.ADA
                self.ADAUsername = row.ADAUsername
                self.email = row.email
                
                # clear status variable
                app.statusVariable.set("")
                app.status.configure(fg="#f6731b")

                # matterNumber has been validated                
                self.validatedToggle = True
                
                # set some global variables
                self.gADA = row.ADA
                self.gADAUsername = row.ADAUsername
                self.gBureau = row.Bureau
                self.gMatterNumber = matterNumber

            # gracefully handle no SQL results for verification
            else:
                print "No SQL data"
                app.statusVariable.set("That matter number does not exist, please try again")
                app.status.configure(fg="Red", font="Constantia 12 bold")
                
                self.leadada = "N/A"
                self.investigationnumber = "N/A"
                self.investigationid = "N/A"
                self.status = "N/A"
                self.casename = "N/A"
                self.Bureau = "N/A"
                self.ADA = "N/A"
                self.ADAUsername = "N/A"
                self.email = "N/A"
                
                self.validatedToggle = False
            
            # set labels for all the case info
            app.adaVariable.set("ADA: " + self.ADA)
            app.emailVariable.set("Email: " + self.email)
            app.bureauVariable.set("Bureau: " + self.Bureau)
            app.caseNameVariable.set("Case Name: " + self.casename)


        #####  FNUMBER  #####
        elif self.isFNumber:
            print "Its a TD investigation number (aka F-number)"
            matterNumber2 = matterNumber[(((matterNumber.find('-'))*(-1))-1):]
            fNumberResults = cursor.execute("""
                
                DECLARE @return_value int
                EXEC @return_value = usp_ValidateTDI @INV_ID = ?;
                SELECT @return_value                
                """
                , matterNumber2)
            self.matterResults = fNumberResults

            # assign all the case info to row variable
            row = self.matterResults.fetchone()
        
            if row:
                self.InvestigationName = row.InvestigationName
                self.InvestigationNumber = row.InvestigationNumber
                self.InvestigationStart = row.InvestigationStart
                self.InvestigationClosed = row.InvestigationClosed
                self.ADA = row.ADA
                self.ADAUsername = row.ADAUsername
                self.email = row.email
                self.BureauFullName = row.BureauFullName
                self.Bureau = row.Bureau

                # clear status variable
                app.statusVariable.set("")
                app.status.configure(fg="#f6731b")
                
                # matterNumber has been validated                
                self.validatedToggle = True
                
                # set some global variables
                self.gADA = row.ADA
                self.gADAUsername = row.ADAUsername
                self.gBureau = row.BureauFullName
                self.gMatterNumber = matterNumber

            # gracefully handle no SQL results for verification            
            else:
                print "No SQL data"
                app.statusVariable.set("That matter number does not exist, please try again")    
                app.status.configure(fg="Red", font="Constantia 12 bold")
                
                self.InvestigationName = "N/A"
                self.InvestigationNumber = "N/A"
                self.InvestigationStart = "N/A"
                self.InvestigationClosed = "N/A"
                self.ADA = "N/A"
                self.ADAUsername = "N/A"
                self.email = "N/A"
                self.BureauFullName = "N/A"
                self.Bureau = "N/A"
                
                self.validatedToggle = False

            # set labels for all the case info
            app.adaVariable.set("ADA: " + self.ADA)
            app.emailVariable.set("Email: " + self.email)
            app.bureauVariable.set("Bureau: " + self.BureauFullName)
            app.caseNameVariable.set("Case Name: " + self.InvestigationName)
        
        elif not matterNumber:
                app.statusVariable.set("The matter number field is blank")
                app.status.configure(fg="Red", font="Constantia 12 bold")
                print "matterNumber is blank"
                
                # clear from Fnumber
                self.InvestigationName = ""
                self.InvestigationNumber = ""
                self.InvestigationStart = ""
                self.InvestigationClosed = ""
                self.ADA = ""
                self.ADAUsername = ""
                self.email = ""
                self.BureauFullName = ""
                self.Bureau = ""
                
                # clear from ICMS
                self.leadada = ""
                self.investigationnumber = ""
                self.investigationid = ""
                self.status = ""
                self.casename = ""
                
                # clear from docket, arrestid, and indictment
                self.docket = ""        
                self.latestindictment = ""
                self.arrestID = ""
                self.defendant = ""
                
                app.adaVariable.set("Docket, Indictment, ArrestID, ICMS, TDI (F-number)")  
                app.emailVariable.set("")
                app.bureauVariable.set("")
                app.caseNameVariable.set("")
                
                self.validatedToggle = False
            
        
        else:
            app.statusVariable.set("Unrecognized DANY matter number, please try again" )
            app.status.configure(fg="Red", font="Constantia 12 bold")
            
            # set labels for all the case info
            app.adaVariable.set("Docket, Indictment, ArrestID, ICMS, TDI (F-number)")  
            app.emailVariable.set("")
            app.bureauVariable.set("")
            app.caseNameVariable.set("")
            
            self.validatedToggle = False


# import entry point 
importTool = ImportTool()

# interface entry point
app = interface(None)
app.title('DANY InPho (v.1.0)')
app.iconbitmap(r'\\DANY.NYCNET\DANYXDrive\Rikers Calls\InPho\InPho Icon.ico')

app.mainloop()
