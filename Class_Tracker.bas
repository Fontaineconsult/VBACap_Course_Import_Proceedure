Attribute VB_Name = "Class Tracker"



Option Compare Database
Function Prepare_Excel()


'Opens original excel doc and deletes first two rows so access can import and use first row as column headers

Dim filePath As String
Dim xl As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet


SysCmd acSysCmdInitMeter, "Formating Excel Document", 1
    
    filePath = "Z:\Databases\Courses\Courses.xls"
    
    Set xl = New Excel.Application

    Set xlBook = xl.Workbooks.Open(filePath)

    Set xlSheet = xlBook.Worksheets(1)

    xl.Visible = False
'''''''''''''''''''''''''''''
Set rCell = Range("C1")

If IsEmpty(rCell) = True Then
xlSheet.Rows("1:2").Delete
xlBook.Save
xlBook.Close

Else
    If MsgBox("Courses.XLS file is already formated. Click OK to continue", vbOKCancel, "Already Formated") = vbOK Then
    xlBook.Close
'''''''''''NEEDS EXIT CODE''''''''''''''''
    End If
End If


''''''''''''''''''''''''''''


    
Set xlApp = Nothing
Set xlSheet = Nothing
SysCmd acSysCmdRemoveMeter
End Function


Function Import_Excel()
Dim Location As String
Location = "Z:\Databases\Courses\Courses.xls"


SysCmd acSysCmdInitMeter, "Importing Excel Data into Access", 1



DoCmd.RunSavedImportExport ("Import_Bulk_Courses")

'DoCmd.TransferSpreadsheetacImport , acSpreadsheetTypeExcel9, "sheet1", "C:\Users\913678186\Desktop\Database Practice\SFO_CS_SR_DP_CLASSES_1193299.xls", True

SysCmd acSysCmdRemoveMeter

End Function



'After Import of bulk data, creates a copy of the original table with structure only (Sheet2)
Function Create_Sheet_Two()

SysCmd acSysCmdInitMeter, "Building Sheet 2", 1


Dim strPath As String
strPath = CurrentProject.FullName
DoCmd.TransferDatabase acImport, "Microsoft Access", strPath, acTable, "sheet1", "sheet2", True
DoCmd.RunSQL ("ALTER TABLE sheet2 ADD COLUMN Gen_Key TEXT(25);")
DoCmd.RunSQL ("ALTER TABLE sheet2 ADD COLUMN Semester TEXT(25);")


SysCmd acSysCmdRemoveMeter
End Function

'searches sheet1 and generates fake Instructor IDs where missing

Function Generate_Missing_IID()

Dim rst As DAO.Recordset
Set rst = CurrentDb.OpenRecordset("sheet1")
Dim IDString As String
Dim GeneratedID As String
Dim GeneratedIDpart As String
Dim NumArray(0 To 7) As String
 
 

Dim Count As Long
Dim Progress_Amount As Integer


rst.MoveLast
Count = rst.RecordCount
SysCmd acSysCmdInitMeter, "Generating Missing Instructor ID", Count
 
 
 
 
rst.MoveFirst
Do Until rst.EOF = True

n = n + 1
SysCmd acSysCmdUpdateMeter, n


    If Len(rst!Instructor_ID) < 9 Then
        If Len(rst!Email) > 0 Then
            IDString = rst!Email
        Else
            IDString = rst!Instructor_Name
        End If
            
        For x = 1 To 8
           
            Inputindvstring = Mid(IDString, x, 1)
            stringnum = Asc(Inputindvstring)
            stringnumcut = Right(stringnum, 1)
            NumArray(x - 1) = stringnumcut
         
        
        Next x
        Final8num = Join(NumArray, "")
        finalnum = "8" & Final8num
        rst.Edit
        rst!Instructor_ID = finalnum
        rst.Update
    End If
    rst.MoveNext
Loop
SysCmd acSysCmdRemoveMeter

End Function



'Adds to sheet2 all the students that we actually need to track

Function Find_Students()


Dim Students As DAO.Recordset
Dim rst As DAO.Recordset
Dim rst2 As DAO.Recordset
Set rst = CurrentDb.OpenRecordset("Sheet1")
Set Students = CurrentDb.OpenRecordset("Captioning Students")
Set rst2 = CurrentDb.OpenRecordset("Sheet2")


Students.MoveLast
Count = Students.RecordCount
SysCmd acSysCmdInitMeter, "Transfering Student's Courses", Count
 
Students.MoveFirst
Do Until Students.EOF = True

    n = n + 1
    SysCmd acSysCmdUpdateMeter, n
    
    StudentID = Students!ID
    DoCmd.RunSQL ("INSERT INTO sheet2 (ID, Student_ID, Student_Name, Student_Email, Descr, Course, Section, Instructor_Name, Phone, Email, Facil_ID, Mtg_Start, Mtg_End, Meetings, Diagnosis, Instructor_ID) SELECT * FROM sheet1 WHERE Student_ID = " & "'" & StudentID & "'" & ";")
    Students.MoveNext
Loop
    
SysCmd acSysCmdRemoveMeter
    
End Function


' generates a derived course key based on the semester, courses name, course section and inserts it into the "GEN_KEY" Column
Function Add_Key(Semester_Input)

Dim Students As DAO.Recordset
Dim rst As DAO.Recordset
Dim rst2 As DAO.Recordset
Set rst = CurrentDb.OpenRecordset("Sheet1")
Set Students = CurrentDb.OpenRecordset("Captioning Students")
Set rst2 = CurrentDb.OpenRecordset("Sheet2")

Dim Count As Long
Dim Progress_Amount As Integer


rst2.MoveLast
Count = rst2.RecordCount
SysCmd acSysCmdInitMeter, "Apply GenKey", Count



rst2.MoveFirst
Do Until rst2.EOF

n = n + 1
SysCmd acSysCmdUpdateMeter, n

COURSE = Replace(rst2!COURSE, " ", "")
SECTION = rst2!SECTION

Semester_Left = Left(Semester_Input, 2)
Semester_Right = Right(Semester_Input, 2)

Semester = Semester_Left & Semester_Right

Gen_Key = Semester & COURSE & SECTION

rst2.Edit
rst2!Gen_Key = Gen_Key
rst2!Semester = Semester_Input
rst2.Update
rst2.MoveNext

Loop
SysCmd acSysCmdRemoveMeter




End Function

'Moves Instructor info from sheet2 into the instructors table. Only takes records that are not already in the instructors table
Function Move_Instructors()

SysCmd acSysCmdInitMeter, "Building Instructors Table", 1


DoCmd.RunSQL ("INSERT INTO Instructors (Instructor_First_Name, Instructor_Last_Name, Instructor_Phone, Instructor_Email, Instructor_ID ) SELECT DISTINCT Instructor_Name, Instructor_Name, Phone, Email, Instructor_ID FROM Sheet2 WHERE Instructor_ID NOT IN (SELECT Instructor_ID FROM Instructors);")

Dim Insttable As Recordset
Set Insttable = CurrentDb.OpenRecordset("Instructors")
Dim name_array() As String
Dim Lastname As String
Dim Firstname As String

SysCmd acSysCmdRemoveMeter


Insttable.MoveLast
Count = Insttable.RecordCount
SysCmd acSysCmdInitMeter, "Apply GenKey", Count





Insttable.MoveFirst
' reformats the instructor string into two parts and removes the comma
Do Until Insttable.EOF = True

n = n + 1
SysCmd acSysCmdUpdateMeter, n


If Insttable!Instructor_First_Name = Insttable!Instructor_Last_Name Then

Inst_Name = Insttable!Instructor_First_Name
name_array() = Split(Inst_Name, ",")

Firstname = name_array(1)
Lastname = name_array(0)

Insttable.Edit
Insttable!Instructor_First_Name = Firstname
Insttable!Instructor_Last_Name = Lastname
Insttable.Update

End If

Insttable.MoveNext

Loop

SysCmd acSysCmdRemoveMeter



End Function


Function Insert_Instructor_And_Gen_Key()

'Inserts the Instructor and Gen_Key into the courses table creating a unique course object

Dim StudentTbl As Recordset
Set StudentTbl = CurrentDb.OpenRecordset("Captioning Students")
Dim Courses As Recordset
Set Courses = CurrentDb.OpenRecordset("Courses")
Dim sht2 As Recordset
Set sht2 = CurrentDb.OpenRecordset("Sheet2")
Dim locarray(0 To 1) As String
Dim DateTime As Date
DateTime = Date


StudentTbl.MoveLast
Count = StudentTbl.RecordCount
SysCmd acSysCmdInitMeter, "Constructing Courses", Count



StudentTbl.MoveFirst
Do Until StudentTbl.EOF = True

n = n + 1
SysCmd acSysCmdUpdateMeter, n
    
' Inserts the course data
    If StudentTbl!Approved = True Then
        StudentID = StudentTbl!ID
        DoCmd.RunSQL ("INSERT INTO Courses (Gen_Key, Course, Section, Instructor_ID, Semester) SELECT DISTINCT Gen_Key, Course, Section, Instructor_ID, Semester FROM Sheet2 WHERE STUDENT_ID = " & "'" & StudentID & "'" & " AND Gen_Key NOT IN (SELECT Gen_Key FROM Courses) ;")
    End If
        StudentTbl.MoveNext
Loop

' if there are multple locations to a class, they will be combined into one cell
Courses.MoveFirst
Do Until Courses.EOF = True
  

    Instructor_ID = Courses!Instructor_ID
    Gen_Key = Courses!Gen_Key
    
    sht2.MoveFirst
    
        Do Until sht2.EOF = True
            If sht2!Gen_Key = Gen_Key And sht2!Instructor_ID = Instructor_ID Then
                If locarray(0) = "" Then
                    locarray(0) = sht2!Facil_ID
                Else
                    locarray(1) = sht2!Facil_ID
                End If
                
                Courses.Edit
                
                If locarray(0) = "" And locarray(1) = "" Then
                    Courses!Location = ""
                ElseIf locarray(1) = "" Then
                    Courses!Location = locarray(0)
                Else
                Courses!Location = locarray(0) + ", " + locarray(1)
                End If
                
                
                If locarray(0) = "ONLINE" Or locarray(1) = "ONLINE" Then
                Courses!Online = True
                End If
                
                Courses.Update
                
            End If
            sht2.MoveNext
        Loop
    Courses.MoveNext
    locarray(0) = ""
    locarray(1) = ""
Loop

' Adds import date

Courses.MoveFirst
Do Until Courses.EOF
    

        If IsNull(Courses!Import_Date) = True Then
            
        Courses.Edit
        Courses!Import_Date = Date
        Courses.Update
            
        Else
        End If
                 
   
Courses.MoveNext
Loop



SysCmd acSysCmdRemoveMeter
Courses.MoveFirst
sht2.MoveFirst
End Function

' Finds, compares, adds the student ID to the correct courses.
Function Add_Students()

Dim StudentTbl As Recordset
Set StudentTbl = CurrentDb.OpenRecordset("Captioning Students")
Dim Courses As Recordset
Set Courses = CurrentDb.OpenRecordset("Courses")
Dim sht2 As Recordset
Set sht2 = CurrentDb.OpenRecordset("Sheet2")


StudentTbl.MoveLast
Count = StudentTbl.RecordCount
SysCmd acSysCmdInitMeter, "Assigning Students", Count



Do Until StudentTbl.EOF = True

n = n + 1
SysCmd acSysCmdUpdateMeter, n

    If StudentTbl!Approved = True Then
    
       Courses.MoveFirst
        Do Until Courses.EOF = True
        
            Gen_Key = Courses!Gen_Key
            Instructor_ID = Courses!Instructor_ID
                sht2.MoveFirst
                
                Do Until sht2.EOF = True
                
                    If sht2!Gen_Key = Gen_Key And sht2!Instructor_ID = Instructor_ID Then
                    
                        If IsNull(Courses!Student_ID) = True Then
                        
                        Courses.Edit
                        Courses!Student_ID = sht2!Student_ID
                        Courses.Update
                    
                        ElseIf IsNull(Courses!Student_ID) = False And Courses!Student_ID <> sht2!Student_ID Then
                        
                        Courses.Edit
                        Courses!Student_ID_2 = sht2!Student_ID
                        Courses.Update
                        
                        End If
                        
                      
                        
                    End If
                    sht2.MoveNext
                Loop
        Courses.MoveNext
        Loop
    End If
 StudentTbl.MoveNext
Loop
        
                    

SysCmd acSysCmdRemoveMeter

End Function

Function Check_Drops()


Dim Courses As Recordset
Set Courses = CurrentDb.OpenRecordset("Courses")
Dim sht2 As Recordset
Set sht2 = CurrentDb.OpenRecordset("Sheet2")


Courses.MoveLast
Count = Courses.RecordCount
SysCmd acSysCmdInitMeter, "Updating Enrollement", Count

Courses.MoveFirst
Do Until Courses.EOF

n = n + 1
SysCmd acSysCmdUpdateMeter, n


    Gen_Key = Courses!Gen_Key
    Instructor_ID = Courses!Instructor_ID
    Student_ID = Courses!Student_ID

    sht2.MoveFirst
    Do Until sht2.EOF
        Status = sht2!Status
        If Gen_Key = sht2!Gen_Key And Instructor_ID = sht2!Instructor_ID And Student_ID = sht2!Student_ID Then
            Courses.Edit
            If Status = "Dropped" Then
                Courses!No_Students_Enrolled = True
            Else
                Courses!No_Students_Enrolled = False
            End If
            Courses.Update
        End If
    sht2.MoveNext
    Loop
Courses.MoveNext
Loop
SysCmd acSysCmdRemoveMeter

End Function

Function Delete_Bulk_Tables(Delete_Bulk)

SysCmd acSysCmdInitMeter, "Cleaning Up", 1

If Delete_Bulk = True Then

DoCmd.DeleteObject acTable, "sheet1"
DoCmd.DeleteObject acTable, "sheet2"

Else
End If

SysCmd acSysCmdRemoveMeter


DoCmd.OpenForm ("Enrolled_Querry")

End Function

