Attribute VB_Name = "CreateRMTSheet"
'*****************************************************
'CreateRMTSheet
'
'Purpose: Subs dedicated to creating a new RMT sheet. Subs will create a new sheet with populated information, and create a
'folder dedicated to the RMT unit. The subs also will update the tracking list and tracking list.
'Author: Ryan Kemmer
'Last Updated: 10/30/2018
'*****************************************************

Public Model As String
Public Serial As String
Public RMTNumber As String
Public Loc As String
Public Service As String
Public Description As String
Public ServiceP As String
Public Additional As String
Public Date_Requested As String

Sub Show_RMTSheet()

RMTSheet.Show

End Sub
Sub Create_Solution_Log()

Application.ScreenUpdating = False

Dim Answer As Integer
Dim tracklist As ListObject
Dim pending As ListObject
Dim ws As Worksheet
Dim FolderN As String
Dim FileN As String

'Set Filename
FolderN = Serial & " " & Model & " - " & Description
FileN = Serial & " " & Model & " - " & Description & ".xlsx"

'Create Folder
Call Create_Folder(FolderN)

'Activate Solutions Log
Worksheets("Solutions Log").Activate

'Populate Sheet
With ActiveSheet
    .Range("E3").Value = Model
    .Range("C3").Value = Serial
    .Range("C4").Value = RMTNumber
    .Range("C5").Value = Service
    .Range("G5").Value = Loc
    .Range("C6").Value = Description
    .Range("C7").Value = ServiceP
End With

If Not Additional = "" Then
    With ActiveSheet.Range("C17")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Value = Additional
    End With
End If

'Copy workbook and save new workbook in pending arrival section
Sheets("Solutions Log").Copy
ActiveWorkbook.SaveAs FileName:="P:\Teamwork\Reliability\Reliability Files\Lab Units\Solution Logs\Pending Arrival\" & _
FolderN & "\" & FileN, FileFormat:=xlOpenXMLWorkbook

'Prompt User to Update Tracking List
Answer = MsgBox("Would you like to update the unit tracking list?", vbQuestion + vbYesNo)
    If Answer = vbYes Then
    
        'Close old workbook
        ActiveWorkbook.Close Savechanges:=True
        Workbooks("Solution Log - Template.xlsm").Activate
        
        'Update tracking list
        Call Update_Tracking_List
        Call Clear_Solutions_Log
        
    Else
        cancel = True
        'Close New Workbook
        ActiveWorkbook.Close
        Workbooks("Solution Log - Template.xlsm").Activate
        Call Clear_Solutions_Log
            
    End If

End Sub

Function Clear_Solutions_Log()

Worksheets("Solutions Log").Activate
With ActiveSheet
    .Range("E3").Value = ""
    .Range("C3").Value = ""
    .Range("C4").Value = ""
    .Range("C5").Value = ""
    .Range("G5").Value = ""
    .Range("C6").Value = ""
    .Range("C7").Value = ""
    .Range("C17").Value = ""
End With

End Function

Sub Update_Tracking_List()

Application.ScreenUpdating = False
Dim tracklist As ListObject
Dim pending As ListObject
Dim ws As Worksheet
Dim FolderN As String
Dim FileN As String

'Set Filename
FolderN = Serial & " " & Model & " - " & Description
FileN = Serial & " " & Model & " - " & Description & ".xlsx"

Worksheets("Unit List").Activate

'Add new row to table
Set ws = ActiveSheet
Set tracklist = ws.ListObjects("Unit_List")
Dim newrow As ListRow
Range("B" & Rows.Count).End(xlUp).Select
Set newrow = tracklist.ListRows.Add

'Add in information
With newrow
    .Range(1) = Date_Requested
    .Range(4) = Serial
    .Range(5) = Model
    .Range(6) = "RMT"
    .Range(7) = "Pending"
    .Range(8) = "Pending"
    .Range(9) = Description
    .Range(13).Select
End With

ws.Hyperlinks.Add Anchor:=Selection, _
        Address:="P:\Teamwork\Reliability\Reliability Files\Lab Units\Solution Logs\Pending Arrival\" & _
        FolderN & "\" & FileN, _
        TextToDisplay:="Link"

'Save changes
ActiveWorkbook.Save

End Sub

Sub Create_Folder(F As String)

MkDir "P:\Teamwork\Reliability\Reliability Files\Lab Units\Solution Logs\Pending Arrival\" & F

End Sub


