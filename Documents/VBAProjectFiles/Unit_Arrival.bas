Attribute VB_Name = "Unit_Arrival"
'*****************************************************
'Unit_Arrival
'
'Purpose: Subs dedicated to automating the unit arrival process. Subs will update the tracking list, update the RMT sheet document, and update folder locations. \
'The subs also allow for the user to send an arrival email.
'Author: Ryan Kemmer
'Last Updated: 10/30/2018
'*****************************************************

Public serial_arrival As String
Public Date_Arrival As String
Public Location_Arrival As Integer
Public Location_String As String
Public Location_Folder_String As String
Public Cell_Contents As String
Public Model_arrival As String

Sub Show_Unit_Arrival()

UnitArrival.Show

End Sub


Sub Unit_Arrival_UnitList()

Dim tracklist As ListObject
Static ws As Worksheet
Dim FirstRow As Long
Dim LastRow As Long
Dim Lrow As Long
Dim row As Long
Dim Description_arrival As String

'Activate Sheet
Workbooks("Solution Log - Template.xlsm").Activate
Worksheets("Unit List").Activate

'Find Serial in unit list
Set ws = ActiveSheet
Set tracklist = ws.ListObjects("Unit_List")

LastRow = tracklist.DataBodyRange.Rows.Count
FirstRow = 2

On Error GoTo Serial_Dosent_Exist

For Lrow = LastRow To FirstRow Step -1
    With tracklist.DataBodyRange.Cells(Lrow, 4)
        If .Value = serial_arrival Then
            row = Lrow
            Exit For
        End If
    End With
Next Lrow

'Error Case for serial does not exist
If row = 0 Then
    GoTo Serial_Dosent_Exist
    Exit Sub
End If

'Update Values in Table
'fetch model number
'Select Cell for hyperlink
Status = "In Progress"

With tracklist
    .DataBodyRange.Cells(row, 2) = Date_Arrival
    .DataBodyRange.Cells(row, 7) = Status
    .DataBodyRange.Cells(row, 8) = Location_String
    .DataBodyRange.Cells(row, 11) = "Yes"
    Model_arrival = .DataBodyRange.Cells(row, 5).Value
    Description_arrival = .DataBodyRange.Cells(row, 9).Value
    .DataBodyRange.Cells(row, 13).Select
End With

'test to see if link is active

If (ActiveCell.Value = "") Then
    Exit Sub
End If

'Follow Link and update with arrival date
On Error GoTo HyperLinkDead

ActiveCell.Hyperlinks(1).Follow

ActiveSheet.Range("G3").Value = Date_Arrival

'Move Folder

On Error GoTo FolderError

Dim FSO As Object
Dim sFolder As String
Dim sDestFolder As String
Dim Path As String
Dim FolderName As String
Dim FileName As String
Dim up1 As String
Dim Solution_Logs_Folder As String

Set FSO = CreateObject("scripting.filesystemobject")
Path = Application.ActiveWorkbook.Path

With FSO
    up1 = .GetParentFolderName(Path)
    Solution_Logs_Folder = .GetParentFolderName(up1)
End With

FileName = FSO.GetFileName(Path) & ".xlsx"
FolderName = Mid(Path, InStrRev(Path, "\") + 1)

ActiveWorkbook.Save
ActiveWorkbook.Close

sFolder = Path
sDestFolder = Solution_Logs_Folder & "\" & Location_Folder_String & "\" & FolderName

FSO.MoveFolder sFolder, sDestFolder
    
'Re Activate Tracking list
Workbooks("Solution Log - Template.xlsm").Activate
Worksheets("Unit List").Activate

'recreate hyperlink
tracklist.DataBodyRange.Cells(row, 13).Select

ws.Hyperlinks.Add Anchor:=Selection, _
        Address:=sDestFolder & "\" & FileName, _
        TextToDisplay:="Link"
        
'Error Cases
Done:
    Exit Sub
Serial_Dosent_Exist:
    MsgBox "Error: Serial does not exist"
    End
HyperLinkDead:
    MsgBox "Hyperlink is dead. Please change location."
FolderError:
    MsgBox "Folder Error"
    End
    
End Sub


Sub Location_Class()

Select Case Location_Arrival
    Case 1
        Location_String = "Lab"
        Location_Folder_String = "In Lab"
    Case 2
        Location_String = "Storage"
        Location_Folder_String = "Storage"
    Case 3
        Location_String = "Harmony Room"
        Location_Folder_String = "In Lab"
    Case 4
        Location_String = "CAL"
        Location_Folder_String = "In Lab"
    Case 5
        Location_String = "TE"
        Location_Folder_String = "In Lab"
End Select

End Sub


Sub Send_Arrival_Email()

Dim OutApp As Object
Dim Outmail As Object
Dim htmlbody As String

Set OutApp = CreateObject("Outlook.Application")
Set Outmail = OutApp.CreateItem(0)

hbody = "<BODY style=font-size:11pt;font-family:Calibri>Hi All, <br><br>" & _
        "Goodyear Reliability has recieved the following units: <br><br>" & _
        Model_arrival & " (" & serial_arrival & ") <br><br>" & _
        "Best,</BODY>"

On Error Resume Next
With Outmail
    .Display
    .To = "goodyearrmt@subzero.com"
    .CC = "bob.zoladz@subzero.com"
    .Bcc = ""
    .Subject = "RMT Unit Arrived"
    .htmlbody = hbody + .htmlbody
    
End With
On Error GoTo 0

Set Outmail = Nothing
Set OutApp = Nothing

End Sub

