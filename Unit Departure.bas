Attribute VB_Name = "Unit_Departure"
'*****************************************************
'Unit_Departure
'
'Purpose: Subs dedicated to automating the unit departure process. Subs will update the tracking list, update the RMT sheet document, and update folder locations.
'The subs also allow for the user to send a departure email.
'Author: Ryan Kemmer
'Last Updated: 5/15/2019
'*****************************************************

Public Serial_Departure As String
Public Date_Departure As String
Public Root_Cause As String
Public Model_Departure As String
Public Description_Departure As String

Sub Show_UnitDeparture()

UnitDeparture.Show

End Sub

Sub Unit_Departure_Unit_List()

Update = True

Dim tracklist As ListObject
Dim ws As Worksheet
Dim FirstRow As Long
Dim LastRow As Long
Dim Lrow As Long
Dim row As Long

Worksheets("Unit List").Activate

Set ws = ActiveSheet
Set tracklist = ws.ListObjects("Unit_List")

LastRow = tracklist.DataBodyRange.Rows.Count
FirstRow = 1

For Lrow = LastRow To FirstRow Step -1
    With tracklist.DataBodyRange.Cells(Lrow, 4)
        If .Value = Serial_Departure Then
            row = Lrow
            Exit For
        End If
    End With
Next Lrow

With tracklist
    .DataBodyRange.Cells(row, 3) = Date_Departure
    .DataBodyRange.Cells(row, 7) = "Completed"
    .DataBodyRange.Cells(row, 8) = "Scrapped"
    .DataBodyRange.Cells(row, 10) = Root_Cause
    .DataBodyRange.Cells(row, 12) = "Yes"
     Model_Departure = .DataBodyRange.Cells(row, 5).Value
    .DataBodyRange.Cells(row, 13).Select
End With

'Follow Link
If (ActiveCell.Value = "") Then
    End
End If

'Follow Link

ActiveCell.Hyperlinks(1).Follow

'Move Folder

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

ActiveWorkbook.Close Savechanges:=True

sFolder = Path
sDestFolder = Solution_Logs_Folder & "\Completed\Cummulative List of Reports\" & FolderName

FSO.MoveFolder sFolder, sDestFolder

'Re Activate Tracking list
Workbooks("Unit Tracking List - Lab Layout .xlsm").Activate
Worksheets("Unit List").Activate

'recreate hyperlink
tracklist.DataBodyRange.Cells(row, 13).Select

ws.Hyperlinks.Add Anchor:=Selection, _
        Address:=sDestFolder & "\" & FileName, TextToDisplay:="Link"

'open workbook
Workbooks.Open FileName:=sDestFolder & "\" & FileName

'prompt to send email
Answer = MsgBox("Would you like to send out RMT departure email?", vbYesNo + vbQuestion, "cancel")
    If Answer = vbYes Then
        Call Send_Departure_Email
    Else
        cancel = True
    End If

'Prompt to update log
MsgBox "Please Update Solutions Log"

Update = False

End Sub

Sub Send_Departure_Email()

Dim OutApp As Object
Dim Outmail As Object
Dim htmlbody As String

Set OutApp = CreateObject("Outlook.Application")
Set Outmail = OutApp.CreateItem(0)

hbody = "<BODY style=font-size:11pt;font-family:Calibri>Hi All, <br><br>" & _
        "Goodyear Reliability has scrapped the following units: <br><br>" & _
        Model_Departure & " (" & Serial_Departure & ") <br><br>" & _
        "Best,</BODY>"

On Error Resume Next
With Outmail
    .Display
    .To = "goodyearrmt@subzero.com"
    .CC = "bob.zoladz@subzero.com"
    .Bcc = ""
    .Subject = "RMT Unit Scrapped"
    .htmlbody = hbody + .htmlbody
    
End With
On Error GoTo 0

Set Outmail = Nothing
Set OutApp = Nothing

End Sub



