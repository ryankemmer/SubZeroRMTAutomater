Attribute VB_Name = "Change_Location"
'*****************************************************
'Change_Location
'
'Purpose: Subs dedicated to changing the location of units by updating the unit tracking list.
'Author: Ryan Kemmer
'Last Updated: 10/30/2018
'*****************************************************

Public Serial_Location_Change As String
Public Current_Location As Integer
Public New_Location_String As String
Public New_Location As Integer
Public Current_Location_Folder As String
Public New_Location_Folder As String

Sub Show_Location_Change()

Location_Change_Form.Show

End Sub

Sub Location_Change_Unit_List()

Dim tracklist As ListObject
Dim ws As Worksheet
Dim FirstRow As Long
Dim LastRow As Long
Dim Lrow As Long
Dim row As Long

'Activate Sheet
Workbooks("Solution Log - Template.xlsm").Activate
Worksheets("Unit List").Activate

'Find Serial in unit list
Set ws = ActiveSheet
Set tracklist = ws.ListObjects("Unit_List")

LastRow = tracklist.DataBodyRange.Rows.Count
FirstRow = 2

For Lrow = LastRow To FirstRow Step -1
    With tracklist.DataBodyRange.Cells(Lrow, 4)
        If .Value = Serial_Location_Change Then
            row = Lrow
            Exit For
        End If
    End With
Next Lrow

'Error Case for serial does not exist

With tracklist
    .DataBodyRange.Cells(row, 8).clear
    .DataBodyRange.Cells(row, 8) = New_Location_String
    .DataBodyRange.Cells(row, 13).Select
End With

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
sDestFolder = Solution_Logs_Folder & "\" & New_Location_Folder & "\" & FolderName

FSO.MoveFolder sFolder, sDestFolder

'Re Activate Tracking list
Workbooks("Solution Log - Template.xlsm").Activate
Worksheets("Unit List").Activate

'recreate hyperlink
tracklist.DataBodyRange.Cells(row, 13).Select

ws.Hyperlinks.Add Anchor:=Selection, _
        Address:=sDestFolder & "\" & FileName, TextToDisplay:="Link"
        
End Sub

Sub Location_Change_Class()

Select Case New_Location
    Case 1
        New_Location_String = "Lab"
        New_Location_Folder = "In Lab"
    Case 2
        New_Location_String = "Storage"
        New_Location_Folder = "Storage"
    Case 3
        New_Location_String = "Harmony Room"
        New_Location_Folder = "In Lab"
    Case 4
        New_Location_String = "CAL"
        New_Location_Folder = "In Lab"
    Case 5
        New_Location_String = "TE"
        New_Location_Folder = "In Lab"
End Select

End Sub
