Attribute VB_Name = "Change_Folder_Location"
'*****************************************************
'Change Folder Location
'
'Purpose: Public Sub created that moves a folder location from lab to storage, or storage to lab.
'The sub moves the dedicated solutions log to the updated folder location, and then updated the hyperlink
'This sub is called when a change is made to the location/status of a unit
'Author: Ryan Kemmer
'Last Updated: 5/15/2019

'*****************************************************

'Called after a hyperlink to a solution log is followed
Public Sub Change_Folder(Location_Folder As String, row As Long)

'Turn off screen updating
Application.ScreenUpdating = False

Dim FSO As Object
Dim sFolder As String
Dim sDestFolder As String
Dim Path As String
Dim FolderName As String
Dim FileName As String
Dim up1 As String
Dim Solution_Logs_Folder As String
Dim ws As Worksheet
Dim tracklist As ListObject

'Set Activesheet to variable (the active sheet should be a solution log)
Set ws = ActiveSheet

'Create file system object to work with the computer files
Set FSO = CreateObject("scripting.filesystemobject")

'Determine path of the active workbook
Path = Application.ActiveWorkbook.Path

'Get the name of the parent folder, and the parents parent folder of the solution log
With FSO
    up1 = .GetParentFolderName(Path)
    Solution_Logs_Folder = .GetParentFolderName(up1)
End With

'Get the name of the solution log file
FileName = FSO.GetFileName(Path) & ".xlsx"

'Get the folder name of the solution log file
FolderName = Mid(Path, InStrRev(Path, "\") + 1)

'Save and close the solution log
ActiveWorkbook.Save
ActiveWorkbook.Close

'Locate the current folder location
sFolder = Path

'Set the path of the new folder location
sDestFolder = Solution_Logs_Folder & "\" & Location_Folder & "\" & FolderName

'Move the folder from the current location to the new location
FSO.MoveFolder sFolder, sDestFolder
    
'Re Activate Tracking list
Workbooks("Unit Tracking List - Lab Layout .xlsm").Activate
Worksheets("Unit List").Activate

'recreate hyperlink

Set ws = ActiveSheet
Set tracklist = ws.ListObjects("Unit_List")

tracklist.DataBodyRange.Cells(row, 13).Select

ws.Hyperlinks.Add Anchor:=Selection, _
        Address:=sDestFolder & "\" & FileName, _
        TextToDisplay:="Link"

End Sub
