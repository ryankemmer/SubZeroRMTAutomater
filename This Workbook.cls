VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, cancel As Boolean)

Worksheets("Unit List").Activate
Range("R3").Value = Application.UserName
Range("R2").Value = FormatDateTime(Now, vbGeneralDate)
Worksheets("Lab Layout").Activate
Range("H2").Value = Application.UserName
Range("H1").Value = FormatDateTime(Now, vbGeneralDate)


End Sub


Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

If Update = True Then
    Exit Sub
End If

Dim KeyCells As Range
Set ws = ActiveSheet
Dim tbl As ListObject
If ActiveSheet.Name = "Unit List" Then
    Set tbl = ws.ListObjects("Unit_List")
    Set KeyCells = tbl.ListColumns(8).Range
Else
    Exit Sub
End If

If Not Application.Intersect(KeyCells, Range(Target.Address)) Is Nothing Then

    Dim row As Long
    row = Target.row - 1

    If tbl.DataBodyRange(row, 13).Hyperlinks.Count = 0 Then
        Exit Sub
    End If
    
    If MsgBox("You are about to update a unit location link, proceed?", vbYesNo) = vbYes Then
        
        If InStr(1, Target.Text, "Storage") > 0 Then
    
            tbl.DataBodyRange(row, 13).Select
            ActiveCell.Hyperlinks(1).Follow
            Call Change_Folder("Storage", row)
            MsgBox ("Folder updated to Storage")
        Else
    
            tbl.DataBodyRange(row, 13).Select
            ActiveCell.Hyperlinks(1).Follow
            Call Change_Folder("In Lab", row)
            MsgBox ("Folder updated to lab")
        
        End If
    
    Else
        Exit Sub
    End If
        
End If

End Sub
