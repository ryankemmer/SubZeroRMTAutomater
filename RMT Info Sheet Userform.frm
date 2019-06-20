VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RMTSheet 
   Caption         =   "Official RMT WebSheet"
   ClientHeight    =   10875
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   9960.001
   OleObjectBlob   =   "RMT Info Sheet Userform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RMTSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cancel_Click()

Unload Me

End Sub

Private Sub clear_Click()

Call UserForm_Initialize

End Sub


Private Sub create_log_click()

Serial = serial_number_box.Value
Model = model_box.Value
RMTNumber = rmtnumber_box.Value
Loc = location_box.Value
Service = service_box.Value
Description = reason_return.Value
ServiceP = service_performed.Value
Additional = additional_info.Value
Date_Requested = mm.Value & "/" & dd.Value & "/" & yyyy.Value

Unload Me

Call Create_Solution_Log

End Sub


Private Sub UserForm_Initialize()

'Initialize text boxes

serial_number_box.Value = ""
model_box.Value = ""
rmtnumber_box.Value = ""
location_box.Value = ""
service_box.Value = ""
reason_return.Value = ""
service_performed.Value = ""
additional_info.Value = ""

'Initialize date boxes
With mm
    Dim i As Integer
        For i = 1 To 12
            .AddItem (i)
        Next i
End With

With dd
    Dim j As Integer
        For j = 1 To 31
            .AddItem (j)
        Next j
End With

With yyyy
    Dim d As Integer
    d = Year(Now)
    .AddItem (d)
    .AddItem (d + 1)
End With

End Sub
