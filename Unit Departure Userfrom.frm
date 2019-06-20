VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UnitDeparture 
   Caption         =   "Unit Departure"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   10540
   OleObjectBlob   =   "Unit Departure Userfrom.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UnitDeparture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub create_log_click()

Serial_Departure = serial_number_box.Value
Date_Departure = mm.Value & "/" & dd.Value & "/" & yyyy.Value
Root_Cause = root_cause_box.Value

Unload Me

Call Unit_Departure_Unit_List

End Sub


Private Sub cancel_Click()

Unload Me

End Sub

Private Sub UserForm_Initialize()

'Initialize text boxes

serial_number_box.Value = ""
root_cause_box.Value = ""

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

Private Sub yyyy_Change()

End Sub

