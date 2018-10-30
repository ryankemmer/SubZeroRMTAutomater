VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UnitDeparture 
   Caption         =   "Unit Departure"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10545
   OleObjectBlob   =   "UnitDeparture.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UnitDeparture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub create_log_click()

Dim Answer As Integer

Serial_Departure = serial_number_box.Value
Date_Departure = mm.Value & "/" & dd.Value & "/" & yyyy.Value
Root_Cause = root_cause_box.Value

Call Unit_Departure_Unit_List

Answer = MsgBox("Would you like to send out RMT departure email?" + vbYesNo + vbQuestion, "cancel")
    If Answer = vbYes Then
        Call Send_Departure_Email
    Else
        cancel = True
    End If

Unload Me

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

