VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UnitArrival 
   Caption         =   "Unit Arrival"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8520
   OleObjectBlob   =   "UnitArrival.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UnitArrival"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub create_log_click()

Dim Answer As Integer

serial_arrival = serial_number_box.Value
Date_Arrival = mm.Value & "/" & dd.Value & "/" & yyyy.Value

If lab = True Then
    Location_Arrival = 1
End If

If storage = True Then
    Location_Arrival = 2
End If

If soundbooth = True Then
    Location_Arrival = 3
End If

If harmony = True Then
    Location_Arrival = 4
End If

If test = True Then
    Location_Arrival = 5
End If

Call Unit_Arrival.Location_Class
Call Unit_Arrival.Unit_Arrival_UnitList

Answer = MsgBox("Would you like to send out RMT arrival email?", vbYesNo + vbQuestion, "cancel")
    If Answer = vbYes Then
        Call Send_Arrival_Email
    Else
        cancel = True
    End If

Unload Me

End Sub

Private Sub Label1_Click()

End Sub
Private Sub serial_number_box_Change()

End Sub

Private Sub storage_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub cancel_Click()

Unload Me

End Sub

Public Sub UserForm_Initialize()

'Initialize text boxes

serial_number_box.Value = ""

'Initialize check boxes

lab.Value = True

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
