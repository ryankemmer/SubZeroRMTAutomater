VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UnitArrival 
   Caption         =   "Unit Arrival"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   8520.001
   OleObjectBlob   =   "Unit Arrival Userform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UnitArrival"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cancel_Click()

Unload Me

End Sub

Private Sub create_log_click()

serial_arrival = serial_number_box.Value
Date_Arrival = mm & "/" & dd & "/" & yyyy

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

Unload Me

Call Location_Class

Call Unit_Arrival_UnitList


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

