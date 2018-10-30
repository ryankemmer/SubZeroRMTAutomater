VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Location_Change_Form 
   Caption         =   "Location Change"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8835
   OleObjectBlob   =   "Location_Change_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Location_Change_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

'Initialize text boxes

serial_number_box.Value = ""

'Initialize check boxes

storage2.Value = True

End Sub

Private Sub cancel_Click()

Unload Me

End Sub


Private Sub create_log_click()

Serial_Location_Change = serial_number_box.Value

'define integer for new location

If lab2 = True Then
    New_Location = 1
End If

If storage2 = True Then
    New_Location = 2
End If

If soundbooth2 = True Then
    New_Location = 3
End If

If harmony2 = True Then
    New_Location = 4
End If

If test2 = True Then
    New_Location = 5
End If

MsgBox Location2

Call Location_Change_Class
Call Location_Change_Unit_List

Unload Me





End Sub

