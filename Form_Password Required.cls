VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Password Required"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Private Sub Enter_Click()
'If Me.Password.Value = "Robo" Then
'
'End If
'End Sub
Private Sub Enter_Click()
If Me.Password.Value = "Robots" Then
DoCmd.Close acForm, "Password Required"
DoCmd.OpenForm "Schedule Update"
Else
MsgBox "The wrong password was attempted", vbInformation, "Incorrect Password"
End If
End Sub
