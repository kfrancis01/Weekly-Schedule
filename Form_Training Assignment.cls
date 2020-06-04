VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Training Assignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub New_Click()
DoCmd.OpenForm "New Training Task", , , , , acDialog
End Sub

Private Sub ShowAll_Click()
DoCmd.OpenForm "Training Assignment", , , , , acDialog
End Sub
