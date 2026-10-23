Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Befehl6_Click()
On Error GoTo Err_Befehl6_Click


    Screen.PreviousControl.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_Befehl6_Click:
    Exit Sub

Err_Befehl6_Click:
    MsgBox Err.Description
    Resume Exit_Befehl6_Click

End Sub
