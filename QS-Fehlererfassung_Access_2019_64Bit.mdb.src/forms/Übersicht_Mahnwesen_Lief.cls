Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Schliessen_Click()
On Error GoTo Err_Schliessen_Click


    DoCmd.Close

Exit_Schliessen_Click:
    Exit Sub

Err_Schliessen_Click:
    MsgBox Err.Description
    Resume Exit_Schliessen_Click

End Sub
