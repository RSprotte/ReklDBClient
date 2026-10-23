Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyDown KeyCode, Shift, Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyUp KeyCode, Shift
End Sub

Private Sub Form_Load()

    Dim objMouseOverForm As clsMouseoverForm
    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

End Sub

Private Sub Form_Open(Cancel As Integer)

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub

Private Sub FormularBearbeiterSchließen_Click()
On Error GoTo Err_FormularBearbeiterSchließen_Click

    FormKoorSave Me
    DoCmd.Close

Exit_FormularBearbeiterSchließen_Click:
    Exit Sub

Err_FormularBearbeiterSchließen_Click:
    MsgBox Err.Description
    Resume Exit_FormularBearbeiterSchließen_Click

End Sub
