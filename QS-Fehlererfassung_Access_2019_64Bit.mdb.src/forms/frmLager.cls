Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmbLagerUebernehmen_Click()

    Reports!Abholauftrag!txtLager = Me.cobLager.Column(1)
    Reports!Abholauftrag!txtDatum = Me.txtAusfuehrenAm
    If IsNull(Me.cobLager.Column(1)) Then
'        MsgBox ("Bitte Lager auswählen!")
        MsgBox getTC_FCFromCode("TC3215"), vbInformation, getTC_FCFromCode("TC3216")
        Exit Sub
    End If
    DoCmd.Close

End Sub



Private Sub Form_Load()

    Me.txtAusfuehrenAm = format(Date, "dd/mm/yyyy")

    Dim objMouseOverForm As clsMouseoverForm
    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub

Private Sub Form_Open(Cancel As Integer)

    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long

    lgPosition = getPointerPos(XPos, YPos)
    FormOpenKoorMousPointer Me, XPos, YPos, 400, 200, False

End Sub

Private Sub Form_ViewChange(ByVal Reason As Long)

MsgBox "Reason: " & Reason

End Sub

Private Sub txtAusfuehrenAm_Click()

    Dim frm As Form

    DoCmd.OpenForm ("frmCalendar")
    Set frm = Forms("frmCalendar")
    Call frm.setParentForm(Me)

End Sub
Public Sub setValue(ByVal CntDate As Date)

    Me.txtAusfuehrenAm = CntDate

End Sub
