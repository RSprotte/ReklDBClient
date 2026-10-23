Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim gs_AddressFeld As String

Private Sub cmbKontaktpersonUebernehmen_Click()

        Dim lForm As Object
        Dim i As Integer
        i = 0

        For Each lForm In Forms
            If lForm.Name = "MailSendenAnNeu" Then Exit For
            i = i + 1
        Next
        If IsNull(Forms(i).Controls(gs_AddressFeld)) Then
            Forms(i).Controls(gs_AddressFeld) = Me.cobNameKontaktperson.Column(0)
            DoCmd.Close
        Else
            Forms(i).Controls(gs_AddressFeld) = Forms(i).Controls(gs_AddressFeld) & "; " & Me.cobNameKontaktperson.Column(0)
        End If

        DoCmd.Close acForm, "frmEmailAdresse"

        Forms(i).Controls(gs_AddressFeld).SetFocus

End Sub

Private Sub Form_Open(Cancel As Integer)
'    Dim varKoord As Variant
'    varKoord = getFktKoordinatenRelHF(Me)
'    If varKoord(3) > 0 And varKoord(2) > 0 Then
'        'Fenster setzen relativ zum Hauptfenster TB:autostart
'        DoCmd.MoveSize varKoord(0), varKoord(1), varKoord(2), varKoord(3)
'    End If
'    DoCmd.Restore
    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long

    lgPosition = getPointerPos(XPos, YPos)
    FormOpenKoorMousPointer Me, XPos, YPos, 500, 150

    gs_AddressFeld = Me.OpenArgs

End Sub

Private Sub Form_Load()

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
