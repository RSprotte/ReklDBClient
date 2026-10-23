Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmbFrage_Click()

    DoCmd.OpenForm ("frmTextKI")

End Sub

Private Sub cmbFrage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbKI_Click()

    Me.Refresh
    If get_pKuerzel() = "ROS" Then
        If Me.cboKI.Column(0) = 4 Then
            Me.txtTextKI = Me.txtText

        Else
            If Me.cboKI.Column(1) = "" Then
                Me.txtTextKI = Me.txtText
            Else
                Me.txtTextKI = Me.cboKI.Column(1) & " " & Me.txtText
            End If
        End If
    Else
        If Me.cboKI.Column(1) = "" Then
            Me.txtTextKI = Me.txtText
        Else
            Me.txtTextKI = Me.cboKI.Column(1) & " " & Me.txtText
        End If
    End If
    Me.txtTextKI.SetFocus
    DoCmd.RunCommand acCmdCopy
    Call KI_Internet

End Sub

Private Sub cmbKI_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbLoeschen_Click()

    Me.txtText = ""
    Me.txtTextKI = ""
    Me.txtText.SetFocus

End Sub

Private Sub cmbLoeschen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Form_Close()

    FormKoorSave Me

End Sub

Private Sub Form_Load()

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub
