Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public Textpos As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'MsgBox "Taste: " & KeyCode
If p_ReklMemoQuelle = "Fehlererfassung - Register" Then

    Select Case KeyCode

        Case 9: 'TAB
            KeyCode = 32
            SendKeys "    "
            KeyCode = 0

    End Select
End If

End Sub

Private Sub Form_Open(Cancel As Integer)
Dim Reklnr As String
Dim Kürzel As String
Dim sMemo As String
Dim sOut As String
Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim NewRekl As Boolean

    'ReklNr = InputBox("Geben Sie die Reklamationsnummer ein!", "Eingabe Vorgang")

    If p_ReklMemoQuelle = "Reklamationen_Übersicht" Then
        Me.Filter = "[ID] = " & Forms("Reklamationen_Übersicht").Controls("ID")
        Me.FilterOn = True
        Me.Requery
        Me.Refresh
        Me.AllowAdditions = False

        'Me.Eintrag.Locked = True


    Else
        Me.Filter = "[ID] = " & Forms("Fehlererfassung - Register").Controls("ID")
        Me.FilterOn = True
        Me.Requery
        Me.Refresh
        Me.AllowAdditions = False
        'Me.Eintrag.Locked = False

        'Me.Protokoll.SetFocus
        'Me.Protokoll.SelStart = Len(Me.Protokoll.Text)
        'SendKeys "{UP}" '{UP}{UP}{UP}{UP}{UP}{UP}{UP}"
        'Me.Eintrag.SetFocus
        'Me.Eintrag.SelStart = Len(Me.Eintrag.Text)
    End If






End Sub

Private Sub Protokoll_LostFocus()

'MsgBox "Pos.: " & Me.Protokoll.SelStart
If p_ReklMemoQuelle = "Fehlererfassung - Register" Then
    Textpos = Me.Protokoll.SelStart
End If

End Sub
Private Sub Befehl7_Click()
On Error GoTo Err_Befehl7_Click

Dim sLeft As String
Dim sRight As String

    'Textpos = Me.Protokoll.SelStart
    Me.Protokoll.SetFocus
    Me.Protokoll.SelStart = Textpos
    Me.Protokoll.SelText = "->"
    'Me.Protokoll
    'sLeft = Left(Me.Protokoll.Text, Textpos)
    'sRight = Right(Me.Protokoll.Text, Len(Me.Protokoll.Text) - Textpos)

    'Me.Protokoll.Text = sLeft & Chr(45) & Chr(62) & sRight
    'Me.Protokoll.Requery
     'MsgBox sLeft
     'MsgBox sRight



Exit_Befehl7_Click:
    Exit Sub

Err_Befehl7_Click:
    MsgBox Err.Description
    Resume Exit_Befehl7_Click

End Sub
