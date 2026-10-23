Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mlngMID As Long 'ID von Datensatz welcher geändert werden soll

Private Sub cmbUebernehmen_Click()

    Dim ldb As DAO.Database
    Dim lZeit As Double

    lZeit = Nz(DLookup("[Zeit]", "AufwandVorgangPositionen", "[VorgangID] = " & mlngMID & ""))


    If IsNull(Me.cboAuswahlStdSatz) Then
        DoCmd.Close
    Else
        Set ldb = openAccDBConnection("openDB")
        ldb.Execute ("UPDATE [AufwandVorgangPositionen] set [StdSatzID]=  " & Me.cboAuswahlStdSatz & " where [VorgangID] = " & mlngMID & " ;")
        ldb.Execute ("UPDATE [AufwandVorgangPositionen] set [StdSatz]=  " & Me.cboAuswahlStdSatz.Column(2) & " where [VorgangID] = " & mlngMID & " ;")
        ldb.Execute ("UPDATE [AufwandVorgangPositionen] set [DatumStdSatz]=  Date() where [VorgangID] = " & mlngMID & " ;")
        ldb.Execute ("UPDATE [AufwandVorgangPositionen] set [Aufwand]= ccur('" & lZeit * Me.cboAuswahlStdSatz.Column(2) & "') where [VorgangID] = " & mlngMID & " ;")
        DoCmd.Close acForm, "frmAuswahlStundensatz"
        Forms("Fehlererfassung - Register").Controls("UF_Fehlererfassung").SetFocus
        Forms("Fehlererfassung - Register").Controls("UF_Fehlererfassung").Requery
    End If

End Sub



Private Sub Form_Load()

    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long

    lgPosition = getPointerPos(XPos, YPos)
    FormOpenKoorMousPointer Me, XPos, YPos, 450, 120, False

End Sub

Private Sub Form_Open(Cancel As Integer)

    Dim lsarr() As String

    If Me.OpenArgs <> "" Then
        lsarr = Split(Me.OpenArgs, "=")
        mlngMID = CLng(lsarr(1))
    End If

End Sub
