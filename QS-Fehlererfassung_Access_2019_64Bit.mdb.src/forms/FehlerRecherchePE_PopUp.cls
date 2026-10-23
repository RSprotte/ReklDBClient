Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Bemerkungen_Vorgang_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_NO)

End Sub

Private Sub ButAbbrechen_Click()

    DoCmd.Close acForm, "FehlerRecherchePE_PopUp", acSaveYes

End Sub

Private Sub ButAbbrechen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub ButOK_Click()

    Call ButUebernehmen_Click
    Call ButAbbrechen_Click

End Sub

Private Sub ButOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub ButUebernehmen_Click()

    Dim ldb As DAO.Database
    Dim lrs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")

    Forms("Fehlererfassung - Register").Controls("ReklMemo#Aktueller Eintrag") = Me.txtFehlerRecherche
    Set lrs = getCtrlDataset(Forms("Fehlererfassung - Register"), Forms("Fehlererfassung - Register").Controls("ReklMemo#Aktueller Eintrag"))

    '20200701-ROS Eingegebener Text wir in Dict abgespeichert, damit dieser beim nächsten öffnen des Textfeldes "ReklMemoAktueller Eintrag" auch sichbar ist
    UpdatetFormTab Forms("Fehlererfassung - Register"), "[ID]=" & Forms("Fehlererfassung - Register").Controls("Mastertab#ID")

    p_KommentarKopieren = True

End Sub



Private Sub ButUebernehmen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub


Private Sub Form_Load()

'    Dim varKoord As Variant
'    varKoord = getFktKoordinatenRelHF(Me)
'    If varKoord(3) > 0 And varKoord(2) > 0 Then
'        'Fenster setzen relativ zum Hauptfenster TB:autostart
'        DoCmd.MoveSize varKoord(0), varKoord(1), varKoord(2), varKoord(3)
'    End If
'    DoCmd.Restore


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
    FormOpenKoorMousPointer Me, XPos - 200, YPos - 150, 710, 730, False

    Me.txtFehlerRecherche = Nz(DLookup("[Aktueller Eintrag]", "ReklMemo", "[ID] = " & get_p_ReklNr() & ""))
    Me.Text6109 = Nz(DLookup("[Protokoll]", "ReklMemo", "[ID] = " & get_p_ReklNr() & ""))
    Me.[Bemerkungen-Vorgang] = Nz(DLookup("[Bemerkungen-Vorgang]", "Mastertab", "[ID] = " & get_p_ReklNr() & ""))

End Sub

Private Sub Text6109_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_NO)

End Sub

Private Sub txtFehlerRecherche_AfterUpdate()

    Forms("Fehlererfassung - Register").Controls("ReklMemo#Aktueller Eintrag") = Me.txtFehlerRecherche

End Sub
