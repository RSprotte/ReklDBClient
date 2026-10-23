Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_DblClick(Cancel As Integer)

    OeffneReklamation (Me.ID)

'Dim lsForm As String
'Dim ldb As DAO.Database
'Dim rs As DAO.Recordset
'Dim lReklNr As Long
'Dim sMemo As String
'
'    lReklNr = Me.ID
'
'    DoCmd.Hourglass True
'
'    If IsLoaded("Fehlererfassung - Register") Then
'        '1. Inhalt Dictionary sichern und anschließend leeren
'        UpdatetFormTab Me, "[ID]=" & Forms("Fehlererfassung - Register").Controls("Mastertab#ID")
'        unlockDataset (Forms("Fehlererfassung - Register").Controls("Mastertab#ID"))
'    End If
'
'    p_ReklNr = lReklNr
'
'    Set ldb = openAccDBConnection("openDB")
''        Ermittlung ob es zu dem Vorgang schon ein Protokolleintrag vorhanden ist
''    ReklNr = Forms("Fehlererfassung - Register")!ID
''    Reklnr = Me.Text4       '20181206-ROS
'    If p_KommentarKopieren Then
'        Set rs = ldb.OpenRecordset("Select * from ReklMemo where [ID] = " & lReklNr & ";", dbOpenSnapshot)
'
'        If rs.RecordCount Then
'            If IsNull(rs("Protokoll").Value) Then
'                sMemo = ""
'            Else
'                sMemo = rs("Protokoll").Value
'            End If
'            If Not IsNull(rs("Aktueller Eintrag").Value) Then
'                sMemo = rs("Aktueller Eintrag").Value & Chr(13) & Chr(10) & Chr(13) & Chr(10) & sMemo
'            End If
'            ldb.Execute ("Update ReklMemo set [Protokoll] = '" & sMemo & _
'                        "' where ID = " & lReklNr & ";")
'            sMemo = ""
'            ldb.Execute ("Update ReklMemo set [Aktueller Eintrag] = NULL where ID = " & lReklNr & ";")
'        End If
'    End If
'
'    Set rs = Nothing
'
'    NewReklamationE "Bearbeiten", get_p_ReklNr()
'
'    Call AnsichtFehlerpositionen
'    DoCmd.Hourglass False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub


Private Sub Form_Open(Cancel As Integer)

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

'    Me.Filter = "[Meldungsnummer Kunde] = '" & sReklNrReferenzPos & "'"
    Me.Filter = "[Referenznummer] = '" & sReklNrReferenzPos & "'"
    Me.FilterOn = True
    Me.Requery
    Me.Refresh
    Me.Controls("Schließen").SetFocus

End Sub

Private Sub Schließen_Click()
On Error GoTo Err_Schließen_Click


    FormKoorSave Me
    sReklNrReferenzPos = ""
    DoCmd.Close

Exit_Schließen_Click:
    Exit Sub

Err_Schließen_Click:
    MsgBox Err.Description
    Resume Exit_Schließen_Click

End Sub

Private Sub cmbFehlerUrsacheAnzeigen_Click()

    DoCmd.OpenForm "frmPopUpWieWarumpassiertederFehler", acNormal, , , acFormReadOnly
    With Forms("frmPopUpWieWarumpassiertederFehler")
        .Filter = "[ID] = " & Me.ID
        .FilterOn = True
        .Requery
        .Refresh
    End With

    With Forms("frmPopUpWieWarumpassiertederFehler").Controls("txtWieWarumpassiertederFehler")
        .SetFocus
        .SelStart = 0
        .SelLength = 0
    End With



End Sub

Private Sub Text8_DblClick(Cancel As Integer)

    OeffneReklamation (Me.Text8)

End Sub


Private Sub OeffneReklamation(ByVal fRekID As Long)

Dim lsForm As String
Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim lReklNr As Long
Dim sMemo As String

    lReklNr = fRekID

    DoCmd.Hourglass True

    If IsLoaded("Fehlererfassung - Register") Then
        '1. Inhalt Dictionary sichern und anschließend leeren
        UpdatetFormTab Me, "[ID]=" & Forms("Fehlererfassung - Register").Controls("Mastertab#ID")
        unlockDataset (Forms("Fehlererfassung - Register").Controls("Mastertab#ID"))
    End If

    p_ReklNr = lReklNr

    Set ldb = openAccDBConnection("openDB")
'        Ermittlung ob es zu dem Vorgang schon ein Protokolleintrag vorhanden ist
'    ReklNr = Forms("Fehlererfassung - Register")!ID
'    Reklnr = Me.Text4       '20181206-ROS
    If p_KommentarKopieren Then
        Set rs = ldb.OpenRecordset("Select * from ReklMemo where [ID] = " & lReklNr & ";", dbOpenSnapshot)

        If rs.RecordCount Then
            If IsNull(rs("Protokoll").Value) Then
                sMemo = ""
            Else
                sMemo = rs("Protokoll").Value
            End If
            If Not IsNull(rs("Aktueller Eintrag").Value) Then
                sMemo = rs("Aktueller Eintrag").Value & Chr(13) & Chr(10) & Chr(13) & Chr(10) & sMemo
            End If
            ldb.Execute ("Update ReklMemo set [Protokoll] = '" & sMemo & _
                        "' where ID = " & lReklNr & ";")
            sMemo = ""
            ldb.Execute ("Update ReklMemo set [Aktueller Eintrag] = NULL where ID = " & lReklNr & ";")
        End If
    End If

    Set rs = Nothing

    NewReklamationE "Bearbeiten", get_p_ReklNr()

    Call AnsichtFehlerpositionen
    DoCmd.Hourglass False

End Sub
