Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub Befehl7_Click()

DoCmd.Close

End Sub

Private Sub Befehl8_Click()
Dim ldb As DAO.Database

 Set ldb = openAccDBConnection("openDB")

If MsgBox(getTC_FCFromCode("TC0226", ldb), vbOKCancel) = vbOK Then       'TC0300
    ldb.Execute ("delete * from Lieferanten where [Lieferant-ID] = " & Me.[Lieferant-ID] & ";")
    Me.Requery
    Me.Refresh
End If

 ' Set ldb = openAccDBConnection("closeDB")

End Sub

Private Sub Befehl9_Click()

DoCmd.OpenForm "Lieferanten"
With Forms("Lieferanten")
    .Filter = "[Lieferant-ID] = " & Me.[Lieferant-ID]
    .FilterOn = True
    .Requery
    .Refresh
End With


End Sub

Private Sub FilterLöschen_Click()

Dim sfilter As String

    sfilter = ""
    Me.Filter = ""
    Me.FilterOn = True
    sfilter = ""
    Me.Filter = sfilter
    Me.Requery
    Me.Refresh


End Sub

Private Sub Form_DblClick(Cancel As Integer) '20160802


    Dim ldb As DAO.Database
    Dim lrs As DAO.Recordset
    Dim strKundenIDneu As String
    Dim strKundenIDalt As String

'    Set ldb = openAccDBConnection("openDB")
'    If strKundenIDalt <> "" Then
'        strKundenIDalt = Forms("Fehlererfassung - Register").Form("frmUFKundenDaten").Controls("qryKundenDaten#KndNr_Avista")
'    End If
    If IsNull(Forms("Fehlererfassung - Register").Form("frmUFKundenDaten").Controls("qryKundenDaten#KndNr_Avista")) Then
        strKundenIDalt = ""
    Else
        strKundenIDalt = Forms("Fehlererfassung - Register").Form("frmUFKundenDaten").Controls("qryKundenDaten#KndNr_Avista")
    End If

    If IsLoaded("Fehlererfassung - Register") Then
        Forms("Fehlererfassung - Register").setNewAnsprechpartner (Me.[Kunden-ID])
        Set lrs = getCtrlDataset(Forms("Fehlererfassung - Register"), Forms("Fehlererfassung - Register").Controls("Mastertab#Kunden-ID#1"))
        Set lrs = Nothing

        DoCmd.Close

        strKundenIDneu = Forms("Fehlererfassung - Register").Form("frmUFKundenDaten").Controls("qryKundenDaten#KndNr_Avista")

        If strKundenIDalt <> strKundenIDneu Then
            Set ldb = openAccDBConnection("openDB")
            ldb.Execute ("Delete * from ReklIDtoTagID where [ReklID] = " & get_p_ReklNr() & ";")
            Call TagEintrag(Forms("Fehlererfassung - Register").Controls("Mastertab#ID"), Forms("Fehlererfassung - Register").Form("frmUFKundenDaten").Controls("qryKundenDaten#KndNr_Avista"))
        Else
            Call TagEintrag(Forms("Fehlererfassung - Register").Controls("Mastertab#ID"), Forms("Fehlererfassung - Register").Form("frmUFKundenDaten").Controls("qryKundenDaten#KndNr_Avista"))
        End If
'        Me.Requery
    Else
        DoCmd.OpenForm ("DialogAnsprechpartnerKunde")
    End If

End Sub

Private Sub Neu_Click()

DoCmd.OpenForm "Lieferanten"
DoCmd.GoToRecord , , acNewRec

AllowAdditions = False

End Sub
