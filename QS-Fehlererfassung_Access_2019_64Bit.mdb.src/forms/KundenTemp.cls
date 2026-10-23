Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_DblClick(Cancel As Integer)

Dim ldb As DAO.Database
Dim rs As DAO.Recordset

 Set ldb = openAccDBConnection("openDB")
Set rs = ldb.OpenRecordset("Select * from [Kunden] where [Kurzzeichen] = '" & Me.[Kuerzel] & "';")
If rs.RecordCount Then
    lKundenTemp = rs("Kunden-ID")
Else
    rs.AddNew
        rs![Kontaktperson] = Me.[pa_org_name1] & ", " & Me.[pa_pers_vorname]
        rs![Firma] = "LQ Mechatronik-Systeme GmbH"
        rs![Kurzzeichen] = Me.[pa_org_name2]
        rs![Telefonnummer] = Me.[adb_tel1]
        rs![EmailAdresse] = Me.[adb_email]
        rs![Anmerkungen] = Me.[cmgroup_name]
        rs![Zuordnung] = "LQ"
        rs![Kundennummer] = "01052"
        rs![Avista-ID] = Me.[pa_id]
        lKundenTemp = rs![Kunden-ID]
    rs.Update
    rs.Close
End If

Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")

Me.SetFocus
DoCmd.Close acForm, "KundenTemp"

End Sub

Private Sub Form_Open(Cancel As Integer)

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rs As ADODB.Recordset

    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    'Recordset-Instanz
    Set rs = New ADODB.Recordset
'    Set rs = adbprov.getRecordsetAvista(AvistaArtikelSQL_Select & AID & "' " & AvistaArtikelSQL_Order)
    Set rs = adbprov.getRecordsetAvista(SelectMAKuerzel & sKundenTemp & "%' ")
    If rs.RecordCount Then
        Set Me.Recordset = rs
        Me.Requery
    End If
    rs.Close
    Set rs = Nothing
    Set adbprov = Nothing

ende:
    Exit Sub
Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")     'TC0292
    Resume ende

End Sub
