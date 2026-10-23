Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_DblClick(Cancel As Integer)
'20160510-RSP

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
Dim rsAvista As ADODB.Recordset
Dim sSQL As String
Dim lbLeitersatz As Boolean

lbLeitersatz = False


Set ldb = openAccDBConnection("openDB")
If sCurrArtikelSource = "Artikel" Then
    Set rs = ldb.OpenRecordset("Select * from [Artikel-all_avista_out] where [LQ Art-Nr] = '" & Me.[LQ_Art_Nr] & "' and [Variante] = '" & Me.[Var Nr] & "' and [Kd Mat-Nr] = '" & Me.[Kunden-Nr] & "';")
    If rs.RecordCount Then
        rs.Edit
    Else

        rs.AddNew
        rs![LQ Art-Nr] = Me.[ast_nummer]
        If Not IsNull(Me.[avar_nummer]) Then
            rs![Variante] = Me.[avar_nummer]
        End If
        rs![Kd Mat-Nr] = Me.[avkd_artnr_kd]
    End If
    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    'Recordset-Instanz
    Set rsAvista = New ADODB.Recordset
    Set rsAvista = adbprov.getRecordsetAvista("Select tm_artikel.ast_nummer, tm_artikel.ast_bezeichnung AS artikel_bezeichnung from tm_artikel WHERE ast_nummer = '" & Me.[LQ_Art_Nr] & "' AND tm_artikel.ast_id IN ( SELECT xm_artikel.ast_id FROM DBA.xm_artikel  WHERE xm_artikel.ast_artikel_typ = 1)")

    If rsAvista.RecordCount Then
        lbLeitersatz = True
    End If
    Set rsAvista = Nothing
    adbprov.closeAvistaConnection
    rs![Bezeichnung 1] = Me.[ast_bezeichnung]
    rs![Hersteller-Teilenummer] = Me.[avar_herst_nummer]

    rs![Kd Mat-Nr] = Me.[avkd_artnr_kd]

    lLQArtikelTemp = rs![Artikel-ID]
    If lbLeitersatz Then rs![AvistaArtikelTyp] = 1
    rs.Update
    rs.Close
Else
    Set rs = ldb.OpenRecordset("Select * from [Baugruppen-avista_out] where [LQ Art-Nr] = '" & Me.[LQ_Art_Nr] & "' and [Var Nr] = '" & Me.Var_Nr & "' and [Kunden-Nr] = '" & Me.[Kunden-Nr] & "';")
    If rs.RecordCount Then
        rs.Edit
    Else
        rs.AddNew
        rs![LQ Art-Nr] = Me.[ast_nummer]
        If Not IsNull(Me.[avar_nummer]) Then
            rs![Var Nr] = Me.[avar_nummer]
        End If
        rs![Kunden-Nr] = Me.[avkd_artnr_kd]
    End If
    rs![Bezeichnung 1] = Me.[ast_bezeichnung]
    rs![Produktgruppe] = Me.[ast_bezeichnung]
    If IsNull(Me.[ast_bezeichnung]) Then
        rs![Produktgruppe] = "-"
    Else
        rs![Produktgruppe] = Me.[ast_bezeichnung]
    End If
    lLQArtikelTemp = rs![Baugruppen-ID]
    rs.Update
    rs.Close
End If

Forms("Fehlererfassung - Register").Refresh

Set rs = Nothing
Set adbprov = Nothing
'closeAccDBConnection db

Me.SetFocus
DoCmd.Close acForm, "Artikel-avista_outTemp"

End Sub

Private Sub Form_Open(Cancel As Integer)

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rs As ADODB.Recordset
    Dim sSQL As String

    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long

    lgPosition = getPointerPos(XPos, YPos)

    FormOpenKoorMousPointer Me, XPos, YPos, 400, 500, False

    CurrentDb.Execute ("delete * from tblArtNrTemp;")
    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    'Recordset-Instanz
    Set rs = New ADODB.Recordset
    If IsLoaded("FehlererfassungIntern") Then
        If Forms("FehlererfassungIntern").Controls("AuswArtikelKND-LQ") = 1 Then '1 = LQ
        '    Set rs = adbprov.getRecordsetAvista(AvistaArtikelSQL_Select & AID & "' " & AvistaArtikelSQL_Order)
            Set rs = adbprov.getRecordsetAvista(AvistaArtikelSQL_Select & sLQArtikelTemp & "' " & AvistaArtikelSQL_Order)
            If rs.RecordCount Then
                Set Me.Recordset = rs
                Me.Requery
            End If
        Else
         '    Set rs = adbprov.getRecordsetAvista(AvistaArtikelSQL_Select & AID & "' " & AvistaArtikelSQL_Order)
            Set rs = adbprov.getRecordsetAvista(AvistaKndArtNrSQL_Select & sLQArtikelTemp & "' " & AvistaArtikelSQL_Order)
            If rs.RecordCount Then
                Set Me.Recordset = rs
                Me.Requery
            End If
        End If
    Else
        Set rs = adbprov.getRecordsetAvista(AvistaArtikelSQL_Select & sLQArtikelTemp & "' " & AvistaArtikelSQL_Order)   '20210428-ROS
        If rs.RecordCount Then
            rs.MoveFirst
            While Not rs.EOF
                 CurrentDb.Execute ("Insert into tblArtNrTemp ([ast_nummer],[avar_nummer],[avkd_artnr_kd],[pg_bez],[ast_bezeichnung],[avar_herst_nummer]) values ('" & rs.Fields(0).Value & "', '" & rs(1) & "', '" & rs(3) & "', '" & rs(5) & "','" & rs(4) & "','" & rs(2) & "');")
                rs.MoveNext
            Wend
'            Set Me.Recordset = rsTemp
            Me.Requery
        End If
    End If
    rs.Close
    Set rs = Nothing
    Set adbprov = Nothing

ende:
    Exit Sub
Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0228
'    MsgBox Err.Description, 16, "Fehler"
    Resume ende

End Sub
