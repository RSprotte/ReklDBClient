Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim blAufAb As Boolean
Dim Position As Integer
Dim sql As String

Const conSQL_1 = "SELECT tblAccessWsID_AvistaWsID.AvistaWsID, tblAccessWsID_AvistaWsID.AvistaWsName," & _
        "tblAccessWsID_AvistaWsID.AvistaGrName, tblAccessWsID_AvistaWsID.AvistaGrBez " & _
        "FROM tblAccessWsID_AvistaWsID AS tblAccessWsID_AvistaWsID " & _
        "WHERE tblAccessWsID_AvistaWsID.AccessWsID = 0 and tblAccessWsID_AvistaWsID.AvistaStatus = 10 ###Gruppe###" & _
        "ORDER BY tblAccessWsID_AvistaWsID.AvistaWsName;"

Const conSQL_2 = "SELECT LQWorkstation.ID, LQWorkstation.LQWorkstation, LQBereich.LQBereich, LQWorkstation.[ID-LQBereich] " & _
        "FROM LQWorkstation INNER JOIN LQBereich ON LQWorkstation.[ID-LQBereich] = LQBereich.ID " & _
        "WHERE (((LQWorkstation.Aktiv) = True)) ###Bereich### " & _
        "GROUP BY LQWorkstation.ID, LQWorkstation.LQWorkstation, LQBereich.LQBereich, LQWorkstation.[ID-LQBereich], LQWorkstation.WSAvistaID " & _
        "HAVING (((LQWorkstation.WSAvistaID) = 0 Or (LQWorkstation.WSAvistaID) Is Null)) " & _
        "ORDER BY LQWorkstation.LQWorkstation;"


Private Sub Befehl17_Click()

    sql = Replace(conSQL_1, "###Gruppe###", "")
    Me.lstDatenAvista.RowSource = sql
    Me.cboAvistaGruppe = ""
    Me.cboAvistaBezeichnung = ""

End Sub

Private Sub Befehl17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Befehl22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Befehl25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cboAvistaBezeichnung_Click()

    sql = Replace(conSQL_1, "###Gruppe###", "and tblAccessWsID_AvistaWsID.AvistaGrBez = '" & Me.cboAvistaBezeichnung & "'")
    Me.lstDatenAvista.RowSource = sql

End Sub

Private Sub cboAvistaGruppe_AfterUpdate()

    sql = Replace(conSQL_1, "###Gruppe###", "and tblAccessWsID_AvistaWsID.AvistaGrName = '" & Me.cboAvistaGruppe & "'")
    Me.lstDatenAvista.RowSource = sql

End Sub



Private Sub cmbAufAb_Click()

    Position = InStr(1, sql, "ORDER BY tblAccessWsID_AvistaWsID.AvistaWsName", 0)
    If Position > 0 Then
        sql = Replace(sql, "ORDER BY tblAccessWsID_AvistaWsID.AvistaWsName", "ORDER BY tblAccessWsID_AvistaWsID.AvistaWsID Sort")
    Else
        sql = Me.lstDatenAvista.RowSource
    End If

    If blAufAb = True Then
        Position = InStr(1, sql, "Desc", 0)
        If Position > 0 Then
            sql = Replace(sql, "Desc", "Asc")
        Else
            sql = Replace(sql, "Sort", "Asc")
        End If
        Me.cmbAufAb.Caption = "Index Auf"
        blAufAb = False
    Else
        Position = InStr(1, sql, "Asc", 0)
        If Position > 0 Then
            sql = Replace(sql, "Asc", "Desc")
        Else
            sql = Replace(sql, "Sort", "Desc")
        End If
        Me.cmbAufAb.Caption = "Index Ab"
        blAufAb = True
    End If

    Me.lstDatenAvista.RowSource = sql



End Sub

Private Sub Befehl25_Click()

    sql = Replace(conSQL_2, "###Bereich###", "")
    Me.lstDatenReklDB.RowSource = sql
    Me.cboLQ_Bereich = ""

End Sub

Private Sub cboLQ_Bereich_AfterUpdate()

    sql = Replace(conSQL_2, "###Bereich###", "and LQBereich = '" & Me.cboLQ_Bereich & "'")
    Me.lstDatenReklDB.RowSource = sql

End Sub
Private Sub Befehl22_Click()

    sql = Me.lstDatenReklDB.RowSource

    Position = InStr(1, sql, "ORDER BY LQWorkstation.LQWorkstation", 0)
    If Position > 0 Then
        sql = Replace(sql, "ORDER BY LQWorkstation.LQWorkstation", "ORDER BY LQWorkstation.ID Sort")
    Else
        sql = Me.lstDatenReklDB.RowSource
    End If

    If blAufAb = True Then
        Position = InStr(1, sql, "Desc", 0)
        If Position > 0 Then
            sql = Replace(sql, "Desc", "Asc")
        Else
            sql = Replace(sql, "Sort", "Asc")
        End If
        Me.Befehl22.Caption = "Index Auf"
        blAufAb = False
    Else
        Position = InStr(1, sql, "Asc", 0)
        If Position > 0 Then
            sql = Replace(sql, "Asc", "Desc")
        Else
            sql = Replace(sql, "Sort", "Desc")
        End If
        Me.Befehl22.Caption = "Index Ab"
        blAufAb = True
    End If


    Me.lstDatenReklDB.RowSource = sql


End Sub

Private Sub cmbAufAb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbSchließen_Click()
On Error GoTo Err_cmbSchließen_Click


    If Me.Dirty Then Me.Dirty = False
    Me.cboAvistaGruppe = ""
    Me.cboAvistaBezeichnung = ""
    Me.cboLQ_Bereich = ""
    FormKoorSave Me
    DoCmd.Close

Exit_cmbSchließen_Click:
    Exit Sub

Err_cmbSchließen_Click:
    MsgBox Err.Description
    Resume Exit_cmbSchließen_Click

End Sub

Private Sub cmbSchließen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbWsAbgleichen_Click()
On Error GoTo err_cmbWsAbgleichen_Click

    Dim rs As DAO.Recordset
    Dim ldb As DAO.Database
    Dim ls_sql As String
    Set ldb = openAccDBConnection("openDB")

        'Declaration Avista DB provider
    Dim adbprov As clsAvistaProvider
    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider

    ldb.Execute ("Update tblAccessWsID_AvistaWsID set [AccessWsID] = " & Me.lstDatenReklDB & "  where [AvistaWsID] = " & Me.lstDatenAvista & " ;")
    ldb.Execute ("Update LQWorkstation set [WSAvistaID] = " & Me.lstDatenAvista & "  where [ID] = " & Me.lstDatenReklDB & " ;")


    Set rs = ldb.OpenRecordset("SELECT * From qryFehlerbildZuAvistaWs WHERE qryFehlerbildZuAvistaWs.[ID] = " & Me.lstDatenReklDB & ";")
    If rs.RecordCount Then
        rs.MoveFirst
        While Not rs.EOF
            ls_sql = "Insert into zlqpp_fe_ws_fehler_info ([lqpp_ws_id],[fef_id],[fewf_aktiv]) values "
            ls_sql = ls_sql & "(" & rs("WSAvistaID") & "," & rs("EinzelTaetigkeitMelderID") & "," & rs("Aktiv") & ")"
            adbprov.CommandExecuteAvista (ls_sql)
            rs.MoveNext
        Wend
    End If


    adbprov.CommandExecuteAvista ("Insert into zlqpp_fe_ws_bereich ([lqpp_ws_id],[feb_id]) Values (" & Me.lstDatenAvista & ", '" & Me.lstDatenReklDB.Column(3) & "');")

    Set adbprov = Nothing
    Set rs = Nothing

    Me.lstDatenReklDB.Requery
    Me.lstDatenAvista.Requery

    Me.UfrmVerkuepfteWSAvistaReklDB.Requery
    Exit Sub

err_cmbWsAbgleichen_Click:
        Resume Next

End Sub

Private Sub cmbWsAbgleichen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbWSAvistaAbrufen_Click()
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rs As ADODB.Recordset
    Dim ldb As DAO.Database
    Dim rs2 As DAO.Recordset
    Dim rs3 As ADODB.Recordset
    Dim Response As Variant

    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    'Fehlerausgabe unterdrücken
    Response = acDataErrAdded
    'Recordset-Instanz
    Set rs = New ADODB.Recordset
    Set rs = adbprov.getRecordsetAvista("SELECT * FROM zlqpp_ap_kopf")

    Set ldb = openAccDBConnection("openDB")
    DoCmd.Hourglass True
        If rs.RecordCount Then
            While Not rs.EOF
                Set rs2 = ldb.OpenRecordset("Select * from [tblAccessWsID_AvistaWsID] where [AvistaWsID]= " & (rs("lqpp_ap_id")) & ";")
                    If rs2.RecordCount Then
                        rs2.Edit
                        rs2("AvistaWsName") = rs("lqpp_ap_name")
                        rs2("AvistaGrBez") = rs("lqpp_ap_bez")
                        rs2("AvistaStatus") = rs("lqpp_ap_status")
                        rs2.Update
                        rs.MoveNext
                    Else
                        Set rs3 = adbprov.getRecordsetAvista("SELECT * FROM [zlqpp_admin_group_kopf] where [lqpp_group_id]= " & (rs("lqpp_ap_group_id")) & " ")
                        rs2.AddNew
                            rs2("AvistaWsID") = rs("lqpp_ap_id")
                            rs2("AvistaWsName") = rs("lqpp_ap_name")
                            rs2("AvistaStatus") = rs("lqpp_ap_status")
                            If rs("lqpp_ap_group_id") = 0 Then
                                rs2("AvistaGrName") = ""
                            Else
                                rs2("AvistaGrName") = rs3("lqpp_group_name")
                            End If
                            rs2("AvistaGrBez") = rs("lqpp_ap_bez")
                        rs2.Update
                        Response = acDataErrContinue
                        Me.Requery
                        Me.Refresh
                        rs.MoveNext
                    End If
            Wend
                Set rs2 = Nothing
        End If
    DoCmd.Hourglass False
    rs.Close
    Set rs = Nothing
    Set adbprov = Nothing

    Me.lstDatenAvista.Requery
    Me.Refresh

ende:
    Exit Sub

Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0259
    Resume ende





End Sub

Private Sub cmbWSAvistaAbrufen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Form_Load()

    FormOpenKoor Me

End Sub
