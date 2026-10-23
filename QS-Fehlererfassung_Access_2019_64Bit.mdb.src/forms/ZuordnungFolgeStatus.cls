Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Befehl59_Click()

    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")
        ldb.Execute ("insert into tblFolgeStatusAbhängig ([IDVorgangsTyp], [IDAktiveStatus]) values (" & Me.ListVorgangstyp & ", " & Me.ListVorgangsTypStatusAuswahl & ");")
        Me.Requery
        Me.Refresh


End Sub

Private Sub Befehl60_Click()

    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")

        ldb.Execute ("Delete from tblFolgeStatusAbhängig where IDVorgangsTyp =" & Me.ListVorgangstyp & " and IDAktiveStatus=" & Me.ListVorgangsTypStatus & ";")
        Me.Requery
        Me.Refresh

End Sub

Private Sub DeleteWsAuswahl_Click()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset


    Set ldb = openAccDBConnection("openDB")

    Set rs = ldb.OpenRecordset("SELECT * from tblFolgeStatusAbhängig WHERE [IDVorgangsTyp] = " & Me.ListVorgangstyp & " and [IDAktiveStatus] =  " & Me.ListLAktiveStatus & " and [IDFolgeStatus] = " & Me.ListFolgeStatus & " ;")
        rs.Edit
        rs!IDFolgeStatus = Null
        rs.Update
        Me.Refresh
        Me.Requery

    Set rs = Nothing

End Sub

Private Sub Form_Load()

    Dim objMouseOverForm As clsMouseoverForm
    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub

Private Sub InsertWsAuswahl_Click()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset


    Set ldb = openAccDBConnection("openDB")

    Set rs = ldb.OpenRecordset("SELECT  tblFolgeStatusAbhängig.[IDAktiveStatus], tblFolgeStatusAbhängig.[IDVorgangsTyp], tblFolgeStatusAbhängig.[IDFolgeStatus] " & _
                              "from tblFolgeStatusAbhängig " & _
                              "WHERE (tblFolgeStatusAbhängig.[IDAktiveStatus]=" & Me.ListLAktiveStatus & ") AND (tblFolgeStatusAbhängig.[IDVorgangsTyp]=" & Me.ListVorgangstyp & ") AND " & _
                              "(tblFolgeStatusAbhängig.[IDFolgeStatus] Is Null);")



    If rs.RecordCount Then
            rs.Edit
            rs!IDFolgeStatus = Me.ListVorgangStatus
            rs.Update
            Me.Refresh
            Me.Requery
        Else
            rs.AddNew
            rs!IDFolgeStatus = Me.ListVorgangStatus
            rs!IDVorgangsTyp = Me.ListVorgangstyp
            rs!IDAktiveStatus = Me.ListLAktiveStatus
            rs.Update
            Me.Refresh
            Me.Requery
    End If

    Set rs = Nothing

End Sub

Private Sub ListLAktiveStatus_AfterUpdate()

Me.ListFolgeStatus.Requery

End Sub

Private Sub ListVorgangstyp_AfterUpdate()

Me.ListLAktiveStatus.Requery
Me.ListVorgangsTypStatus.Requery
Me.ListVorgangStatus.Requery

End Sub

Private Sub ListVorgangsTypStatus_AfterUpdate()

Me.ListFolgeStatus.Requery
Me.ListVorgangStatus.Requery

End Sub

Private Sub Objekt_schließen___Formular_Click()

    FormKoorSave Me
    DoCmd.Close

End Sub
Private Sub NeuerStatus_Click()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim strInbox As String

'    strInbox = InputBox("Bitte neuen Status eingeben")
    strInbox = InputBox(getTC_FCFromCode("IC3847"), getTC_FCFromCode("IC3848"))

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Kategorien-Vorgangstatus")
        rs.AddNew
        rs![Kategorie-Vorgangstatus] = strInbox
        rs.Update
        Me.Requery

    Set rs = Nothing


End Sub
