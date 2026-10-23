Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub MaßnahmeAktiv_AfterUpdate()
    Dim ldb As DAO.Database
    Dim rs  As DAO.Recordset
    'Maßnahme der Fehlerart zuordnen

    s_Fehlerart = Forms("Fehlererfassung - Register").Controls("Fehlerart")
    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from MaßnahmenFehlerarten where [Fehlerart] = '" & s_Fehlerart & "' and MaßnahmeID = " & Me.MaßnahmeID & ";")
    If rs.RecordCount = 0 Then
        ldb.Execute ("Insert into MaßnahmenFehlerarten (Fehlerart, MaßnahmeID) Values ('" & s_Fehlerart & "', " & Me.MaßnahmeID & ");")
    Else
'        MsgBox "Maßnahme ist der Fehlerart: " & s_Fehlerart & " schon zugeordnet"  'TC0290
        MsgBox getTC_FCFromCode("TC0252", ldb) & " " & s_Fehlerart & " " & getTC_FCFromCode("TC0396", ldb)
    End If

    Set rs = Nothing
     ' Set ldb = openAccDBConnection("closeDB")

    DoCmd.Close
    Forms("Fehlererfassung - Register").Refresh


End Sub

Private Sub MaßnahmeAktiv_Click()

    Dim ldb As DAO.Database
    Dim rs  As DAO.Recordset
    'Maßnahme der Fehlerart zuordnen

    s_Fehlerart = Forms("Fehlererfassung - Register").Controls("Fehlerart")
     Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from MaßnahmenFehlerarten where [Fehlerart] = '" & s_Fehlerart & "' and MaßnahmeID = " & Me.MaßnahmeID & ";")
    If rs.RecordCount = 0 Then
        ldb.Execute ("Insert into MaßnahmenFehlerarten (Fehlerart, MaßnahmeID) Values ('" & s_Fehlerart & "', " & Me.MaßnahmeID & ");")
    Else
'        MsgBox "Maßnahme ist der Fehlerart: " & s_Fehlerart & " schon zugeordnet"   'TC0291
        MsgBox getTC_FCFromCode("TC0252", ldb) & " " & s_Fehlerart & " " & getTC_FCFromCode("TC0396", ldb)
    End If

    Set rs = Nothing
     ' Set ldb = openAccDBConnection("closeDB")

    DoCmd.Close
    Forms("Fehlererfassung - Register").Refresh

End Sub

Private Sub DatensatzLöschen_Click()

    Dim ldb As DAO.Database
    Set ldb = openAccDBConnection("openDB")
    ldb.Execute ("Delete from MaßnahmenFehlerarten where " & _
                "Fehlerart = '" & Me.Fehlerklasse & "' and MaßnahmeID = " & Me.MaßnahmeID & ";")

     ' Set ldb = openAccDBConnection("closeDB")
    Me.Requery
    Me.Refresh
End Sub
