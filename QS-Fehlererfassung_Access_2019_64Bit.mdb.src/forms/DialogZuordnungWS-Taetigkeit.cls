Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub DeleteWsAuswahl_Click()

On Error GoTo err1_DeleteWsAuswahl_Click
Dim ldb As DAO.Database

     Set ldb = openAccDBConnection("openDB")
    'Wenn die Zuordnung einer Taetigkeit zur Worksstation gelöscht wird, müssen folgende abhängige Daten nachfolgend gelöscht werden
    '--> Die Tatetigkeit muss innerhalb der mastertab von allen Vorgängen in den die WS eingetragen ist gelöscht werden
    ldb.Execute ("Update Mastertab set TaetigkeitID = Null, DefTaetigkeitID = Null " & _
                            "Where TaetigkeitID = " & Me.ListWsTaetigkeiten & ";")
    ldb.Execute ("Delete from ZuordnungWsTaetigkeit where WsID=" & Me.ListLQWorkstation2 & " and TaetigkeitID=" & Me.ListWsTaetigkeiten & ";")
    Me.Requery
    Me.Refresh


err1_DeleteWsAuswahl_Click:
     ' Set ldb = openAccDBConnection("closeDB")

End Sub

Private Sub InsertWsAuswahl_Click()

On Error GoTo err1_InsertWsAuswahl_Click
Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")
    ldb.Execute ("insert into ZuordnungWsTaetigkeit ([WsID], [TaetigkeitID]) values (" & Me.ListLQWorkstation2 & ", " & Me.ListOffeneTaetigkeiten & ");")
    Me.Requery
    Me.Refresh


err1_InsertWsAuswahl_Click:
     ' Set ldb = openAccDBConnection("closeDB")

End Sub

Private Sub ListLQBereich2_AfterUpdate()
Me.ListLQFunktion2 = -1
Me.ListLQWorkstation2 = -1
Me.Requery
Me.Refresh
End Sub

Private Sub ListLQEinheit2_AfterUpdate()
'MsgBox (Me.ListLQEinheit.ListIndex)
Me.ListLQBereich2 = -1
Me.ListLQFunktion2 = -1
Me.ListLQWorkstation2 = -1

Me.Requery
Me.Refresh
End Sub


Private Sub ListLQFunktion2_AfterUpdate()
Me.ListLQWorkstation2 = -1
Me.Requery
Me.Refresh

End Sub

Private Sub ListLQWorkstation2_Click()

Me.Requery
Me.Refresh

End Sub

Private Sub Objekt_schließen___Formular_Click()
DoCmd.Close
End Sub
