Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database

Private Sub Abbrechen_Click()

DoCmd.Close

End Sub

Private Sub OK_Click()
Dim ldb As DAO.Database
Dim rsNewFehlerbildAnalyseID As DAO.Recordset
Dim rs As DAO.Recordset
Dim rs2 As DAO.Recordset
Dim l_FehlerbildAnalyseID As Long

 Set ldb = openAccDBConnection("openDB")
If Me.FehlerbildAnalyse <> "" Then
    With Forms("FehlerbilderAnalyse")
        If DialogFehlerbildAnalyse = "Neu" Then
            'Neues Fehlerbild zur Tätigkeit anlegen
            ldb.Execute ("Insert into ZuordnungAnalyseFehlerbild (FehlerbildID, Beschreibung, AnalyseFehlerbild) " & _
                        "values (" & DefTaetigkeitID & ", '" & Me.Beschreibung & "', '" & Me.FehlerbildAnalyse & "');")
            'Ermitteln ID des neuen Fehlerbildes
            Set rsNewFehlerbildAnalyseID = ldb.OpenRecordset("SELECT Max(ZuordnungAnalyseFehlerbild.ID) AS MaxvonID FROM ZuordnungAnalyseFehlerbild;")
            'ID des neu angelegten Fehlerbild sichern
            l_FehlerbildAnalyseID = rsNewFehlerbildAnalyseID("MaxvonID")
        Else
            ldb.Execute ("UPDATE ZuordnungAnalyseFehlerbild set AnalyseFehlerbild='" & Me.FehlerbildAnalyse & "', Beschreibung='" & Me.Beschreibung & "'" & _
                        " WHERE ID =" & FehlerbildAnalyseID & ";")
        End If
        .Requery
        .Refresh
    End With
    DoCmd.Close acForm, "Dialog_FehlerbildAnalyse"
    Forms("FehlerbilderAnalyse").SetFocus
Else
    MsgBox getTC_FCFromCode("TC0335", ldb)  'MsgBox ("Daten nicht vollständig!")   'TC0336
End If
Set rs = Nothing
Set rs2 = Nothing
Set rsNewFehlerbildAnalyseID = Nothing
 ' Set ldb = openAccDBConnection("closeDB")

End Sub
