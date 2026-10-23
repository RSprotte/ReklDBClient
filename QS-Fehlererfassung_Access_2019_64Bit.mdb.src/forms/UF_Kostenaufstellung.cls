Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private pKostentypID As Integer

Private Sub Form_DblClick(Cancel As Integer)

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim lrs As DAO.Recordset


    Set ldb = openAccDBConnection("openDB")

    If MsgBox(getTC_FCFromCode("TC4212"), vbYesNo, getTC_FCFromCode("TC4208")) = vbYes Then
'        If Me.KostenTypID = 10 Then
'            Forms("Fehlererfassung - Register").Controls("Mastertab#Belastungsbetrag LQ") = 0
'            Set lrs = getCtrlDataset(Forms("Fehlererfassung - Register"), Forms("Fehlererfassung - Register").Controls("Mastertab#Belastungsbetrag LQ"))
'        ElseIf Me.KostenTypID = 11 Then
'            Forms("Fehlererfassung - Register").Controls("Mastertab#Reparaturbetrag") = 0
'            Set lrs = getCtrlDataset(Forms("Fehlererfassung - Register"), Forms("Fehlererfassung - Register").Controls("Mastertab#Reparaturbetrag"))
'        ElseIf Me.KostenTypID = 5 Then
'            Forms("Fehlererfassung - Register").Controls("Mastertab#Material-Kosten#1") = 0
'            Set lrs = getCtrlDataset(Forms("Fehlererfassung - Register"), Forms("Fehlererfassung - Register").Controls("Mastertab#Material-Kosten#1"))
'        End If
        DoCmd.SetWarnings False
        ldb.Execute ("Delete * from tblKostenaufstellung where [KostenaufstellungID] = " & Me.KostenaufstellungID & ";")
        DoCmd.SetWarnings True
    Else
        Exit Sub
    End If

    Set lrs = Nothing
    Me.Requery

    Call PosOrdnen

    Call UpdateKosten

End Sub
Private Sub PosOrdnen()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
        Dim i As Integer

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from tblKostenaufstellung where [ReklID] = " & Forms("Fehlererfassung - Register").Controls("Mastertab#ID") & ";")

    If rs.RecordCount Then
        i = 10
        rs.MoveLast
        rs.MoveFirst
        Do While Not rs.EOF
            rs.Edit
            rs("Pos") = i
            rs.Update
            i = i + 10
            rs.MoveNext
        Loop
    End If

    Me.Requery

    Set rs = Nothing

End Sub
Private Sub KostenTyp_AfterUpdate()

    If Me.KostenTypID = 9 Then
        Me.Kosten = 7   'Pauschalbetrag damit überhaupt Kosten erfasst sind
    Else
        Me.Kosten = 0
    End If

    Call UpdateKosten

End Sub

Private Sub Kosten_AfterUpdate()

    Call UpdateKosten

End Sub
Private Sub UpdateKosten()


'    Dim lrs As DAO.Recordset
'
'    If fKostenTyp = 5 Then
'        Forms("Fehlererfassung - Register").Controls("Mastertab#Material-Kosten#1") = fKosten
'        Set lrs = getCtrlDataset(Forms("Fehlererfassung - Register"), Forms("Fehlererfassung - Register").Controls("Mastertab#Material-Kosten#1"))
'    End If
'
'    If fKostenTyp = 10 Then
'        Forms("Fehlererfassung - Register").Controls("Mastertab#Belastungsbetrag LQ") = fKosten
'        Set lrs = getCtrlDataset(Forms("Fehlererfassung - Register"), Forms("Fehlererfassung - Register").Controls("Mastertab#Belastungsbetrag LQ"))
'    End If
'
'    If fKostenTyp = 11 Then
'        Forms("Fehlererfassung - Register").Controls("Mastertab#Reparaturbetrag") = fKosten
'        Set lrs = getCtrlDataset(Forms("Fehlererfassung - Register"), Forms("Fehlererfassung - Register").Controls("Mastertab#Reparaturbetrag"))
'    End If
'
'    Set lrs = Nothing
'
    Forms("Fehlererfassung - Register").Controls("txtBearbeitungKosten").SetFocus
    Forms("Fehlererfassung - Register").Controls("txtBearbeitungKosten").Requery
    Forms("Fehlererfassung - Register").Controls("txt_Kosten-Aufwand").SetFocus
    Forms("Fehlererfassung - Register").Controls("txt_Kosten-Aufwand").Requery


End Sub



Private Sub KostenTyp_Dirty(Cancel As Integer)

    pKostentypID = Me.KostenTypID

End Sub
