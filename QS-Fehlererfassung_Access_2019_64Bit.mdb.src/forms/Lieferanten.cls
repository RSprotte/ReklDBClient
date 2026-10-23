Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Form_Close()
    FormKoorSave Me
End Sub

Private Sub Form_Load()

'    Dim varKoord As Variant
'    varKoord = getFktKoordinatenRelHF(Me)
'    If varKoord(3) > 0 And varKoord(2) > 0 Then
'        'Fenster setzen relativ zum Hauptfenster TB:autostart
'        DoCmd.MoveSize varKoord(0), varKoord(1), varKoord(2), varKoord(3)
'    End If
'    DoCmd.Restore

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

End Sub

Private Sub Info_DblClick(Cancel As Integer)

    DoCmd.OpenForm "Lieferanten"
    With Forms("Lieferanten")
        .Filter = "[Lieferant-ID] = " & Me.[Lieferant-ID]
        .FilterOn = True '20190827-ROS_1
        .Requery
        .Refresh
    End With

End Sub

Private Sub Kontrollkästchen178_AfterUpdate()
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim Liefindex As Integer
     Set ldb = openAccDBConnection("openDB")

    If Not IsNull(Me.Kontrollkästchen178) Then
        If Me.Kontrollkästchen178 = True Then
            ldb.Execute ("UPDATE Mastertab SET Mastertab.[Kategorie-Fehleranalyse] = 'Lieferantenfehler LQ' WHERE (Mastertab.[Lieferanten-Code 1]=" & Me.[Lieferant-ID] & ");")
        Else
            ldb.Execute ("UPDATE Mastertab SET Mastertab.[Kategorie-Fehleranalyse] = 'Lieferantenfehler TRUMPF' WHERE (Mastertab.[Lieferanten-Code 1]=" & Me.[Lieferant-ID] & ");")
        End If
    End If

     ' Set ldb = openAccDBConnection("closeDB")

End Sub

Private Sub Lieferant_Name_AfterUpdate()
On Error GoTo err_Lieferant_Name_AfterUpdate

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim Liefindex As Integer
 Set ldb = openAccDBConnection("openDB")


Set rs = ldb.OpenRecordset("Select * from Lieferanten Where Lieferantenname = '" & Me.Lieferantenname & "';")
If rs.RecordCount = 0 And IsNull(Me.[Lieferant-ID]) Then
    Set rs = ldb.OpenRecordset("SELECT Lieferanten.[Lieferant-ID] FROM Lieferanten ORDER BY Lieferanten.[Lieferant-ID] DESC;")
    rs.MoveFirst
    Liefindex = rs(0).Value + 1
    Me.[Lieferant-ID] = Liefindex
    Me.Requery
    Me.Refresh
End If

Set rs = Nothing

end_Lieferant_Name_AfterUpdate:
Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")

Exit Sub

err_Lieferant_Name_AfterUpdate:
If Err = 3314 Then
    GoTo end_Lieferant_Name_AfterUpdate
End If
End Sub

Private Sub Matchcode_AfterUpdate()
On Error GoTo err_Matchcode_AfterUpdate

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim Liefindex As Integer
 Set ldb = openAccDBConnection("openDB")
Dim lResume As Boolean

lResume = False

Set rs = ldb.OpenRecordset("Select * from Lieferanten Where Matchcode = '" & Me.Matchcode & "';")
If rs.RecordCount = 0 And IsNull(Me.[Lieferant-ID]) Then
    Set rs = ldb.OpenRecordset("SELECT Lieferanten.[Lieferant-ID] FROM Lieferanten ORDER BY Lieferanten.[Lieferant-ID] DESC;")
    If rs.RecordCount = 0 Then
        Liefindex = 1
    Else
        rs.MoveFirst
        Liefindex = rs(0).Value + 1
    End If
    Me.[Lieferant-ID] = Liefindex
    Me.Requery
    Me.Refresh
    Me.Filter = "[Lieferant-ID] = " & Liefindex
    Me.FilterOn = True
    Me.AllowAdditions = False
    Me.Requery
End If

Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")

Exit Sub

err_Matchcode_AfterUpdate:
If Err = 3314 Then
    If IsNull(Me.Lieferantenname) Then
        Me.Lieferantenname = "Name erforderlich"
    End If
    If IsNull(Me.[L-Kategorie]) Then
        Me.[L-Kategorie] = "Kategorie erforderlich"
    End If
    If lResume Then
        'MsgBox ("Es fehlen noch mehr Mussfelder!")
        MsgBox getTC_FCFromCode("TC0471")
    Else
        lResume = True
        Resume
    End If
End If

End Sub

Private Sub txt_L_Kategorie_Click()
On Error GoTo err_txt_L_Kategorie

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim Liefindex As Integer
 Set ldb = openAccDBConnection("openDB")


Set rs = ldb.OpenRecordset("Select * from Lieferanten Where Matchcode = '" & Me.Matchcode & "';")
If rs.RecordCount = 0 And IsNull(Me.[Lieferant-ID]) Then
    Set rs = ldb.OpenRecordset("SELECT Lieferanten.[Lieferant-ID] FROM Lieferanten ORDER BY Lieferanten.[Lieferant-ID] DESC;")
    rs.MoveFirst
    Liefindex = rs(0).Value + 1
    Me.[Lieferant-ID] = Liefindex
    Me.Requery
    Me.Refresh
End If

Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")

Exit Sub

err_txt_L_Kategorie:
If Err = 3314 Then
    Resume Next
End If

End Sub
