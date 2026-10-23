Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public Sub setValue(ByVal CntDate As Date)

    Me.txtDatum = CntDate
    Call ZeitenReklamationen
    Call Datenabgleichen

End Sub
Private Sub chkMittagspause_AfterUpdate()

    Call ZeitenReklamationen
    Call Datenabgleichen

End Sub

Private Sub Form_Close()
    FormKoorSave Me
End Sub

Private Sub Form_Load()

    Me.chkMittagspause = False
    Me.txtDatum = format(Date, "dd/mm/yyyy")
    Call ZeitenReklamationen
    Call Datenabgleichen
    Call FormEinUndAusBlenden(Me)
    FormOpenKoor Me

End Sub

Private Sub Form_Open(Cancel As Integer)

    Me.chkMittagspause = False
    Me.txtDatum = format(Date, "dd/mm/yyyy")
    Call ZeitenReklamationen
    Call Datenabgleichen

End Sub

Private Sub Datenabgleichen()

    Dim Mittag As Double
    Dim Mittagspause As Date
    Dim SummeFehlzeit As Date

    If Me.chkMittagspause = True Then
        Mittag = "0,75"
    Else
        Mittag = "0"
    End If

'20200903-ROS Änderung der Mittagspause kann nun Unabhängig vom Kürzel eingestellt werden
'    If get_pKuerzel = "LH" Then
'        Mittagspause = Me.txtZeitaufwandReklamationen + Me.txtZeitaufwandGU + Me.txtZeitaufwandRep
'    Else
        Mittagspause = Me.txtZeitaufwandReklamationen + Me.txtZeitaufwandGU + Me.txtZeitaufwandRep + format(Mittag / 24)
'    End If

    SummeFehlzeit = Me.txtUhrzeitGehen - Me.txtUhrzeitKommen - Mittagspause

    If (Me.txtUhrzeitGehen - Me.txtUhrzeitKommen) > Mittagspause Then
        Me.txtFehlzeit = SummeFehlzeit
    Else
        Me.txtFehlzeit = 0
    End If


End Sub

Private Sub txtDatum_Change()

    Call ZeitenReklamationen
    Call Datenabgleichen

End Sub



Private Sub txtDatum_Click()

    Dim frm As Form

    DoCmd.OpenForm ("frmCalendar")
    Set frm = Forms("frmCalendar")
    Call frm.setParentForm(Me)

End Sub

Private Sub txtUhrzeitGehen_AfterUpdate()

    Call ZeitenReklamationen
    Call Datenabgleichen

End Sub

Private Sub txtUhrzeitKommen_AfterUpdate()

    Call ZeitenReklamationen
    Call Datenabgleichen

End Sub

Private Sub ZeitenReklamationen()

    Dim ldb As DAO.Database
    Dim rsRekla As DAO.Recordset
    Dim rsGU As DAO.Recordset
    Dim rsRep As DAO.Recordset
    Dim Zeit As Double


    Set ldb = openAccDBConnection("openDB")

'    Set rs = SELECT Sum(AufwandVorgangPositionen.Zeit) AS SummevonZeit From [AufwandVorgangPositionen] WHERE ([AufwandVorgangPositionen.Datum] = Date() AND [AufwandVorgangPositionen.Kürzel]= '" & get_pKuerzel() & "');", dbOpenSnapshot)
    Set rsRekla = ldb.OpenRecordset("SELECT AufwandVorgangPositionen.Datum, Sum(AufwandVorgangPositionen.Zeit) AS SummevonZeit " & _
                    "FROM (Mastertab LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID) LEFT JOIN AufwandVorgangPositionen ON Mastertab.ID = AufwandVorgangPositionen.MID " & _
                    "WHERE (((AufwandVorgangPositionen.Kürzel) = '" & get_pKuerzel() & "') And ((ReklMemo.VorgangstypID) = 1 Or (ReklMemo.VorgangstypID) = 6 Or (ReklMemo.VorgangstypID) = 8)) " & _
                    "GROUP BY AufwandVorgangPositionen.Datum " & _
                    "HAVING (((AufwandVorgangPositionen.Datum)= #" & format(Me.txtDatum, "mm\/dd\/yyyy") & "#));")



    If rsRekla.RecordCount Then
        Zeit = rsRekla("SummevonZeit")
        Me.txtZeitaufwandReklamationen = format(Zeit / 24) + 0
    Else
        Me.txtZeitaufwandReklamationen = 0
    End If

    Zeit = 0

    Set rsRep = ldb.OpenRecordset("SELECT AufwandVorgangPositionen.Datum, Sum(AufwandVorgangPositionen.Zeit) AS SummevonZeit " & _
                "FROM (Mastertab LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID) LEFT JOIN AufwandVorgangPositionen ON Mastertab.ID = AufwandVorgangPositionen.MID " & _
                "WHERE (((AufwandVorgangPositionen.Kürzel) = '" & get_pKuerzel() & "') And (ReklMemo.VorgangstypID) = 3)" & _
                "GROUP BY AufwandVorgangPositionen.Datum " & _
                "HAVING (((AufwandVorgangPositionen.Datum)= #" & format(Me.txtDatum, "mm\/dd\/yyyy") & "#));")

    If rsRep.RecordCount Then
        Zeit = rsRep("SummevonZeit")
        Me.txtZeitaufwandRep = format(Zeit / 24) + 0
    Else
        Me.txtZeitaufwandRep = 0
    End If

    Zeit = 0

    Set rsGU = ldb.OpenRecordset("SELECT AufwandVorgangPositionen.Datum, Sum(AufwandVorgangPositionen.Zeit) AS SummevonZeit " & _
                "FROM (Mastertab LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID) LEFT JOIN AufwandVorgangPositionen ON Mastertab.ID = AufwandVorgangPositionen.MID " & _
                "WHERE (((AufwandVorgangPositionen.Kürzel) = '" & get_pKuerzel() & "') And (ReklMemo.VorgangstypID) = 2)" & _
                "GROUP BY AufwandVorgangPositionen.Datum " & _
                "HAVING (((AufwandVorgangPositionen.Datum)= #" & format(Me.txtDatum, "mm\/dd\/yyyy") & "#));")

    If rsGU.RecordCount Then
        Zeit = rsGU("SummevonZeit")
        Me.txtZeitaufwandGU = format(Zeit / 24) + 0
    Else
        Me.txtZeitaufwandGU = 0
    End If


    Set rsRep = Nothing
    Set rsGU = Nothing
    Set rs = Nothing

End Sub
