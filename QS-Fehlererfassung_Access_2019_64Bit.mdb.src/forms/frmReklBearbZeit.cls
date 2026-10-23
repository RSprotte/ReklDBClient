Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub cmbFirmakonfig_Click()  '20181018-ROS

Dim ldb As DAO.Database
Dim rs As DAO.Recordset


Set ldb = openAccDBConnection("openDB")

    Select Case Rahmen72

        Case 1
            Set rs = ldb.OpenRecordset("Select * from tblReklBearbZeit where [KundenID] = " & Me.cboFirma & ";")

            If rs.RecordCount = 0 Then
                'Kunde-Konfig wird neu angelegt
                ldb.Execute ("Insert into tblReklBearbZeit ([KundenID]) values ( " & Me.cboFirma & ") ;")
                Me.Requery
                Me.Recordset.FindFirst "KundenID= " & Me.cboFirma
            End If

            Me.Filter = "[Firma_Name] = '" & Me.cboFirma.Column(1) & "'"
        Case 2
            Set rs = ldb.OpenRecordset("Select * from tblReklBearbZeit where [LiefID] = " & Me.cboFirma & ";")

            If rs.RecordCount = 0 Then
                'Lieferant-Konfig wird neu angelegt
                ldb.Execute ("Insert into tblReklBearbZeit ([LiefID]) values (" & Me.cboFirma & " ) ;")
                Me.Requery
                Me.Recordset.FindFirst "LiefID= " & Me.cboFirma
            End If

            Me.Filter = "[Lieferantenname] = '" & Me.cboFirma.Column(1) & "' "
    End Select

Me.FilterOn = True
Set rs = Nothing

End Sub

Private Sub cmbReklBearbZeitSchließen_Click()

    FormKoorSave Me
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close

End Sub

Private Sub cmbSpeichernReklBearbZeit_Click()

    DoCmd.RunCommand acCmdSaveRecord
    Me.Requery
    Me.Refresh

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub

Private Sub Form_Load()


    Dim ldb As DAO.Database
    'Dim rs As DAO.Recordset
    Dim objMouseOverForm As clsMouseoverForm

    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

'Set ldb = openAccDBConnection("openDB")
'
'    Dim varKoord As Variant
'    varKoord = getFktKoordinatenRelHF(Me)
'    If varKoord(3) > 0 And varKoord(2) > 0 Then
'        'Fenster setzen relativ zum Hauptfenster TB:autostart
'        DoCmd.MoveSize varKoord(0), varKoord(1), varKoord(2), varKoord(3)
'    End If
'    DoCmd.Restore

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub

Private Sub Form_Open(Cancel As Integer)

    Call Rahmen72_Click

End Sub



Private Sub optLieferant_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub


Private Sub optKunde_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub


Private Sub Rahmen72_Click()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim strListe As String


        Select Case Rahmen72

            Case 1
                Set ldb = openAccDBConnection("openDB")
                Set rs = ldb.OpenRecordset("SELECT [tblFirma].[Firma_ID], [tblFirma].[Firma_Name] FROM tblFirma ORDER BY [Firma_Name];")

                      Do While Not rs.EOF
                         strListe = strListe & rs("Firma_ID") & ";" & rs("Firma_Name") & ";"
                         rs.MoveNext
                      Loop
                   With Me.cboFirma
                      .RowSourceType = "Value List"
                      .ColumnCount = 2
                      .ColumnWidths = "0cm;5cm"
                      .BoundColumn = 1
                      .RowSource = strListe
                   End With

                   Me.cboFirma = Me.cboFirma.ItemData(0)
                   Me.txtFirma.Caption = "Kunde"
                   Me.txtFirmaID.Caption = "Kunde"
                   Me.txtFirmakonfig.Caption = "Konfiguration Kunde"
                   Me.cmbFirmakonfig.Caption = "Kunde konfigurieren"
                   Me.txtLieferantenname.Visible = False
                   Me.Firma_Name.Visible = True

                   Me.Filter = "[KundenID]<>0"
                   Me.FilterOn = True

                Set rs = Nothing
            Case 2
                Set ldb = openAccDBConnection("openDB")
                Set rs = ldb.OpenRecordset("SELECT [Lieferanten].[Lieferant-ID], [Lieferanten].[Lieferantenname], [Lieferanten].[GESPERRT]" & _
                                        "FROM Lieferanten WHERE (((Lieferanten.GESPERRT)=False)) ORDER BY [Lieferantenname];")

                      Do While Not rs.EOF
                         strListe = strListe & rs("Lieferant-ID") & ";" & rs("Lieferantenname") & ";"
                         rs.MoveNext
                      Loop
                   With Me.cboFirma
                      .RowSourceType = "Value List"
                      .ColumnCount = 2
                      .ColumnWidths = "0cm;5cm"
                      .BoundColumn = 1
                      .RowSource = strListe
                   End With

                   Me.cboFirma = Me.cboFirma.ItemData(0)
                   Me.txtFirma.Caption = "Lieferant"
                   Me.txtFirmaID.Caption = "Lieferant"
                   Me.cmbFirmakonfig.Caption = "Lieferant konfigurieren"
                   Me.txtFirmakonfig.Caption = "Konfiguration Lieferant"
                   Me.txtLieferantenname.Visible = True
                   Me.cboFirma.SetFocus
                   Me.Firma_Name.Visible = False

                   Me.Filter = "[LiefID]<>0"
                   Me.FilterOn = True

                Set rs = Nothing
            End Select

End Sub
