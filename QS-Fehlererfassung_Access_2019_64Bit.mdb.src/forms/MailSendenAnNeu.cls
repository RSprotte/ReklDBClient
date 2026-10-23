Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim gsaAnhang() As String
Dim gsaColAnhang As Collection
'Dim sAnhang2 As String
'Dim m_parentForm As Form_frmTicketListe
Dim m_parentForm As Form_frmTicketErstellen '20251216-ROS
Dim m_TicketNew As Boolean

'EventHandler
Public Event MessageChanged(Control As String, Message As String)


Private Sub BefLoescheAnhang_Click()


Dim lItem As Object
Dim i, i2 As Integer
Dim sSelectedItem As String
Dim sFile As String
Dim sPfadDatei As String


sSelectedItem = Me.ListAnhaenge.ItemData(Me.ListAnhaenge.ListIndex)
gsaColAnhang.Remove (sSelectedItem)
Me.ListAnhaenge.RemoveItem (Me.ListAnhaenge.ListIndex)

RaiseEvent MessageChanged("DeleteAnhang", "")

End Sub

Private Sub BefNeuerAnhang_Click()
On Error GoTo err_BefNeuerAnhang_Click

Dim i As Long
Dim sAnhang As String
Dim sAnhang2 As String
Dim sFile As String
Dim n As Long
Dim sPfadDatei As String
Dim sFiles As Object
Dim sFile2 As Variant
Dim ReklnrKnd As String

    With Forms("Fehlererfassung - Register")        '20190827-ROS
        If IsNull(.Controls("Mastertab#Meldungsnummer Kunde#2")) Or .Controls("Mastertab#Meldungsnummer Kunde#2") = "" Then
            sPfadDatei = getReklVerzeichnis(.Controls("Mastertab#ID"), format(.Controls("Mastertab#Auftragsdatum#3"), "yyyy"), "")
        Else
            sPfadDatei = getReklVerzeichnis(.Controls("Mastertab#ID"), format(.Controls("Mastertab#Auftragsdatum#3"), "yyyy"), .Controls("Mastertab#Meldungsnummer Kunde#2"))
        End If
    End With

    Set sFiles = DateiMultiOeffnen(, , , sPfadDatei)

    For Each sFile2 In sFiles.SelectedItems
'        Debug.Print sFile2
    Next

    If Not IsNull(sFiles) Then
        For Each sFile2 In sFiles.SelectedItems
            n = InStrRev(sFile2, "\")
            If n > 0 Then
                sFile = Right$(sFile2, Len(sFile2) - n)
            Else
                sFile = sFile2
            End If
            gsaColAnhang.Add sFile2, sFile
            Me.ListAnhaenge.AddItem sFile
            RaiseEvent MessageChanged("AnhangNeu", "")
'            Debug.Print sFile2
        Next
    Else
        MsgBox getTC_FCFromCode("TC0293")
        'MsgBox "Keine Datei ausgewählt"    'TC0293
    End If

end_BefNeuerAnhang_Click:
    Exit Sub
err_BefNeuerAnhang_Click:
    MsgBox "Err: " & Err & " - " & Error
    GoTo end_BefNeuerAnhang_Click

End Sub

Private Sub BefNeuerAnhang_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbKontaktpersonAuswählen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbKontaktpersonAuswählenBCC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbKontaktpersonAuswählenCC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbKorkenzieher_Click()

    Me.txtBody = Me.txtBody & vbCrLf & vbCrLf & getTC_FCFromCode("TC0514")
    Me.txtBody.SetFocus
    Me.txtBody.SelStart = Len(Me.txtBody)
'    Me.txtBody.SelStart = 0
    SendKeys "{ENTER}"     ' optional: Enter drücken
    SendKeys "{NUMLOCK}"

End Sub

Private Sub cmbKorkenzieher_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Form_Close()
    On Error Resume Next

    FormKoorSave Me
    m_parentForm.MailSendenAnNeuClosed = "Close"

End Sub

Private Sub Form_Load()


    ReDim Preserve gsaAnhang(0)
    Set gsaColAnhang = New Collection


    Dim objMouseOverForm As clsMouseoverForm
    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

'    Dim varKoord As Variant
'    varKoord = getFktKoordinatenRelHF(Me)
'    If varKoord(3) > 0 And varKoord(2) > 0 Then
'        'Fenster setzen relativ zum Hauptfenster TB:autostart
'        DoCmd.MoveSize varKoord(0), varKoord(1), varKoord(2), varKoord(3)
'    End If
'
'    DoCmd.Restore

    If get_ServiceZuordnung = True And getCurrLanguage = "German" Then
        Me.cmbKorkenzieher.Visible = True
    Else
        Me.cmbKorkenzieher.Visible = False
    End If

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub


Private Sub SendXDRanLiefMail_Click()

On Error GoTo Err_Befehl10_Click

    Dim lclsMailProv As clsMailProviderOutlook
    Dim i As Integer

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim strTextHtml As String

     'Html File erzeugen für E-Mail
    HtmlTextCreate Me.txtBody, strTextHtml
    Me.txtHTMLText = strTextHtml
    'Ändeerung Parent mitteilen
    RaiseEvent MessageChanged(Me.txtHTMLText.Name, Me.txtHTMLText)
    'Instanz auf einen Outlook MailProvider
    Set lclsMailProv = New clsMailProviderOutlook
    'Mailinhalte füllen
    lclsMailProv.setAn = Me.txtAN
    lclsMailProv.setVon = Me.txtVon
    If Not IsNull(Me.txtCC) Then
             lclsMailProv.setCC = Me.txtCC
         Else
             lclsMailProv.setCC = ""
    End If
    lclsMailProv.setBCC = ""
    lclsMailProv.setBetreff = Me.txtBetreff
    lclsMailProv.setBody = Me.txtHTMLText & "<br>" & "<br>" & "<br>" & "<br>" & Me.txtLinkZumVorgangsordner
    lclsMailProv.fillMailItem
    'Anhänge füllen
    For i = 1 To gsaColAnhang.Count
        lclsMailProv.AddItemMailItem (gsaColAnhang.Item(i))
    Next i
    'Mail im Vorgangsordner sichern
    lclsMailProv.SaveAsMailItem get_DateiPfad & "E-Mail\", get_pKuerzel()
    'Abfrage ob der Entwicklermodus aktiviert ist. Wenn ja, dann Mail nicht versenden
    If DLookup("Value", "KonfigurationReklDBLocal", "Konfiguration = 'Entwicklung'") = "Nein" Then
        DoCmd.Hourglass True
        lclsMailProv.sendMailItem
        DoCmd.Hourglass False
    End If
    Set lclsMailProv = Nothing
    RaiseEvent MessageChanged(Me.ActiveControl.Name, "Gesendet")


    FormKoorSave Me
    m_parentForm.closeMailSendenAnNeu = "Close"
    closeWinHandle Me

    Exit Sub

Exit_Befehl10_Click:

    Exit Sub

Err_Befehl10_Click:
    MsgBox Err.Description  'FC0161
    Resume Exit_Befehl10_Click


End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub


Public Property Let parentForm(fparentForm As Form)

    If m_parentForm Is Nothing Then Set m_parentForm = fparentForm

End Property


Public Sub loadForm(ByRef frs As DAO.Recordset, Optional ByRef frs2 As DAO.Recordset = Null)
'Prozedur lädt das Mailformalar mit den notwendigen Informationen wie
'Formale Parameter:
'fRS:
'fRS2:
'Adresse
'Absender
'Mailtext
'und Anhänge
'########################################################

    On Error Resume Next

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsReklBericht As DAO.Recordset
    Dim sFile As String
    Dim sPfad As String
    Dim sPfadDatei As String
    Dim n As Long
    Dim lTicketID As Long  'RSP_8_2_5

    lTicketID = 0 'RSP_8_2_5

    Set ldb = openAccDBConnection("openDB")
    '## Prüfen ob es sich um eine Lieferantenticket == 2 handelt
    '## --> Laden Report
    If frs("TicketTypID") = 2 Then
        Set rs = ldb.OpenRecordset("Select * from Mastertab where [ID] = " & frs("ReklID") & ";")
        If rs.RecordCount > 0 Then
           If IsNull(rs("Meldungsnummer Kunde")) Or rs("Meldungsnummer Kunde") = "" Then
                sPfadDatei = getReklVerzeichnis(rs("ID"), format(rs("Auftragsdatum"), "yyyy"), "")
            Else
                sPfadDatei = getReklVerzeichnis(rs("ID"), format(rs("Auftragsdatum"), "yyyy"), rs("Meldungsnummer Kunde"))
            End If
        End If
        Set rs = Nothing

'Alter Code
'        Me.ListAnhaenge.AddItem frs("ReklID") & "_Reklamationsbericht_Lieferant.xlsx"
'        gsaColAnhang.Add sPfadDatei & frs("ReklID") & "_Reklamationsbericht_Lieferant.xlsx", sFile

'20180119-ROS Reklamationsbericht in tblTicketAnhaenge anlegen wenn nicht vorhanden
'###################################
        'RSP_8_2_5
        Set rsReklBericht = ldb.OpenRecordset("Select * from tblTicketAnhaenge where TicketID = " & frs("TicketID") & " and " & _
                            "ReklID = " & frs("ReklID") & " and AnhangPfad = '" & sPfadDatei & frs("ReklID") & "_Reklamationsbericht_Lieferant.pdf" & "';")
        If rsReklBericht.RecordCount = 0 Then
            rsReklBericht.AddNew
            rsReklBericht("TicketID") = frs("TicketID")
            rsReklBericht("ReklID") = frs("ReklID")
            rsReklBericht("AnhangPfad") = sPfadDatei & frs("ReklID") & "_Reklamationsbericht_Lieferant.pdf"
            rsReklBericht.Update
            'db.Execute ("Insert into tblTicketAnhaenge ([TicketID],[ReklID], [AnhangPfad]) values (" & Forms("frmTicketListe")![TicketID] & " ," & fRS("ReklID") & ", '" & sPfadDatei & fRS("ReklID") & "_Reklamationsbericht_Lieferant.xlsx" & "');")
        End If
        Set rsReklBericht = Nothing
'####################################

    End If

    'Laden Formular mit den notwendigen Informationen
    Set rs = ldb.OpenRecordset("select * From tblTicketAnhaenge where TicketID = " & frs("TicketID") & " ;")
    If Not rs Is Nothing Then
        While Not rs.EOF
            sPfad = rs("AnhangPfad")
            n = InStrRev(rs("AnhangPfad"), "\")
            sFile = Right$(rs("AnhangPfad"), Len(rs("AnhangPfad")) - n)
            Me.ListAnhaenge.AddItem sFile
            'Container der Anhänge füllen
            gsaColAnhang.Add sPfad, sFile
            rs.MoveNext
        Wend
    End If
    Set rs = Nothing


    Me.txtAN = frs("AN")
    Me.txtCC = frs("CC")
    Me.txtBetreff = frs("Betreff")
    Me.txtBody = frs("Body")
    Me.txtBCC = frs("BCC")
    Me.txtLinkZumVorgangsordner = frs("Link")
    If get_ServiceZuordnung = True And getCurrLanguage = "German" Then
        Me.txtVon = getTC_FCFromCode("TC0510")        'service@de.lq-group.com
    Else
        Me.txtVon = frs("VON")
    End If

    Me.txtTicketID = frs("TicketID")

End Sub

Public Property Get GettxtCC() As String
    GettxtCC = Me.txtCC
End Property
Public Property Let LettxtCC(fs As String)
    Me.txtCC = fs
End Property
Public Property Get GettxtBCC() As String
    GettxtBCC = Me.txtBCC
End Property
Public Property Let LettxtBCC(fs As String)
    Me.txtBCC = fs
End Property
Public Property Get GettxtAN() As String
    GettxtAN = Me.txtAN
End Property
Public Property Let LettxtAN(fs As String)
    Me.txtAN = fs
End Property
Public Property Get GettxtVON() As String
    GettxtVON = Me.txtVon
End Property
Public Property Let LettxtVON(fs As String)
    Me.txtVon = fs
End Property
Public Property Get GettxtBetreff() As String
    GettxtBetreff = Me.txtBetreff
End Property
Public Property Let LettxtBetreff(fs As String)
    Me.txtBetreff = fs
End Property
Public Property Get GettxtBody() As String
    GettxtBody = Me.txtBody
End Property
Public Property Let LettxtBody(fs As String)
    Me.txtBody = fs
End Property
Public Property Get GetNewAnhang() As String

    GetNewAnhang = gsaColAnhang.Item(gsaColAnhang.Count)

End Property
Public Property Get GetAnhangByItem(fItem As Integer) As String
On Error GoTo err_GetAnhang
'    GetNewAnhang = gsaAnhang(UBound(gsaAnhang))
    GetAnhangByItem = gsaColAnhang.Item(fItem)
    Exit Property
err_GetAnhang:
    GetAnhangByItem = ""
End Property


Private Sub SendXDRanLiefMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub txtAn_LostFocus()

    RaiseEvent MessageChanged(Me.ActiveControl.Name, Me(Me.ActiveControl.Name).Text)

End Sub

Private Sub txtBCC_LostFocus()

    RaiseEvent MessageChanged(Me.ActiveControl.Name, Me(Me.ActiveControl.Name).Text)

End Sub

Private Sub txtCC_LostFocus()

    RaiseEvent MessageChanged(Me.ActiveControl.Name, Me(Me.ActiveControl.Name).Text)

End Sub

Private Sub txtVon_AfterUpdate()

    RaiseEvent MessageChanged(Me.ActiveControl.Name, Me(Me.ActiveControl.Name).Text)
    'm_parentForm.MailSendenChanged = "txtVon"

End Sub
Private Sub txtAn_AfterUpdate()

    RaiseEvent MessageChanged(Me.ActiveControl.Name, Me(Me.ActiveControl.Name).Text)
    'm_parentForm.MailSendenChanged = "txtAn"

End Sub
Private Sub txtCC_AfterUpdate()

    RaiseEvent MessageChanged(Me.ActiveControl.Name, Me(Me.ActiveControl.Name).Text)
'    m_parentForm.MailSendenChanged = "txtCC"

End Sub
Private Sub txtBCC_AfterUpdate()

    RaiseEvent MessageChanged(Me.ActiveControl.Name, Me(Me.ActiveControl.Name).Text)
'    m_parentForm.MailSendenChanged = "txtBCC"

End Sub
Private Sub txtBetreff_AfterUpdate()

    RaiseEvent MessageChanged(Me.ActiveControl.Name, Me(Me.ActiveControl.Name).Text)
'    m_parentForm.MailSendenChanged = "txtBetreff"

End Sub
Private Sub txtBody_AfterUpdate()

    RaiseEvent MessageChanged(Me.ActiveControl.Name, Me(Me.ActiveControl.Name).Text)
'    m_parentForm.MailSendenChanged = "txtBody"

End Sub
Private Sub ListAnhaenge_AfterUpdate()

    RaiseEvent MessageChanged("ListAnhaenge", "")
'    m_parentForm.MailSendenChanged = "ListAnhaenge"

End Sub
Private Sub cmbKontaktpersonAuswählen_Click()

    DoCmd.OpenForm "frmEmailAdresse", acNormal, , , , , Me.txtAN.Name
'    Me.SetFocus
'    Dim varKoord As Variant
'    varKoord = getFktKoordinatenRelHF(Me)
'    If varKoord(3) > 0 And varKoord(2) > 0 Then
'        'Fenster setzen relativ zum Hauptfenster TB:autostart
'        DoCmd.MoveSize varKoord(0), varKoord(1), varKoord(2), varKoord(3)
'    End If
'    DoCmd.Restore
    Forms("frmEmailAdresse").SetFocus

End Sub

Private Sub cmbKontaktpersonAuswählenBCC_Click()

    DoCmd.OpenForm "frmEmailAdresse", acNormal, , , , , Me.txtBCC.Name
'    Me.SetFocus
'    Dim varKoord As Variant
'    varKoord = getFktKoordinatenRelHF(Me)
'    If varKoord(3) > 0 And varKoord(2) > 0 Then
'        'Fenster setzen relativ zum Hauptfenster TB:autostart
'        DoCmd.MoveSize varKoord(0), varKoord(1), varKoord(2), varKoord(3)
'    End If
'    DoCmd.Restore
    Forms("frmEmailAdresse").SetFocus

End Sub

Private Sub cmbKontaktpersonAuswählenCC_Click()

    DoCmd.OpenForm "frmEmailAdresse", acNormal, , , , , Me.txtCC.Name
'    Me.SetFocus
'    Dim varKoord As Variant
'    varKoord = getFktKoordinatenRelHF(Me)
'    If varKoord(3) > 0 And varKoord(2) > 0 Then
'        'Fenster setzen relativ zum Hauptfenster TB:autostart
'        DoCmd.MoveSize varKoord(0), varKoord(1), varKoord(2), varKoord(3)
'    End If
'    DoCmd.Restore
    Forms("frmEmailAdresse").SetFocus

End Sub

Public Property Get GetTicketNew() As Boolean
    GetTicketNew = m_TicketNew
End Property
Public Property Let LetTicketNew(fTicketNew As Boolean)
    m_TicketNew = fTicketNew
End Property

Public Sub setFocusMe()

    Me.SetFocus

End Sub
