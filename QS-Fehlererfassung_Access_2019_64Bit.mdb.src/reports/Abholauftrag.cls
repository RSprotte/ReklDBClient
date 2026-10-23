Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

#If VBA7 Or Win64 Then
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
                                                        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
                                                        ByVal cy As Long, ByVal wFlags As Long) As Long

    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

#Else
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
                                                        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
                                                        ByVal cy As Long, ByVal wFlags As Long) As Long

    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If




Public blFenster As Boolean


Private Const WM_DESTROY = &H2
Private Const WM_CLOSE = &H10
Private Const SWP_NOZORDER As Long = &H4
Private Const SW_HIDE = 0       ' Fenster ausblenden
Private Const SW_NORMAL = 1     ' Fenster wiederherstellen bzw. einblenden
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOW = 5

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Sub cmbBerichtSchließen_Click()

    DoCmd.Close

End Sub

Private Sub cmbMailLogiFuhrpark_Click()

'        Dim olApp As Object
        Dim strBody As String
        Dim Anhang As String
        Dim Datum As String
        Dim strBetreff As String
        Dim sPfadDatei As String
        Dim strKontakt As String
        Dim strCCKontakt As String
        Dim strVon As String


        strVon = getTC_FCFromCode("TC0510")        'service@de.lq-group.com
        strBetreff = "Abholauftrag " & Me.ID
        sPfadDatei = get_DateiPfad & "E-Mail\"
        strKontakt = "Logistik-Fuhrpark@de.lq-group.com"
        strCCKontakt = "Marc.Entenmann@de.lq-group.com"
'        strKontakt = "Roberto.Santarelli@de.lq-group.com"

        Datum = format(Me.txtDatum, "dd.mm.yyyy")

        Anhang = TempVars("Pfad")

        strBody = "<span style=""font-size:10pt; font-family:'Arial'"">" & _
                "Hallo,<br><br>" & "bitte Abholauftrag im Anhang am " & Datum & " durchführen.<br><br>" & "Vielen Dank im Voraus.<br><br>" & _
                "Für Rückfragen stehen wir Ihnen gerne zur Verfügung.<br><br>" & _
                "Mit freundlichen Grüßen <br>" & _
                "i.A. " & "xxxx <br><br>" & _
                "Mitarbeiter(in) Service <br>" & _
                "service@de.lq-group.com _Telefon +49.7143.968440 <br><br>" & _
                "LQ Mechatronik-Systeme GmbH<br>" & _
                "Carl-Benz-Straße 6 _D-74354 Besigheim<br>" & _
                "Fax +49.7143.9683-96 _www.lq-group.com<br><br>" & _
                "<span style=""font-size:7pt; font-family:'Arial'"">" & _
                "LQ Mechatronik-Systeme GmbH, Sitz: Freiberg am Neckar<br>" & _
                "Registergericht: Amtsgericht Stuttgart - HRB-Nr. 205651<br>" & _
                "Geschäftsführer: Friedrich P. Link, José Quero<br><br>" & _
                "Wichtiger Hinweis: Diese E-Mail und etwaige Anlagen enthalten<br>" & _
                "firmenvertrauliche Informationen. Sollten Sie diese E-Mail<br>" & _
                "irrtümlich erhalten haben, benachrichtigen Sie uns bitte durch<br>" & _
                "eine Antwort-Mail und löschen Sie diese E-Mail nebst Anlagen<br>" & _
                "von Ihrem System. Vielen Dank."

        strBody = Replace(strBody, "xxxx", Forms("Fehlererfassung - Register").Controls("Mastertab#Zuständiger Service-MA"))

        Call EMailSendenUndSpeichern(strVon, strKontakt, strBetreff, strBody, sPfadDatei, strCCKontakt, Anhang)

'        Set olApp = CreateObject("Outlook.Application")
'        With olApp.CreateItem(0)
'            .SentOnBehalfOfName = "service@de.lq-group.com"
'            .To = strKontakt
''            .cc = "Marc.Entenmann@de.lq-group.com"
'            .subject = strBetreff
'            .HTMLBody = strBody
'            .Attachments.Add Anhang
'            .Display
'            .SaveAs sPfadDatei & strBetreff & ".msg"
'            .Send
'            MsgBox getTC_FCFromCode("TC0428")
'        End With

        TempVars.Remove ("Pfad")

'        Set olApp = Nothing

        Call cmbBerichtSchließen_Click

End Sub

Private Sub cmbUpdate_Click()

     DoCmd.OpenForm ("frmLager")

End Sub



Private Sub Report_Load()

    Me.cmbSpeichernPDF.Visible = True
    Me.LQ_Art_Nr.Visible = True
    Me.Bezeichnung_1.Visible = True
    Me.BG.Visible = True
    Me.BG_Bezeichnung.Visible = True


    If IsNull(Me.LQ_Art_Nr) Then
        Me.LQ_Art_Nr.Visible = False
        Me.Bezeichnung_1.Visible = False
        Me.ArtKD.Visible = False
        Me.BG.Visible = True
        Me.BG_Bezeichnung.Visible = True
        Me.BGKD.Visible = True
    Else
        Me.LQ_Art_Nr.Visible = True
        Me.Bezeichnung_1.Visible = True
        Me.ArtKD.Visible = True
        Me.BG.Visible = False
        Me.BG_Bezeichnung.Visible = False
        Me.BGKD.Visible = False
    End If

End Sub

Private Sub cmbSpeichernPDF_Click()

        Dim strPfad As String
        Dim Jahr As String
        Dim strDatum As String
        Dim strkdnummer As String
        Dim Reklnr As Long


        Reklnr = Me.ID
        If IsNull(Me.Meldungsnummer_Kunde) Then
            strkdnummer = ""
        Else
            strkdnummer = Me.Meldungsnummer_Kunde
        End If
        strDatum = format(Date, "yyyymmdd")
        Jahr = format(Me.Auftragsdatum, "yyyy")


        Me.ID.SetFocus
'        Me.cmbSpeichernPDF.Visible = False
'        Me.cmbBerichtSchließen.Visible = False
'        Me.cmbUpdate.Visible = False

        strPfad = get_DateiPfad & strDatum & "-" & get_pKuerzel & "_Abholauftrag_" & Reklnr & ".pdf"

        DoCmd.OutputTo acOutputReport, "Abholauftrag", acFormatPDF, strPfad

        TempVars.Add "Pfad", strPfad

'        Me.cmbUpdate.Visible = True
'        Me.cmbSpeichernPDF.Visible = True
'        Me.cmbBerichtSchließen.Visible = True
        Me.cmbBerichtSchließen.SetFocus
'        Me.cmbMailLogiFuhrpark.Visible = True

End Sub

Private Sub Report_Open(Cancel As Integer)
    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long

    Call ReportEinUndAusBlenden(Me)

    lgPosition = getPointerPos(XPos, YPos)
'   Bei Report ist eine Größen anpassung des Fensters nicht möglich
    'Ermittlung des Maximale X+ und X- Koordinate Anzeigebereich (Breite)
'    If XPos + 228 > GetSystemMetrics(78) + GetSystemMetrics(76) Then
'        XPos = (GetSystemMetrics(78) + GetSystemMetrics(76)) - 228
'    End If

'    252/913/210/228/462
'     varKoord(0) = von links ,varKoord(1) = von oben ,varKoord(2) = breite ,varKoord(3) = höhe ,varKoord(4) = von rechts
    XPos = XPos - 250
    YPos = YPos - 250

    SetWindowPos Me.hwnd, 0, XPos, YPos, 0, 0, SWP_NOZORDER
End Sub
