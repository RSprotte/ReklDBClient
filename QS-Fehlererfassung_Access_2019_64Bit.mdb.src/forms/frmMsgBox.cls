Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_ret As VbMsgBoxResult

Property Get Retval() As Long
    Retval = m_ret
End Property



Private Sub cmd1_Click()
    m_ret = Val(Me.cmd1.Tag)
    Me.Visible = False
End Sub

Private Sub cmd2_Click()
    m_ret = Val(Me.cmd2.Tag)
    Me.Visible = False
End Sub

Private Sub cmd3_Click()
    m_ret = Val(Me.cmd3.Tag)
    Me.Visible = False
End Sub



Private Sub Form_Load()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim objMouseOverForm As clsMouseoverForm

    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

End Sub

Private Sub Form_Open(Cancel As Integer)

    Dim strOpenArgs As String
    Dim MsgConst As VbMsgBoxStyle
    Dim strMsg As String
    Dim strTitle As String
    Dim intPosSplitter As Integer
    Dim strText As String
    Dim strHeader As String
    Dim strFontnameText As String
    Dim lngFontsizeText As Long
    Dim bolFontBoldText As Boolean
    Dim bolFontItalicText As Boolean
    Dim strFontnameHeader As String
    Dim lngFontsizeHeader As Long
    Dim bolFontBoldHeader As Boolean
    Dim bolFontItalicHeader As Boolean
    Dim vExtentText As TSize
    Dim vExtentHeader As TSize
    Dim lngHeightFooter As Long
    Dim lngWidthIcon As Long
    Dim lngHeightHeader As Long
    Dim i As Integer
    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long

    Const clngD As Long = 50


    strOpenArgs = CStr(Nz(Me.OpenArgs))

    If Len(strOpenArgs) = 0 Then
        Cancel = True
        Exit Sub
    End If

    MsgConst = Val(Split(strOpenArgs, "|")(0))
    strMsg = CStr(Split(strOpenArgs, "|")(1))
    strTitle = Split(strOpenArgs, "|")(2)

    Me.Caption = strTitle

    strFontnameHeader = Me.lblHeader.FontName
    lngFontsizeHeader = Me.lblHeader.FontSize
    bolFontBoldHeader = Me.lblHeader.FontBold
    bolFontItalicHeader = Me.lblHeader.FontItalic
    strFontnameText = Me.lblText.FontName
    lngFontsizeText = Me.lblText.FontSize
    bolFontBoldText = Me.lblText.FontBold
    bolFontItalicText = Me.lblText.FontItalic

    lngHeightFooter = Me.Section(acFooter).Height

    intPosSplitter = InStr(1, strMsg, "@")
    If intPosSplitter > 0 Then
        strHeader = NarrowText(CStr(Split(strMsg, "@")(0)))
        For i = LBound(Split(strMsg, "@")) + 1 To UBound(Split(strMsg, "@"))
            strText = strText & Split(strMsg, "@")(i) & vbCrLf
        Next i
        Do While Right(strText, 2) = Chr(13) & Chr(10)
            strText = MID(strText, 1, Len(strText) - 2)
        Loop
        strText = Left(strText, Len(strText))
        strText = NarrowText(strText)
        Me.lblHeader.Visible = True
        Me.lblHeader.Caption = strHeader
        vExtentHeader = GetTextSize(strHeader, strFontnameHeader, lngFontsizeHeader, bolFontBoldHeader, bolFontItalicHeader)
        Me.lblHeader.Height = vExtentHeader.Y * 1.2
        Me.lblHeader.Width = vExtentHeader.X
        Me.lblHeader.Top = clngD
        lngHeightHeader = clngD + Me.lblHeader.Height
    Else
        strText = NarrowText(strMsg)
        Me.lblHeader.Visible = False
        Me.lblText.Top = clngD
    End If
    Me.lblText.Caption = strText
    vExtentText = GetTextSize(strText, strFontnameText, lngFontsizeText, bolFontBoldText, bolFontItalicText)
    Me.lblText.Top = lngHeightHeader + clngD
    Me.lblText.Height = vExtentText.Y * 1.2
    Me.lblText.Width = vExtentText.X

    Me.InsideHeight = clngD + vExtentText.Y * 1.2 + clngD + vExtentHeader.Y * 1.2 + clngD + lngHeightFooter

    Me.picCritical.Visible = (vbCritical = (MsgConst And &HF0))
    Me.picExclamation.Visible = (vbExclamation = (MsgConst And &HF0))
    Me.picInformation.Visible = (vbInformation = (MsgConst And &HF0))
    Me.picQuestion.Visible = (vbQuestion = (MsgConst And &HF0))

    Dim bolpic As Boolean
    bolpic = (MsgConst And &HF0) > 0
    If bolpic Then
        lngWidthIcon = clngD + Me.picCritical.Width
    End If
    Me.InsideWidth = clngD + lngWidthIcon + Largest(vExtentText.X * 1.2, vExtentHeader.X, 4000) + clngD
    Me.lblHeader.Left = clngD + lngWidthIcon
    Me.lblText.Left = clngD + lngWidthIcon
    Me.lblHeader.Width = Me.InsideWidth - 2 * clngD - lngWidthIcon
    Me.lblText.Width = Me.lblHeader.Width

    Select Case (MsgConst And &HF)
        Case vbOKOnly
            Me.cmd1.Visible = True
            Me.cmd1.Caption = "OK"
            Me.cmd1.Tag = vbOK
        Case vbYesNo
                Me.cmd1.Visible = True
                Me.cmd1.Caption = getTC_FCFromCode("TC0484")    'Ja
                Me.cmd1.Tag = vbYes
                Me.cmd2.Visible = True
                Me.cmd2.Caption = getTC_FCFromCode("TC0485")    'Nein
                Me.cmd2.Tag = vbNo
                If Me.Caption = "Label" Then
                    Me.cmd1.Visible = True
                    Me.cmd1.Caption = getTC_FCFromCode("TC0484")    'Ja
                    Me.cmd1.Tag = vbYes
                    Me.cmd2.Visible = True
                    Me.cmd2.Caption = getTC_FCFromCode("TC0485")    'Nein
                    Me.cmd2.Tag = vbNo
                End If
        Case vbYesNoCancel
                If Me.Caption = "3D-Report" Then
                    Me.cmd1.Visible = True
                    Me.cmd1.Caption = "DE"
                    Me.cmd1.Tag = vbYes
                    Me.cmd2.Visible = True
                    Me.cmd2.Caption = "EN"
                    Me.cmd2.Tag = vbNo
                    Me.cmd3.Visible = True
                    Me.cmd3.Caption = getTC_FCFromCode("TC0489")    'Abbrechen
                    Me.cmd3.Tag = vbCancel

                ElseIf Me.Caption = "E-Mail an KKI" Then
                    Me.cmd1.Visible = True
                    Me.cmd1.Caption = getTC_FCFromCode("TC0490")    '"Erfassung"
                    Me.cmd1.Tag = vbYes
                    Me.cmd2.Visible = True
                    Me.cmd2.Caption = getTC_FCFromCode("TC0491")    '"Abschluss"
                    Me.cmd2.Tag = vbNo
                    Me.cmd3.Visible = True
                    Me.cmd3.Caption = getTC_FCFromCode("TC0489")    'Abbrechen
                    Me.cmd3.Tag = vbCancel

                ElseIf Me.Caption = "Filter bearbeiten" Then
                    Me.cmd1.Visible = True
                    Me.cmd1.Caption = getTC_FCFromCode("TC0495")    '"De/Aktivieren"
                    Me.cmd1.Tag = vbYes
                    Me.cmd2.Visible = True
                    Me.cmd2.Caption = getTC_FCFromCode("TC0494")    '"Filter Löschen"
                    Me.cmd2.Tag = vbNo
                    Me.cmd3.Visible = True
                    Me.cmd3.Caption = getTC_FCFromCode("TC0489")    'Abbrechen
                    Me.cmd3.Tag = vbCancel

                ElseIf Me.Caption = "Anrede" Then
                    Me.cmd1.Visible = True
                    Me.cmd1.Caption = getTC_FCFromCode("TC0492")    '"Herr"
                    Me.cmd1.Tag = vbYes
                    Me.cmd2.Visible = True
                    Me.cmd2.Caption = getTC_FCFromCode("TC0493")    '"Frau"
                    Me.cmd2.Tag = vbNo
                    Me.cmd3.Visible = True
'                    Me.cmd3.Caption = getTC_FCFromCode("TC0489")    'Abbrechen
                    Me.cmd3.Caption = getTC_FCFromCode("TC0501")    '"Dame/Herr"
                    Me.cmd3.Tag = vbCancel

                ElseIf Me.Caption = "Aktuelle Stückliste" Then
                    Me.cmd1.Visible = True
                    Me.cmd1.Caption = getTC_FCFromCode("TC4189")    'Baugruppe
                    Me.cmd1.Tag = vbYes
                    Me.cmd2.Visible = True
                    Me.cmd2.Caption = getTC_FCFromCode("TC0258")    'Artikel
                    Me.cmd2.Tag = vbNo
                    Me.cmd3.Visible = True
                    Me.cmd3.Caption = getTC_FCFromCode("TC0489")    'Abbrechen
                    Me.cmd3.Tag = vbCancel

                Else
                    Me.cmd1.Visible = True
                    Me.cmd1.Caption = getTC_FCFromCode("TC0484")    'Ja
                    Me.cmd1.Tag = vbYes
                    Me.cmd2.Visible = True
                    Me.cmd2.Caption = getTC_FCFromCode("TC0485")    'Nein
                    Me.cmd2.Tag = vbNo
                    Me.cmd3.Visible = True
                    Me.cmd3.Caption = getTC_FCFromCode("TC0489")    'Abbrechen
                    Me.cmd3.Tag = vbCancel
                End If

        Case vbOKCancel
            Me.cmd1.Visible = True
            Me.cmd1.Caption = "OK"
            Me.cmd1.Tag = vbOK
            Me.cmd2.Visible = True
            Me.cmd2.Caption = getTC_FCFromCode("TC0489")    'Abbrechen
            Me.cmd2.Tag = vbCancel
        Case vbRetryCancel
            Me.cmd1.Visible = True
            Me.cmd1.Caption = getTC_FCFromCode("TC0496")    '"Wiederholen"
            Me.cmd1.Tag = vbRetry
            Me.cmd2.Visible = True
            Me.cmd2.Caption = getTC_FCFromCode("TC0489")    'Abbrechen
            Me.cmd2.Tag = vbCancel
        Case vbAbortRetryIgnore
            Me.cmd1.Visible = True
            Me.cmd1.Caption = getTC_FCFromCode("TC0489")    'Abbrechen
            Me.cmd1.Tag = vbCancel
            Me.cmd2.Visible = True
            Me.cmd2.Caption = getTC_FCFromCode("TC0496")    'Wiederholen
            Me.cmd2.Tag = vbRetry
            Me.cmd3.Visible = True
            Me.cmd3.Caption = getTC_FCFromCode("TC0497")    'Ignorieren
            Me.cmd3.Tag = vbIgnore
    End Select
    If Me.cmd3.Visible = True Then
        Me.cmd2.Left = Me.InsideWidth / 2 - Me.cmd2.Width / 2
        Me.cmd1.Left = Me.cmd2.Left - Me.cmd1.Width - clngD
        Me.cmd3.Left = Me.cmd2.Left + Me.cmd2.Width + clngD
    ElseIf Me.cmd2.Visible = True Then
        Me.cmd1.Left = Me.InsideWidth / 2 - Me.cmd1.Width - clngD / 2
        Me.cmd2.Left = Me.InsideWidth / 2 + clngD / 2
    Else
        Me.cmd1.Left = Me.InsideWidth / 2 - Me.cmd1.Width / 2
    End If


    lgPosition = getPointerPos(XPos, YPos)
    FormOpenKoorMousPointer Me, XPos, YPos, 300, 100


End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub
