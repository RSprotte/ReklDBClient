Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Dim pfrm As Form

Private Sub cmbCancel_Click()

    psgetValueFromInputBox = "X"
    DoCmd.Close

End Sub

Private Sub cmbCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbOK_Click()


    If IsNull(Me.Controls("txtBoxValue")) Or Me.Controls("txtBoxValue") = "" Then
        psgetValueFromInputBox = ""
    Else
        psgetValueFromInputBox = Me.Controls("txtBoxValue")
    End If

    DoCmd.Close

    'bolStop = False

End Sub

Private Sub cmbOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Drücken der Entertaste

    If KeyCode = 13 Then
        Me.Refresh
        Call cmbOK_Click
    End If

End Sub

Private Sub Form_Open(Cancel As Integer)

    Dim lsOpenArgs() As String
    Dim lsParam() As String
    Dim lsForm As String
    Dim lsCaption As String
    Dim lbKuerzel As Boolean
    Dim lsText As String

    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long
    Dim strWert As String
    Dim varKoord As Variant
    Dim strDaten As String
    Dim varKoordInputbox As Variant


    If Not IsNull(Me.OpenArgs) Then
        lsOpenArgs = Split(Me.OpenArgs, "|")
        For i = 0 To UBound(lsOpenArgs)
            lsParam = Split(lsOpenArgs(i), ":")
            Select Case lsParam(0)
                Case "Form"
                    lsForm = lsParam(1)
                    'Set pfrm = New Form
                Case "Kuerzel"
                    If lsParam(1) = "True" Then
                        lbKuerzel = True
                    Else
                        lbKuerzel = False
                    End If
                Case "Text"
                    Me.txtBoxtext = lsParam(1)
                Case "Caption"
                    Me.Caption = lsParam(1)
            End Select
        Next i
    Else
        DoCmd.Close
    End If

    strDaten = Nz(DLookup("Koordinaten", "Tab_Koordinaten", "[Formname] = '" & lsForm & "'"))
    varKoordInputbox = Split(strDaten, "/")


    lgPosition = getPointerPos(XPos, YPos)

    'Positionierung beim öffnen von Access bei ersten mal Kürzel Eingabe.
    If lbKuerzel = True Then
        strWert = Nz(DLookup("[Koordinaten]", "TAB_Koordinaten", "[Formname] = 'TB:autostart'"))
        varKoord = Split(strWert, "/")
        FormOpenKoorMousPointer Me, varKoord(0) + 300, varKoord(1) + 160, varKoordInputbox(2), varKoordInputbox(3), False
    Else
        FormOpenKoorMousPointer Me, XPos - 190, YPos - 80, varKoordInputbox(2), varKoordInputbox(3), False '380, 160
    End If


'    bolStop = True
'    Forms("frmInputBox").SetFocus
'    While bolStop
'     DoEvents
'    Wend


    Me.SetFocus

End Sub
