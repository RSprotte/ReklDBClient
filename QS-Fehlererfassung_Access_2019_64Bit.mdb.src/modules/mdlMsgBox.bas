Attribute VB_Name = "mdlMsgBox"
Option Compare Database
Option Explicit

Public Type TSize
    X As Double
    Y As Double
End Type

Public Sub TestEigeneMsgBox()
'    EigeneMsgBox "Test ohne Icon nur mit OK-Schaltfläche"
'    EigeneMsgBox "Header@Test ohne Icon nur mit OK-Schaltfläche"
'    EigeneMsgBox "Test mit Icon nur mit OK-Schaltfläche", vbOKOnly
'    EigeneMsgBox "Header@Test mit Icon nur mit OK-Schaltfläche", vbOKOnly
End Sub

'Splittet einen String so in Zeilen auf, dass ein Text entsteht,
' der eine bestimmte maximale Breite hat
'MaxChars: Maximale Zeichenzahl pro Zeile (Vorgabe 80)
'(Sinnvoll etwa für Msgbox)
Function NarrowText(Text As String, Optional MaxChars As Long = 80) As String
    Dim vTxt As Variant
    Dim i As Long, n As Long, m As Long, mOld As Long
    Dim sTmp As String

'    On Error GoTo Fehler

    NarrowText = Text
    vTxt = Split(Text, vbCrLf)
    NarrowText = ""
    For i = 0 To UBound2(vTxt)
        sTmp = vTxt(i)
        m = 1
        If Len(sTmp) > MaxChars Then
            n = InStr(m, sTmp, " ")
            If n = 0 Then
                NarrowText = NarrowText & sTmp & vbCrLf
            Else
                mOld = 1
                m = m + MaxChars
                n = InStrRev(sTmp, " ", m)
                Do While n <> 0
                    NarrowText = NarrowText & Trim(MID(sTmp, mOld, n - mOld)) & vbCrLf
                    mOld = n
                    m = m + MaxChars
                    n = InStrRev(sTmp, " ", m)
                Loop
                NarrowText = NarrowText & Trim(MID(sTmp, mOld + 1)) & vbCrLf
            End If
        Else
            NarrowText = NarrowText & vTxt(i) & vbCrLf
        End If
    Next i
    Do While Right(NarrowText, 2) = Chr(13) & Chr(10)
        NarrowText = MID(NarrowText, 1, Len(NarrowText) - 2)
    Loop
'
'Ende:
'    Exit Function
'
'Fehler:
'    ErrNotify Err, "mdlGlobal", "NarrowText", eHiddenError
'    Resume Ende
End Function

'Ermittelt die horizontale Ausdehnung eines Strings bei angegebenen Schriftattributen
'Rückgabe über benutzerdefinierten Typ TSize (= .x, .y) in Twips
Function GetTextSize(strText As String, FontName As String, FontSize As Long, _
                     Optional IsBold As Boolean, Optional IsItalic As Boolean) As TSize
    Dim H As Long, W As Long, WMax
    Dim vText() As String
    Dim n As Long, i As Long

'    On Error GoTo Fehler

    WizHook.Key = 51488399
    vText = Split(strText, vbCrLf)
    n = UBound2(vText)
    For i = 0 To n
        WizHook.TwipsFromFont FontName, FontSize, 400 + (IsBold * 300), IsItalic, _
                              False, 0, vText(i), 0, W, H
        If W > WMax Then WMax = W
    Next i
    GetTextSize.Y = H * (n + 1)
    GetTextSize.X = WMax

'Ende:
'    Exit Function
'
'Fehler:
'    ErrNotify Err, "mdlGlobal", "GetTextSize", eHiddenError
'    Resume Ende
'
End Function

'Überschreibt die VBA-Methode Msgbox! Ggfls. auskommentieren
Function EigeneMsgBox(prompt, Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
                Optional Title, _
                Optional HelpFile, _
                Optional Context) As VbMsgBoxResult
    On Error Resume Next

    If IsMissing(Title) Then Title = "dmsBase-Meldung"
    DoCmd.OpenForm "frmMsgbox", , , , , acDialog, CStr(Buttons) & "|" & CStr(prompt) & "|" & CStr(Title)
    EigeneMsgBox = Forms!frmMsgbox.Retval
    DoCmd.Close acForm, "frmMsgbox"
End Function

'Gleiche Funktion, wie VBA.UBound(), nur dass kein Fehler bei nicht initialisiertem
' Array auftritt, sondern -1 zurückgegeben wird
Function UBound2(varArray As Variant) As Long
    On Error Resume Next
    UBound2 = UBound(varArray)
    If Err.Number <> 0 Then UBound2 = -1
End Function

Public Sub BeispielTextabmessungen()
    Dim TTextabmessungen As TSize
    TTextabmessungen = GetTextSize("Dies ist ein Beispieltext", "Tahome", "12", False, False)
'    Debug.Print "X-Abmessung: " & TTextabmessungen.x
'    Debug.Print "Y-Abmessung: " & TTextabmessungen.y
End Sub

Public Function Largest(ParamArray Var() As Variant) As Variant
    Dim i As Integer
    Dim varMax As Variant
    varMax = Var(LBound(Var))
    For i = LBound(Var) + 1 To UBound(Var)
        If Var(i) > varMax Then
            varMax = Var(i)
        End If
    Next i
    Largest = varMax
End Function

Public Function Smallest(ParamArray Var() As Variant) As Variant
    Dim i As Integer
    Dim varMin As Variant
    varMin = Var(LBound(Var))
    For i = LBound(Var) + 1 To UBound(Var)
        If Var(i) < varMin Then
            varMin = Var(i)
        End If
    Next i
    Smallest = varMin
End Function
