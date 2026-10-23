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

Private Sub Abbrechen_Click()

'    FormKoorSave Me
    DoCmd.Close acForm, "DialogAnsprechpartnerKunde"

End Sub

Private Sub Form_Close()
    FormKoorSave Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub

Private Sub Form_Open(Cancel As Integer)

    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long


    lgPosition = getPointerPos(XPos, YPos)
    FormOpenKoorMousPointer Me, XPos, YPos, 600, 320


    With Me
        .Filter = "[Firma_ID] = " & plDialogSelectedKundenID
        .FilterOn = True
        .AllowAdditions = False
        .Requery
        .Refresh
    End With

    If psDialogAnsprechpartnerKunde = "Bearbeiten" Then
        With Forms("KundenUebersicht").Form.KundeAnsprechpartnerÜbersicht
            Me.SKontaktperson = .Controls("Kontaktperson")
            Me.SKuerzel = .Controls("Kurzzeichen")
            Me.STelefonnummer = .Controls("Telefonnummer")
            Me.SE_Mail = .Controls("EmailAdresse")
        End With
    End If

End Sub

Private Sub Speichern_Click()
On Error GoTo err_Speichern_Click

    Dim lb_DSBearbeiten As Boolean
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from Kunden where [Kontaktperson] = '" & Me.SKontaktperson & "' and [EmailAdresse] = '" & Me.SE_Mail & "' and [FirmaID] = " & Me.Firma_ID & ";")
    If psDialogAnsprechpartnerKunde = "Bearbeiten" Then
        ldb.Execute ("update Kunden set [Kontaktperson] = '" & Me.SKontaktperson & "', [Kurzzeichen] = '" & Me.SKuerzel & _
            "',[Telefonnummer] = '" & Me.STelefonnummer & "',[EmailAdresse] = '" & Me.SE_Mail & _
            "' where [Kunden-ID] = " & plDialogSelectedAPID & ";")
    Else
        Set rs = ldb.OpenRecordset("Select * from Kunden where [Kontaktperson] = '" & Me.SKontaktperson & "' and [EmailAdresse] = '" & Me.SE_Mail & "' and [FirmaID] = " & Me.Firma_ID & ";")
        If rs.RecordCount Then
            MsgBox (getTC_FCFromCode("TC0302") & "'" & Me.SKontaktperson & "'" & getTC_FCFromCode("TC0302"))
        Else
            ldb.Execute ("Insert into Kunden ([Kontaktperson],[FirmaID],[Kurzzeichen],[Telefonnummer],[EmailAdresse], [Aktiv]) " & _
                        "values ('" & Me.SKontaktperson & "', " & Me.Firma_ID & ", '" & Me.SKuerzel & "', '" & Me.STelefonnummer & "', '" & Me.SE_Mail & "', True);")
        End If
    End If

'    rs.MoveLast
'    With Forms("KundenUebersicht").Form("KundeAnsprechpartnerÜbersicht")
'        .SetFocus
'        .RecordsetClone
'        .MoveLast
'        .Bookmark = .Bookmark
'    End With

    FormKoorSave Me
    DoCmd.Close
    Set rs = Nothing

Exit Sub

err_Speichern_Click:
    If Err = 1000 Then
'        MsgBox ("Die Kontaktperosn / Ansprechpartner '" & Me.SKontaktperson & "' gibt es schon für den Kunden. Kontakt kann nicht angelegt werden!") 'TC0202
        MsgBox (getTC_FCFromCode("TC0302") & "'" & Me.SKontaktperson & "'" & getTC_FCFromCode("TC0302"))
    End If
End Sub
