Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


#If VBA7 Or Win64 Then
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
                                                        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
                                                        ByVal cy As Long, ByVal wFlags As Long) As Long

    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

#Else
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
                                                        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
                                                        ByVal cy As Long, ByVal wFlags As Long) As Long

    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If

Private Const WM_DESTROY = &H2
Private Const WM_CLOSE = &H10
Private Const SWP_NOZORDER As Long = &H4
Private Const SW_HIDE = 0       ' Fenster ausblenden
Private Const SW_NORMAL = 1     ' Fenster wiederherstellen bzw. einblenden
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOW = 5


Private MousePosition As POINTAPI
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private m_Table As String
Private m_Year As Long
Private m_Month As Long
Private m_SelectedDate As Date
Private m_Backcolor As Long
Private m_Fontname As String
Private m_FontSize As Long
Private m_Date As Date
Private m_parentForm As Form

Public Event DateSelected(ByVal Value As Date)
Public Event YearChanged(ByVal NewYear As Long)
Public Event MonthChanged(ByVal NewMonth As Long)
Public Event NewHeight(ByVal H As Long)

Public Sub setParentForm(ByRef ffrm As Form)

    Set m_parentForm = ffrm

End Sub

Private Sub Form_Load()
    Dim i As Long

    m_Month = VBA.Month(Date)
    m_Year = VBA.Year(Date)
    m_Table = "tblCalendar"

    For i = 1950 To 2030
        Me.cbYear.AddItem i
    Next i
    Me.cbYear.Value = m_Year

    For i = 1 To 12
        Me.cbMonth.AddItem CStr(i) & ";" & _
            format(DateSerial(m_Year, i, 1), "mmmm") & ";"
    Next i
    Me.cbMonth.Value = m_Month

    m_Fontname = Me.txtD1.FontName
    m_FontSize = Me.txtD1.FontSize

    FillCalendar format(Date, "dd/mm/yyyy")

    On Error Resume Next
    m_Backcolor = Me.Parent.Section(acDetail).BackColor
    If Err.Number <> 0 Then Exit Sub
    Me.Section(acDetail).BackColor = m_Backcolor
    Me.Section(acHeader).BackColor = m_Backcolor
    Me.Section(acFooter).BackColor = m_Backcolor
End Sub

Property Get Table() As String
    Table = m_Table
End Property
Property Let Table(ByVal Value As String)
    m_Table = Value
    FillCalendar
End Property

Property Get Year() As Long
    Year = m_Year
End Property
Property Let Year(ByVal Value As Long)
    m_Year = Value
    Me.cbYear.Value = Value
    FillCalendar
End Property

Property Get Month() As Long
    Month = m_Month
End Property
Property Let Month(ByVal Value As Long)
    m_Month = Value
    Me.cbMonth.Value = Value
    FillCalendar
End Property

Property Get FontName() As String
    FontName = m_Fontname
End Property
Property Let FontName(ByVal Value As String)
    m_Fontname = Value
    FormatControls
End Property

Property Get FontSize() As Long
    FontSize = m_FontSize
End Property
Property Let FontSize(ByVal Value As Long)
    m_FontSize = Value
    FormatControls
End Property

Property Get BackColor() As Long
    BackColor = m_Backcolor
End Property
Property Let BackColor(ByVal Value As Long)
    m_Backcolor = Value
    Me.Section(acDetail).BackColor = Value
    Me.Section(acFooter).BackColor = Value
    Me.Section(acHeader).BackColor = Value
End Property

Property Get SelectedDate() As Date
    SelectedDate = m_SelectedDate
End Property

Public Sub SelectDate(ByVal D As Date)
    Me.Year = VBA.Year(D)
    Me.Month = VBA.Month(D)
    m_SelectedDate = D

    FillCalendar D
    Me.LblDate.Caption = format(m_SelectedDate, "dddd\, dd.mm.yyyy")
End Sub

Private Sub cbMonth_AfterUpdate()
    m_Month = Me.cbMonth.Value
    FillCalendar
    RaiseEvent MonthChanged(m_Month)
End Sub

Private Sub cbYear_AfterUpdate()
    m_Year = Me.cbYear.Value
    FillCalendar
    RaiseEvent YearChanged(m_Year)
End Sub

Private Sub cmdMonthBack_Click()
    m_Month = m_Month - 1
    If m_Month < 1 Then m_Month = 1
    Me.cbMonth.Value = m_Month
    FillCalendar
    RaiseEvent MonthChanged(m_Month)
End Sub

Private Sub cmdMonthNext_Click()
    m_Month = m_Month + 1
    If m_Month > 12 Then m_Month = 12
    Me.cbMonth.Value = m_Month
    FillCalendar
    RaiseEvent MonthChanged(m_Month)
End Sub

Private Sub cmdYearBack_Click()
    m_Year = m_Year - 1
    Me.cbYear.Value = m_Year
    FillCalendar
    RaiseEvent YearChanged(m_Year)
End Sub

Private Sub cmdYearNext_Click()
    m_Year = m_Year + 1
    Me.cbYear.Value = m_Year
    FillCalendar
    RaiseEvent YearChanged(m_Year)
End Sub

Private Sub FillCalendar(Optional SelDate As Variant)
    Dim rs As DAO.Recordset
    Dim i As Long
    Dim j As Long
    Dim X As Long
    Dim Y As Long
    Dim n As Long
    Dim StartDate As Date
    Dim H As Long

    If Len(m_Table) = 0 Then Exit Sub
    Me.RecordSource = m_Table

    n = Weekday(DateSerial(m_Year, m_Month, 1), vbMonday)
    StartDate = DateSerial(m_Year, m_Month, 1) - n

    CurrentDb.Execute "DELETE FROM " & m_Table
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & m_Table, dbOpenDynaset)
    For j = 0 To 5
        rs.AddNew
        For i = 0 To 6
            StartDate = StartDate + 1
            If Not IsMissing(SelDate) Then
                If StartDate = SelDate Then X = i + 1: Y = j
            End If
            n = VBA.Month(StartDate)
            rs.Fields("D" & CStr(1 + i)).Value = Day(StartDate) * IIf(n <> m_Month, -1, 1)
        Next i
        rs.Update
        If n <> m_Month Then Exit For
'        If (m_Month = 12) And (n = 1) Then Exit For
    Next j

    rs.Close
    Set rs = Nothing
    Me.Requery

    Me.LblDate.Caption = format(Date, "dddd") & " " & format(Date, "dd/mm/yyyy")

    If Not IsMissing(SelDate) Then
        Dim ctl As Access.TextBox
        Me.Recordset.AbsolutePosition = Y
        Set ctl = Me.Controls("txtD" & CStr(X))
        ctl.SetFocus
        ctl.SelStart = 0: ctl.SelLength = 2
    Else
        m_SelectedDate = -1
    End If

    H = Me.Section(acHeader).Height + Me.Section(acFooter).Height
    H = H + (Me.Recordset.RecordCount * 315)
    Me.InsideHeight = H
'    RaiseEvent NewHeight(H)
End Sub

Private Sub FormatControls()
    Dim ctl As Access.Control

    For Each ctl In Me.Controls
        Select Case ctl.ControlType
        Case acTextBox, acLabel, acCheckBox
            ctl.FontName = m_Fontname
            ctl.FontSize = m_FontSize
            ctl.BackColor = m_Backcolor
        End Select
    Next ctl
End Sub

Private Sub Form_Open(Cancel As Integer)

    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long

    lgPosition = getPointerPos(XPos, YPos)
    FormOpenKoorMousPointer Me, XPos, YPos, 210, 228, False

End Sub

Private Sub txtD1_Click()
    fuClick
    DoCmd.Close acForm, "frmCalendar"
    Call m_parentForm.setValue(m_Date)
End Sub
Private Sub txtD2_Click()
    fuClick
    DoCmd.Close acForm, "frmCalendar"
    Call m_parentForm.setValue(m_Date)
End Sub
Private Sub txtD3_Click()
    fuClick
    DoCmd.Close acForm, "frmCalendar"
    Call m_parentForm.setValue(m_Date)
End Sub
Private Sub txtD4_Click()
    fuClick
    DoCmd.Close acForm, "frmCalendar"
    Call m_parentForm.setValue(m_Date)
End Sub
Private Sub txtD5_Click()
    fuClick
    DoCmd.Close acForm, "frmCalendar"
    Call m_parentForm.setValue(m_Date)
End Sub
Private Sub txtD6_Click()
    fuClick
    DoCmd.Close acForm, "frmCalendar"
    Call m_parentForm.setValue(m_Date)
End Sub
Private Sub txtD7_Click()
    fuClick
    DoCmd.Close acForm, "frmCalendar"
    Call m_parentForm.setValue(m_Date)
End Sub

Private Sub fuClick()
    Dim n As Long
    Dim nAdd As Long
    Dim Location As String

    n = Me.ActiveControl.Value
    If (Me.CurrentRecord < 2) And (n > 20) Then nAdd = -1
    If (Me.CurrentRecord > 5) And (n < 8) Then nAdd = 1
    m_SelectedDate = DateSerial(m_Year, m_Month + nAdd, n)
    Me.LblDate.Caption = format(m_SelectedDate, "dddd\, dd.mm.yyyy")
    RaiseEvent DateSelected(m_SelectedDate)

    Location = DLookup("Value", "KonfigurationReklDBLocal", "Konfiguration = 'Location'")
    If Location = "English" Then
        m_Date = format(m_SelectedDate, "mm\/dd\/yyyy")
    Else
        m_Date = format(m_SelectedDate, "dd\/mm\/yyyy")
    End If


End Sub
