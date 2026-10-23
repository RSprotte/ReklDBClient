Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmbKonfigurieren_Click()

    Select Case Rahmen20
        Case 1
            Call KonfigMsgBox
        Case 2
            Call KonfigInputBox
    End Select

End Sub

Private Sub cmbKonfigurieren_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

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

Private Sub Form_Load()

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub

Private Sub Rahmen20_Click()

    Select Case Rahmen20
        Case 1
            Me.Rahmen9.Visible = True
            Me.picExclamation.Visible = True
            Me.picInformation.Visible = True
            Me.picQuestion.Visible = True
            Me.Bezeichnungsfeld10.Visible = True
            Me.Bezeichnungsfeld13.Visible = True
            Me.Bezeichnungsfeld15.Visible = True
            Me.Bezeichnungsfeld17.Visible = True
            Me.Option12.Visible = True
            Me.Option14.Visible = True
            Me.Option16.Visible = True
        Case 2
            Me.Rahmen9.Visible = False
            Me.picExclamation.Visible = False
            Me.picInformation.Visible = False
            Me.picQuestion.Visible = False
            Me.Bezeichnungsfeld10.Visible = False
            Me.Bezeichnungsfeld13.Visible = False
            Me.Bezeichnungsfeld15.Visible = False
            Me.Bezeichnungsfeld17.Visible = False
            Me.Option12.Visible = False
            Me.Option14.Visible = False
            Me.Option16.Visible = False
    End Select

End Sub

Private Sub KonfigInputBox()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim TextCode As String
    Dim Funktion As String
    Dim InputBoxName As String

    Funktion = "getTC_FCFromCode(""x"")"

    Set ldb = openAccDBConnection("openDB")


    Set rs = ldb.OpenRecordset("Select * from RSP_Lokalisierung where [BasisText] = '" & Me.txtEingabe & "' and  ([TypeOfObject] = 'MsgBox' or [TypeOfObject] = 'InputBox');")
    If rs.RecordCount = 0 Then
            rs.AddNew
            rs("TypeOfObject") = "InputBox"
            rs("OjectID_Name") = "Form"
            rs("TextID") = "IC" & rs("Text-ID")
            TextCode = rs("TextID")
            rs("BasisText") = Me.txtEingabe
            rs("LängeText") = Len(Me.txtEingabe)
            rs.Update
            Set rs = Nothing
    Else
        TextCode = rs("TextID")
    End If
    If Not IsNull(Me.txtMsgName) Then
        Set rs = ldb.OpenRecordset("Select * from RSP_Lokalisierung where [BasisText] = '" & Me.txtMsgName & "' and  ([TypeOfObject] = 'MsgBox' or [TypeOfObject] = 'InputBox');")
        If rs.RecordCount = 0 Then
            rs.AddNew
            rs("TypeOfObject") = "InputBox"
            rs("OjectID_Name") = "Form"
            rs("TextID") = "IC" & rs("Text-ID")
            InputBoxName = rs("TextID")
            rs("BasisText") = Me.txtMsgName
            rs("LängeText") = Len(Me.txtMsgName)
            rs.Update
            Set rs = Nothing
        Else
            InputBoxName = rs("TextID")
        End If
    End If


    If IsNull(Me.txtMsgName) Then
        Me.txtAusgabe = "InputBox " & "(" & Replace(Funktion, "x", TextCode) & ")"
    Else
        Me.txtAusgabe = "InputBox " & "(" & Replace(Funktion, "x", TextCode) & ", " & Replace(Funktion, "x", InputBoxName) & ")"
    End If

    'String in Zwischenablage Kopieren
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    clipboard.SetText Me.txtAusgabe
    clipboard.PutInClipboard

End Sub
Private Sub KonfigMsgBox()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim TextCode As String
    Dim Funktion As String
    Dim MsgBild As String
    Dim MsgName As String

    Funktion = "getTC_FCFromCode(""x"")"

    Set ldb = openAccDBConnection("openDB")

    Select Case Rahmen9
        Case 1
            MsgBild = "vbQuestion"
        Case 2
            MsgBild = "vbInformation"
        Case 3
            MsgBild = "vbExclamation"
    End Select

    Set rs = ldb.OpenRecordset("Select * from RSP_Lokalisierung where [BasisText] = '" & Me.txtEingabe & "' and  ([TypeOfObject] = 'MsgBox' or [TypeOfObject] = 'InputBox');")
    If rs.RecordCount = 0 Then
        rs.AddNew
        rs("TypeOfObject") = "MsgBox"
        rs("OjectID_Name") = "Form"
        rs("TextID") = "TC" & rs("Text-ID")
        TextCode = rs("TextID")
        rs("BasisText") = Me.txtEingabe
        rs("LängeText") = Len(Me.txtEingabe)
        rs.Update
        Set rs = Nothing
    Else
        TextCode = rs("TextID")
    End If

    If Not IsNull(Me.txtMsgName) Then
    Set rs = ldb.OpenRecordset("Select * from RSP_Lokalisierung where [BasisText] = '" & Me.txtMsgName & "' and  ([TypeOfObject] = 'MsgBox' or [TypeOfObject] = 'InputBox');")
        If rs.RecordCount = 0 Then
            rs.AddNew
            rs("TypeOfObject") = "MsgBox"
            rs("OjectID_Name") = "Form"
            rs("TextID") = "TC" & rs("Text-ID")
            MsgName = rs("TextID")
            rs("BasisText") = Me.txtMsgName
            rs("LängeText") = Len(Me.txtMsgName)
            rs.Update
            Set rs = Nothing
        Else
            MsgName = rs("TextID")
        End If
    End If

    If IsNull(Me.txtMsgName) Then
        Me.txtAusgabe = "MsgBox " & Replace(Funktion, "x", TextCode) & ", " & MsgBild
    Else
        Me.txtAusgabe = "MsgBox " & Replace(Funktion, "x", TextCode) & ", " & MsgBild & ", " & Replace(Funktion, "x", MsgName)
    End If

    'String in Zwischenablage Kopieren
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    clipboard.SetText Me.txtAusgabe
    clipboard.PutInClipboard

End Sub
