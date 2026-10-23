Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmbDatenÜbernehmen_Click()
    Dim MailAn As String
    Dim MailCC As String
    Dim MailBCC As String
    Dim s As String

    s = get_EMailAdressen2(MailAn, MailCC, MailBCC)

    ps_MailAn = MailAn
    ps_MailCC = MailCC
    ps_MailBCC = MailBCC

    DoCmd.Close

End Sub
Private Function get_EMailAdressen2(ByRef MailAn, Optional ByRef MailCC, Optional ByRef MailBCC) As String       '20181217-ROS

        If IsNull(Me.Controls("LQWorkstationID")) Or (Me.Controls("LQWorkstationID")) = (-1) Then

                If IsNull(Forms("frmLagerPrüfung")![LQBereichID]) Or (Forms("frmLagerPrüfung")![LQBereichID]) = (-1) Then
                    MailAn = DLookup("MailTOLQEinheit", "LQEinheit", "[ID] = " & Me.Controls("LQEinheitID") & "")
                    MailCC = DLookup("MailCCLQEinheit", "LQEinheit", "[ID] = " & Me.Controls("LQEinheitID") & "")
                    MailBCC = DLookup("MailEskalationLQEinheit", "LQEinheit", "[ID] = " & Me.Controls("LQEinheitID") & "")
                Else
                    MailAn = DLookup("MailTOLQBereich", "LQBereich", "[ID] = " & Me.Controls("LQBereichID") & "")
                    MailCC = DLookup("MailCCLQBereich", "LQBereich", "[ID] = " & Me.Controls("LQBereichID") & "")
                    MailBCC = DLookup("MailEskalationLQBereich", "LQBereich", "[ID] = " & Me.Controls("LQBereichID") & "")
                End If
        Else
            MailAn = Nz(DLookup("MailTOLQWS", "LQWorkstation", "[ID] = " & Me.Controls("LQWorkstationID") & ""))
            MailCC = Nz(DLookup("MailCCLQWS", "LQWorkstation", "[ID] = " & Me.Controls("LQWorkstationID") & ""))
            MailBCC = Nz(DLookup("MailEskalationLQWS", "LQWorkstation", "[ID] = " & Me.Controls("LQWorkstationID") & ""))
        End If

        If IsNull(MailAn) Or ((MailAn) = "") Then
            MailAn = DLookup("MailTOLQEinheit", "LQEinheit", "[ID] = " & Me.Controls("LQEinheitID") & "")
            MailCC = DLookup("MailCCLQEinheit", "LQEinheit", "[ID] = " & Me.Controls("LQEinheitID") & "")
            MailBCC = DLookup("MailEskalationLQEinheit", "LQEinheit", "[ID] = " & Me.Controls("LQEinheitID") & "")
        End If

    get_EMailAdressen2 = ""

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyDown KeyCode, Shift, Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyUp KeyCode, Shift
End Sub

Private Sub Form_Load()
    Me.LQEinheitID = Null
    Me.LQBereichID = Null
    Me.LQWorkstationID = Null
    ps_MailAn = ""
    ps_MailCC = ""
    ps_MailBCC = ""

    Dim objMouseOverForm As clsMouseoverForm

    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With


End Sub

Private Sub Form_Open(Cancel As Integer)

    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long

    lgPosition = getPointerPos(XPos, YPos)
    FormOpenKoorMousPointer Me, XPos, YPos, 460, 190, False

End Sub

Private Sub LQBereichID_AfterUpdate()
    Me.LQWorkstationID = -1
    Me.Refresh
    Me.Requery
End Sub

Private Sub LQEinheitID_AfterUpdate()
    Me.LQBereichID = -1
    Me.LQWorkstationID = -1
    Me.Refresh
    Me.Requery
End Sub

Private Sub LQWorkstationID_AfterUpdate()
    Me.Refresh
    Me.Requery
End Sub
