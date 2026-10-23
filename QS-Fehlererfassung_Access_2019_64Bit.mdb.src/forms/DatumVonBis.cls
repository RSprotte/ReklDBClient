Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ActiveXStr0_Click()
Me.Von = Me.ActiveXStr0.Value
End Sub
Private Sub ActiveXStr1_Click()
Me.Bis = Me.ActiveXStr1.Value
End Sub


Private Sub Befehl5_Click()

Forms("Reklamationen_Übersicht_Konfigurator").Controls("Von") = Me.Von
Forms("Reklamationen_Übersicht_Konfigurator").Controls("Bis") = Me.Bis
DoCmd.Close

End Sub

Private Sub Form_Load()

If Not (Forms("Reklamationen_Übersicht_Konfigurator").Von = "") Then
    Me.ActiveXStr0 = Forms("Reklamationen_Übersicht_Konfigurator").Von
    Me.Von = Forms("Reklamationen_Übersicht_Konfigurator").Von
End If
If Not (Forms("Reklamationen_Übersicht_Konfigurator").Bis = "") Then
    Me.ActiveXStr1 = Forms("Reklamationen_Übersicht_Konfigurator").Bis
    Me.Bis = Forms("Reklamationen_Übersicht_Konfigurator").Bis
End If

End Sub
