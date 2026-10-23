Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Open(Cancel As Integer)    'Kürzel ROS

    DoCmd.ShowToolbar "Ribbon", acToolbarNo
    If Environ("USERNAME") = "Roberto Santarelli" Then
        get_pKuerzel "ROS", True
        DoCmd.Close
        DoCmd.OpenForm ("TB:autostart")
    End If
    If Environ("USERNAME") = "Ralf Sprotte" Then
        get_pKuerzel "RSP", True
        DoCmd.Close
        DoCmd.OpenForm ("TB:autostart")
    End If

End Sub

Private Sub txtEingabeKz_AfterUpdate()

    get_pKuerzel Me.txtEingabeKz, True
    DoCmd.Close
    DoCmd.OpenForm ("TB:autostart")

End Sub
