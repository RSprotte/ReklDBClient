Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmbVerknüpfungLoeschen_Click()      '20190513-ROS

    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")

    'Declaration Avista DB provider
    Dim adbprov As clsAvistaProvider

    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider

    'Verknüfung der WS zwischen Avista und Access in ReklaDB löschen
    ldb.Execute ("Update tblAccessWsID_AvistaWsID set [AccessWsID] = 0 where [AvistaWsID] = " & Me.txtAvistaWsID & " ;")
    ldb.Execute ("Update LQWorkstation set [WSAvistaID] = 0 where [ID] = " & Me.txtWSID & ";")

    'Verknüfung der WS zwischen Avista und Access in Avista löschen -> Fehlerbilder
    adbprov.CommandExecuteAvista ("Delete * FROM zlqpp_fe_ws_fehler_info where [lqpp_ws_id] = " & Me.txtAvistaWsID & " ;")


    'Verknüfung der WS zwischen Avista und Access in Avista löschen -> WS zu Bereich
    adbprov.CommandExecuteAvista ("Delete * FROM zlqpp_fe_ws_bereich where [lqpp_ws_id] = " & Me.txtAvistaWsID & " ;")

    Forms("frmAccessWsID_AvistaWSID").Requery
    Me.Refresh
    Me.Requery
    Set adbprov = Nothing


End Sub

Private Sub cmbVerknüpfungLoeschen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub
