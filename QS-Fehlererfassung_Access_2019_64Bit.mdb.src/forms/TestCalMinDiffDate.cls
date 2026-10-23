Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub cmbBerechnen_Click()  '20170220-ROS

    Dim liHour, liDays, liMinute As Integer


    Me.txtAnzMin = CalcArbeitszeitDiffDate(Me.txtDatStart, Me.txtDatEnde)
    Me.txtAnzTage = getHourMinuteByAnzMinute(Me.txtAnzMin, liDays, liHour, liMinute)
    Me.txtTage = liDays
    Me.txtGesamtStd = liHour
    Me.txtGesamtMin = liMinute


End Sub
