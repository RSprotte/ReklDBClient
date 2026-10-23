Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Form_Close()
    Dim ldb As DAO.Database

'    FormKoorSave Me
    Set ldb = openAccDBConnection("openDB")
    ldb.Execute ("delete * from tblTempAnspLief;")

End Sub

Private Sub Form_DblClick(Cancel As Integer)

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim lAP_ID As Long

'20191015-ROS#############################################################################################################
'Wenn der im AVISTA ausgewählte Ansprechpartner Lieferant nicht in der Tabelle "LieferantAnsprechpartner" angelet ist,
'wird dieser neu angelegt. In jedem Fall werden die Daten von AVISTA aktuallisiert
    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from [LieferantAnsprechpartner] where [AnsprechpartnerID] = " & Me.[pa_id] & ";")
    If rs.RecordCount Then
        rs.Edit
    Else
        rs.AddNew
        rs![AnsprechpartnerID] = Me.[pa_id]
        rs![LieferantID] = Forms(sCurrentFehlererfassung).Controls("Mastertab#Lieferanten-Code 1")
    End If
    lAP_ID = rs![AnsprechpartnerID]
    rs![Name] = Me.[pa_org_name1]
    rs![Vorname] = Me.[pa_pers_vorname]
    rs![Email] = Me.[adb_email]
    rs![Telefon] = Me.[adb_tel1]
    rs![Position] = Me.[asp_position]
    rs.Update
    rs.Close
    Set rs = Nothing
'##########################################################################################################################
    Forms(sCurrentFehlererfassung).Controls("Mastertab#LiefAnsprechp-ID#3") = Me.[pa_id]
    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Forms(sCurrentFehlererfassung), Forms(sCurrentFehlererfassung).Controls("Mastertab#LiefAnsprechp-ID#3"))

    Set lrs = Nothing


Me.SetFocus
DoCmd.Close acForm, "LieferantAnsprechpartnerTemp"

db.Execute ("delete * from tblTempAnspLief;")

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
    FormOpenKoorMousPointer Me, XPos, YPos, 1070, 290

End Sub
