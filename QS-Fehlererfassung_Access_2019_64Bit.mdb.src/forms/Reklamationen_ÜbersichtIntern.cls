Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

#If VBA7 Or Win64 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
      (ByVal hwnd As Long, ByVal lpOperation As String, _
      ByVal lpFile As String, ByVal lpParameters As String, _
      ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
      (ByVal hwnd As Long, ByVal lpOperation As String, _
      ByVal lpFile As String, ByVal lpParameters As String, _
      ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If


Public Sub setValue(ByVal CntDate As Date)

    Me.ActiveControl = CntDate

End Sub
Private Sub Befehl60_Click()

If IsLoaded("Reklamationen_Übersicht_Konfigurator") Then
    Forms("Reklamationen_Übersicht_Konfigurator").SetFocus
Else
    DoCmd.OpenForm "Reklamationen_Übersicht_Konfigurator"
    Forms("Reklamationen_Übersicht_Konfigurator").Bis = Me.Bis
    Forms("Reklamationen_Übersicht_Konfigurator").Von = Me.Von
    Forms("Reklamationen_Übersicht_Konfigurator").Fehlerquelle = 1
    Forms("Reklamationen_Übersicht_Konfigurator").Fehlerquelle.Enabled = False
    Forms("Reklamationen_Übersicht_Konfigurator").optKundenFilter = 0   '20160318-ROS  eingefügt
End If

    Me.cboFilterForm = ""

End Sub



Private Sub Bis_Click()
   Dim frm As Form

    DoCmd.OpenForm ("frmCalendar")
    Set frm = Forms("frmCalendar")
    Call frm.setParentForm(Me)
End Sub

Private Sub cboFilterForm_Click()  '20180905-ROS
    Dim ldb As DAO.Database
    Dim Date1 As String
    Dim Schaltjahr As String
    Dim strFormFilter As String
    Dim rsFormFilter As DAO.Recordset
    Dim ld_date As String


    Set ldb = openAccDBConnection("openDB")
    Set rsFormFilter = ldb.OpenRecordset("Select * from tblFilterForm where [FilterBezeichnung] = '" & Me.cboFilterForm & "' And [FormName]= '" & Me.Name & "';")

    strFormFilter = rsFormFilter("FilterParameter")

    If rsFormFilter("Dynamisch") Then
        strFormFilter = Replace(rsFormFilter("FilterParameter"), "#Von#", "#" & format(Me.Von, "mm\/dd\/yyyy") & "#")
        strFormFilter = Replace(strFormFilter, "#Bis#", "#" & format(Me.Bis, "mm\/dd\/yyyy") & "#")
    Else
        strFormFilter = rsFormFilter("FilterParameter")
        Me.Von = rsFormFilter("DatumVon")
        Me.Bis = rsFormFilter("DatumBis")
    End If
    DoCmd.OpenForm "Reklamationen_ÜbersichtIntern", acNormal
    Forms("Reklamationen_ÜbersichtIntern").Filter = strFormFilter ' & sAuftrag
    Forms("Reklamationen_ÜbersichtIntern").FilterOn = True
    Forms("Reklamationen_ÜbersichtIntern").FilterView = rsFormFilter("FilterParameter")
    Me.Requery
    Me.Refresh


    'DoCmd.Maximize

    Set rsFormFilter = Nothing

'    20180913-ROS
    Dim QS As DAO.QueryDef
    Set QS = ldb.QueryDefs("RSP_Ausw_Basisdaten_1_Export")
    If InStr(QS.sql, "Where") Then
'            QS.SQL = Left(QS.SQL, InStr(QS.SQL, "Where") + 5) & Forms(sÜbersicht).Filter = sFilter & ";"
        QS.sql = Left(QS.sql, InStr(QS.sql, "Where") + 5) & strFormFilter & ";"
    Else
'            QS.SQL = Left(QS.SQL, InStr(QS.SQL, ";") - 1) & " Where " & Forms(sÜbersicht).Filter = sFilter & ";"
        QS.sql = Left(QS.sql, InStr(QS.sql, ";") - 1) & " Where " & strFormFilter & ";"
    End If

End Sub

Private Sub cboPersFilterSpeichern_Click() '20180905-ROS
    Dim ldb As DAO.Database
    Dim strFilterBezeichnung As String
    Dim strFilterparameter As String

    strFilterparameter = Me.Filter
    'strFilterparameter = Me.FilterView.Value
    strFilterparameter = Replace(strFilterparameter, "'", "''")

    Set ldb = openAccDBConnection("openDB")
'    strFilterBezeichnung = InputBox("Bitte Filterbezeichnung eingeben")
'    strFilterBezeichnung = InputBox(getTC_FCFromCode("IC3845"), getTC_FCFromCode("IC3846"))
    strFilterBezeichnung = getValueFromInputBox(getTC_FCFromCode("IC3845"), getTC_FCFromCode("IC3846"))

    If strFilterBezeichnung = "" Then Exit Sub

    li_retMsgBox = MsgBox(getTC_FCFromCode("TC4209"), vbYesNo + vbDefaultButton2)

    If li_retMsgBox = 6 Then
        ldb.Execute ("Insert into tblFilterForm ([SachbearbeiterID], [FormName], [FilterParameter], [FilterBezeichnung], [Aktiv], [Dynamisch], [DatumVon], [DatumBis])" & _
                    "values(" & get_SachbearbeiterID() & ", '" & Me.Name & "' , '" & strFilterparameter & "', '" & strFilterBezeichnung & "' , True, " & _
                    "False, #" & format(Me.Von, "mm\/dd\/yyyy") & "#, #" & format(Me.Bis, "mm\/dd\/yyyy") & "#);")
    Else

        ldb.Execute ("Insert into tblFilterForm ([SachbearbeiterID], [FormName], [FilterParameter], [FilterBezeichnung], [Aktiv], [Dynamisch], [DatumVon], [DatumBis])" & _
                    "values(" & get_SachbearbeiterID() & ", '" & Me.Name & "' , '" & Replace(Me.FilterView, "'", "''") & "', '" & strFilterBezeichnung & "' , True, True, '','');")
    End If

    Me.Requery
    Me.Refresh

End Sub

Private Sub ExportExcel_Click()

    Dim rst As DAO.Recordset
    Dim customQuery As String
    Dim cnt As Integer

    Dim appExcel As Excel.Application
    Dim wbk As Excel.Workbook
    Dim wks As Excel.Worksheet
    Dim rng As Excel.Range

    Dim FileName As String

    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "RSP_Ausw_Basisdaten_1_Export", format(Now(), "yyyy/mm/dd_hh.mm") & "_Übersicht_Reklamationen.xls", True, _
                 "Export_Übersicht"

    ShellExecute 0, "open", format(Now(), "yyyy/mm/dd_hh.mm") & "_Übersicht_Reklamationen.xls", "", "", 1

'    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, _
'                 "RSP_Ausw_Basisdaten_1_Export", Date & "_Übersicht_Reklamationen.xls", True, _
'                 "Export_Übersicht"
'    ShellExecute 0, "open", "C:\Users\ralf sprotte\Documents\" & Date & "_Übersicht_Reklamationen.xls", "", "", 1
    'Set appExcel = New Excel.Application
    'Set wbk = appExcel.Workbooks.Open("C:\Users\ralf sprotte\Documents\" & Date & "_Übersicht_Reklamationen.xls")

    Set wks = Nothing
    Set wkb = Nothing


End Sub

Private Sub FilterLoeschen_Click()

'    Dim Date1 As String
'    Dim Schaltjahr As String
'
'    Date1 = "29.02."
'    Schaltjahr = Format(Date, "dd.mm.")
'    If Schaltjahr = Date1 Then
'        sVon = Format(Date - 1, "dd.mm") & "." & LTrim(Str(Val(Format(Date, "yyyy")) - 1))
'    Else
'        sVon = Format(Date, "dd.mm") & "." & LTrim(Str(Val(Format(Date, "yyyy")) - 1))
'    End If
'
'    Me.Von = sVon
'    Me.Bis = Date
'
'    Me.Filter = "[intern] = True AND [ErfasserIntern] Is Not Null"
'    Me.FilterView = Me.Filter
'    Me.cboFilterForm = ""
'    Me.Requery
'    Me.Refresh

    Dim Date1 As String
    Dim Schaltjahr As String

'    Date1 = "29.02."
'    Schaltjahr = Format(Date, "dd.mm.")
'    If Schaltjahr = Date1 Then
'        sVon = Format(Date - 1, "dd.mm") & "." & LTrim(Str(Val(Format(Date, "yyyy")) - 1))
'    Else
'        sVon = Format(Date, "dd.mm") & "." & LTrim(Str(Val(Format(Date, "yyyy")) - 1))
'    End If

    Date1 = "02/29/"
    Schaltjahr = format(Date, "mm\/dd\/")
    If Schaltjahr = Date1 Then
        sVon = format(Date - 1, "mm\/dd\/") & LTrim(Str(Val(format(Date, "yyyy")) - 1))
    Else
        sVon = format(Date, "mm\/dd\/") & LTrim(Str(Val(format(Date, "yyyy")) - 1))
    End If
    sBis = Date
    sAuftrag = " AND ([Auftragsdatum] between #" & format(sVon, "mm\/dd\/yyyy") & "# and #" & _
                                                format(sBis, "mm\/dd\/yyyy") & "#)"


    Me.Von = sVon
    Me.Bis = format(Date, "mm\/dd\/yyyy")

'    Me.Filter = "([ErfasserIntern] Is Null)"
    Forms("Reklamationen_Übersichtintern").Filter = "([ErfasserIntern] Is Not Null" & sAuftrag & ")"
    Me.FilterView = Me.Filter
    Me.cboFilterForm = ""
    Me.Requery
    Me.Refresh
End Sub



Private Sub FilterView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CursorWechseln(IDC_HAND)
End Sub

Private Sub Form_Activate()
'DoCmd.Maximize
End Sub


Private Sub Form_Close()

On Error GoTo err_Form_Close

    DoCmd.Close acForm, "Reklamationen_Übersicht_Konfigurator"

err_Form_Close:
    Exit Sub

End Sub

Private Sub Form_Current()
'DoCmd.Maximize
End Sub

Private Sub Form_DblClick(Cancel As Integer)

    Dim lsForm As String

''    If Not IsNull(Me.ErfasserIntern) Then
''        lsForm = "FehlererfassungIntern"
''    Else
'        lsForm = "Fehlererfassung - Register"
''    End If
'
'    p_ReklNr = Me.ID
'    '20151002-RSP Filter direkt setzen
''    If GetStateVorgang(get_p_ReklNr()) = "Abgeschlossen" And lsForm = "Fehlererfassung - Register" Then
'    If GetStateVorgang(get_p_ReklNr()) = "Abgeschlossen" Then       '20190704-ROS
'        DoCmd.OpenForm lsForm, acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormReadOnly
'    Else
'        If proofAndLockDataset(get_p_ReklNr()) = "" Then
'            ReklBearbeiten = True
'            InsertNewReklMemoDS (get_p_ReklNr())
''            DoCmd.OpenForm lsForm, acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormEdit
'            DoCmd.OpenForm lsForm
'        Else
'            ReklBearbeiten = False
''            DoCmd.OpenForm lsForm, acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormReadOnly
'            DoCmd.OpenForm lsForm
'        End If
'    End If
'
'    'set_FehlererfassungEnviroment       '20190407-ROS
'
'    With Forms(lsForm)
'        'DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , , acFormEdit
''        .Filter = "[Mastertab].[ID] = " & get_p_ReklNr ()
''        .FilterOn = True
''        .Requery
''        .Refresh
'        .AllowAdditions = False
'        'DoCmd.Maximize
'        'Sichern des aktuellen Vorgang-Status
'        If IsNull(Forms(lsForm).Controls("Mastertab#Kategorie-Status")) Then
'            p_VorgangStatus = ""
'        Else
'            p_VorgangStatus = Forms(lsForm).Controls("Mastertab#Kategorie-Status")
'        End If
'    End With



    p_ReklNr = Me.ID
    NewReklamationE "Bearbeiten", get_p_ReklNr()

    Forms("Fehlererfassung - Register").SetFocus


end2_Form_DblClick:

End Sub

Private Sub Form_Load()

    Select Case Application.CurrentProject.Name
        Case "QS-FehlererfassungIntern.mdb"
            Me.Filter = "[intern] = True AND [ErfasserIntern] Is Not Null"
            Me.FilterOn = True
        Case Else
            Me.Filter = ""
    End Select
    Me.FilterView = Me.Filter
    'DoCmd.Maximize

    Dim objMouseOverForm As clsMouseoverForm
    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)


End Sub


Private Sub Form_Open(Cancel As Integer)

    Select Case Application.CurrentProject.Name
        Case "QS-FehlererfassungIntern.mdb"
            Me.Filter = "[intern] = True AND [ErfasserIntern] Is Not Null"
            Me.FilterOn = True
        Case Else
            Me.Filter = ""
    End Select
    Me.FilterView = Me.Filter

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub

Private Sub Form_Timer()

    UpdateForm_Click

End Sub

Private Sub IntervallOFF_Click()

Me.TimerInterval = 0
Me.UpdateForm.SetFocus
Me.IntervallOFF.Enabled = False

End Sub

Private Sub Kommentar_Click()

    p_ReklMemoQuelle = "Reklamationen_Übersicht"
    DoCmd.OpenForm "ReklMemo", acNormal

End Sub

Private Sub Objekt_schließen___Formular_Click()
    FormKoorSave Me
    DoCmd.Close
End Sub

Private Sub UpdateForm_Click()

    Call OpenForm_Reklamationen_Übersicht
'Me.Requery
'Me.Refresh
End Sub

Private Sub Von_Click()
   Dim frm As Form

    DoCmd.OpenForm ("frmCalendar")
    Set frm = Forms("frmCalendar")
    Call frm.setParentForm(Me)
End Sub
