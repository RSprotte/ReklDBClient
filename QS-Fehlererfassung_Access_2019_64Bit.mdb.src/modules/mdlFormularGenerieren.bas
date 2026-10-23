Attribute VB_Name = "mdlFormularGenerieren"
Option Compare Database
Option Explicit

Public Sub FormularErstellen(fstrFormname As String)

    Dim lfrm As Form
    Dim lstrFormTemp As String
    Set lfrm = CreateForm
    lstrFormTemp = lfrm.Name
    DoCmd.Restore
    DoCmd.Save acForm, lstrFormTemp
    DoCmd.Close acForm, lstrFormTemp
    DoCmd.Rename fstrFormname, acForm, lstrFormTemp

End Sub

Public Sub FormularEinrichten(ByVal fFormname As String, fTablname As String)

    Dim strFormname As String
    Dim strRecordSource As String
    Dim strName As String
    Dim strTabellenName As String

    'Eingabe des Formularnamen muss gleichen Namen haben wie Tabelle
'    strName = InputBox("Formularname", "Bitte Formularname eingeben.")
    strName = fFormname

'    strTabellenName = "tbl" & strName
'    strTabellenName = InputBox("Quelle", "Bitte Quelle eingeben.")
    strTabellenName = fTablname

    'Prüfen ob Tabelle oder Query zum erstellen des Formulars vorhanden ist
    If QueryExistsADO(strTabellenName) = True Then
        strFormname = "frm" & strName
        strRecordSource = "SELECT * FROM " & strTabellenName & ""
    ElseIf TableExistsADO(strTabellenName) = True Then
        strFormname = "frm" & strName
        strRecordSource = "SELECT * FROM " & strTabellenName & ""
    Else
        MsgBox "Tabelle zu erstellen des Formular nicht vorhanden oder Tabellenname wurde falsch geschrieben"
        Exit Sub
    End If

    On Error Resume Next
    DoCmd.DeleteObject acForm, strFormname
    On Error GoTo 0
    FormularErstellen strFormname
    FormularansichtEinrichten strFormname, strRecordSource
    DoCmd.OpenForm strFormname, acFormDS, , , acFormReadOnly
    DoCmd.Restore
    DoCmd.MoveSize 6500, 3500, 10000, 7500


End Sub

Public Sub FormularansichtEinrichten(fstrFormname As String, fstrRecordsource As String)

    Dim ldb As DAO.Database
    Dim rst As DAO.Recordset
    Dim fld As DAO.Field
    Dim frm As Form
    Dim intControlType As AcControlType
    Dim ctl As Control
    Dim lngLabelX As Long
    Dim lngLabelY As Long
    Dim lngControlX As Long
    Dim lngControlY As Long

    Dim sngTop As Single
    Dim intFieldsize As Integer
    DoCmd.OpenForm fstrFormname, acDesign, , , , acHidden
'    DoCmd.OpenForm fstrFormname, acDesign
    Set frm = Forms(fstrFormname)
    Set ldb = CurrentDb
    Set rst = ldb.OpenRecordset(fstrRecordsource, dbOpenDynaset)
    Call GetSizeAllLabels(rst, frm.DefaultControl(acTextBox), lngLabelX, lngLabelY)
    sngTop = 100

    For Each fld In rst.Fields
        intControlType = acTextBox

        On Error Resume Next
        intControlType = fld.Properties("DisplayControl")
        On Error GoTo 0
        Set ctl = Application.CreateControl(fstrFormname, intControlType, acDetail, , , lngLabelX + 1000, sngTop)
        Call GetMaximalSizeControl(db, fstrRecordsource, ctl, fld, lngControlX, lngControlY)
        intFieldsize = 0
        On Error Resume Next
        intFieldsize = fld.FieldSize
        If Err.Number = 0 Then 'Memo
            lngControlX = lngControlX / 4
            lngControlY = lngControlY * 4
        End If
'        On Error GoTo 0
'        ctl.Width = lngControlX
'        ctl.Height = lngControlY
        ctl.D
        'Feste Felstergröße vergeben
        ctl.Width = 3000
        ctl.Height = 350
        'Wudre Auskommentiert da nur ungebundene Felder erstellt werden sollen
'        ctl.ControlSource = fld.SourceField
        Select Case intControlType
            Case acComboBox
                ctl.RowSource = fld.Properties("RowSource")
                ctl.ColumnCount = fld.Properties("ColumnCount")
                ctl.ColumnWidths = fld.Properties("ColumnWidths")
                ctl.Name = "cbo" & fld.Name
            Case acListBox
                ctl.RowSource = fld.Properties("RowSource")
                ctl.ColumnCount = fld.Properties("ColumnCount")
                ctl.ColumnWidths = fld.Properties("ColumnWidths")
                ctl.Name = "lst" & fld.Name
            Case acCheckBox
                ctl.Name = "chk" & fld.Name
            Case acTextBox
                ctl.Name = "txt" & fld.Name
                ctl.ControlSource = fld.Name
        End Select

        Call BuildLabel(frm, fld, ctl, intControlType, lngLabelX, lngLabelY, sngTop)
        sngTop = sngTop + ctl.Height + 200 'für Abstand zwischen den Feldern Vertikal

        ctl.TextAlign = 1
        ctl.FontName = "Arial"
        ctl.FontSize = 12

    Next fld
    'Wudre Auskommentiert da nur ungebundene Felder erstellt werden sollen
'    frm.RecordSource = fstrRecordsource

    frm.RecordSource = fstrRecordsource '"SELECT * from TBL_TEMP"

'    frm.DefaultView = acDefViewSingle
    frm.PopUp = True
    frm.NavigationButtons = True
    frm.RecordSelectors = False
    frm.ScrollBars = 2
    frm.RecordSelectors = True
    frm.DefaultView = 1
    DoCmd.Close acForm, frm.Name, acSaveYes

End Sub

Public Function BuildLabel(frm As Form, fld As DAO.Field, ctl As Control, intControlType As AcControlType, lngLabelX As Long, lngLabelY As Long, sngTop As Single) As Label

    Dim lbl As Control
    Dim strCaption As String
    Set lbl = Application.CreateControl(frm.Name, acLabel, acDetail, ctl.Name, , 100, sngTop, lngLabelX + 800, lngLabelY)
    strCaption = fld.Name
    On Error Resume Next
    strCaption = fld.Properties("Caption")
    On Error GoTo 0
    If frm.DefaultControl(intControlType).AddColon Then
        strCaption = strCaption & ":"
    End If

    lbl.Caption = strCaption
    lbl.Name = "lbl" & fld.Name
    lbl.TextAlign = 3
    lbl.FontName = "Arial"
    lbl.FontSize = 12
    lbl.FontWeight = 600
    Set BuildLabel = lbl

End Function

Public Function GetSizeAllLabels(rst As DAO.Recordset, ctl As Control, lngX As Long, lngY As Long)

    Dim fld As DAO.Field
    Dim strCaption As String
    Dim lngXTemp As Long
    Dim lngYTemp As Long
    For Each fld In rst.Fields
        strCaption = fld.Name
        On Error Resume Next
        strCaption = fld.Properties("Caption")
        On Error GoTo 0
        WizHook.Key = 51488399
        WizHook.TwipsFromFont ctl.FontName, ctl.FontSize, ctl.FontWeight, ctl.FontItalic, ctl.FontUnderline, 0, strCaption, 0, lngXTemp, lngYTemp
        If lngXTemp > lngX Then
            lngX = lngXTemp
        End If
        If lngYTemp > lngY Then
            lngY = lngYTemp
        End If
    Next fld

End Function

Public Function GetMaximalSizeControl(fdb As DAO.Database, strRecordSource As String, ByVal ctl As Control, fld As DAO.Field, lngDXMax As Long, lngDYMax As Long)

    Dim rst As DAO.Recordset
    Dim lngDX As Long
    Dim lngDY As Long
    Dim strText As String
    Dim varFeldname As String
    Dim strSpaltenbreiten() As String
    Dim i As Integer
    Dim intControlType As AcControlType

    lngDXMax = 0
    lngDYMax = 0

    intControlType = ctl.ControlType
    Select Case intControlType
        Case acCheckBox
            Set ctl = CreateControl(ctl.Parent.Name, acTextBox, acDetail)
    End Select

    On Error Resume Next
    intControlType = fld.Properties("DisplayControl")
    On Error GoTo 0
    Select Case intControlType
        Case acComboBox, acListBox
            Set rst = fdb.OpenRecordset(fld.Properties("RowSource"))
            strSpaltenbreiten = Split(fld.Properties("ColumnWidths"), ";")
            For i = LBound(strSpaltenbreiten) To UBound(strSpaltenbreiten)
                If Not strSpaltenbreiten(i) = 0 Then
                    varFeldname = i
                    Exit For
                End If
            Next i
        Case Else
            Set rst = fdb.OpenRecordset(strRecordSource)
            varFeldname = fld.Name
    End Select
    Do While Not rst.EOF
        If IsNumeric(varFeldname) Then
            strText = Nz(rst.Fields(CInt(varFeldname)).Value)
        Else
            strText = Nz(rst.Fields(varFeldname).Value)
        End If
        WizHook.Key = 51488399
        WizHook.TwipsFromFont ctl.FontName, ctl.FontSize, ctl.FontWeight, ctl.FontItalic, ctl.FontUnderline, 0, strText, 0, lngDX, lngDY
        If lngDX > lngDXMax Then
            lngDXMax = lngDX
        End If
        If lngDY > lngDYMax Then
            lngDYMax = lngDY
        End If
        rst.MoveNext
    Loop
    Select Case intControlType
        Case acComboBox
            lngDXMax = lngDXMax + lngDYMax
        Case acListBox
            lngDXMax = lngDXMax + lngDYMax
            lngDYMax = lngDYMax * 8
        Case acCheckBox
            DeleteControl ctl.Parent.Name, ctl.Name
    End Select
    lngDXMax = (lngDXMax + 100) * 1.3
    lngDYMax = lngDYMax * 1.1

End Function

Public Function TableExistsADO(TableName As String) As Boolean

    On Error Resume Next
    TableExistsADO = IsObject(CurrentData.AllTables(TableName))

End Function

Public Function QueryExistsADO(QueryName As String) As Boolean

    On Error Resume Next
    QueryExistsADO = IsObject(CurrentData.AllQueries(QueryName))

End Function
