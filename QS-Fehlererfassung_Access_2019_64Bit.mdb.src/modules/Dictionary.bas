Attribute VB_Name = "Dictionary"
Option Compare Database

'#######################################################################################
'Richtlinie für das Anlegen von neuen Formularen unter Verwendung der Dictionary.
'
'-   Formularname        frm + Name          frmAnmeldedaten
'-   Tabellenname        tbl + Name des Formular tblAnmeldedaten
'
'Aufbau Tabelle:
'-   Feld 1          immer ID mit Primärschlüssel
'-   Feld 2          z.B. Name
'-   Feld 3          z.B. Vornname
'usw.
'
'Namensgebung der ungebundenen Felder im Formular:
'Felder müssen so benannt werden wie in der Tabelle + davor die Typkürzel des Feldtyp:
'
'z.B. txtName        txt steht für Textfeld
'
'weitere Typkürzel:

'Typkürzel   Objekt – deutsch        Objekt - englisch
'app         Applikation             Application
'bas         (oder mdl) Modul        Module
'cbo         Kombinationsfeld        ComboBox
'chk         Kontrollkästchen        CheckBox
'cmd         Befehlsschaltfläche     CommandButton
'ctl(s)      Steuerelement(e)        Control(s)
'fra         Optionsgruppe           OptionGroup (Frame)
'frm(s)      Formular(e)             Form(s)
'grl         Gruppenebene            GroupLevel
'img         Bild                    Image
'lbl         Bezeichnungsfeld        Label
'lin         Linie                   Line
'lst         Listenfeld              ListBox
'ole         OLE-Objekt              ObjectFrame
'opt         Optionsfeld             OptionButton
'shp         Rechteck                Rectangle
'qry         Abfrage                 Query
'rpt(s)      Bericht(e)              Report(s)
'scr         Bildschirm              Screen
'sec         Bereich                 Section
'sfr         Unterformular           SubForm
'srp         Unterbericht            SubReport
'txt         Textfeld                Textbox
'tgl         Umschaltfläche          ToggleButton
'
'Beispiele:
'txtFName        Textfeld Familienname
'lstMitarbeiter  Listenfeld Mitarbeiter
'#######################################################################################
'_____________________________________________________________________________________________________________________________________________________________________________________


'#### Dictionary für Tabellenfelder
Public dictFormularRef As clsDictFormulareRef

'Public Sub addControlsList(f_Feldname As String, f_Feldtyp As String, f_Feld As Control, f_Tabellenname As String, Optional f_clsFormDict As clsDictFieldList = Nothing, Optional f_Form As String)
'Public Sub addControlsList(f_Feldname As String, f_Feldtyp As String, f_Feld As Control, ByRef fSourceDef As SourceDef, Optional f_clsFormDict As clsDictFieldList = Nothing)
Public Sub addControlsList(f_Feldname As String, f_Feldtyp As String, f_Feld As Control, ByRef fForm As Form, Optional f_clsFormDict As clsDictFieldList = Nothing)

    Dim dictFieldIn As clsField
    Dim ls_Feld() As String



    If f_clsFormDict Is Nothing Then Set f_clsFormDict = dictFormularRef.GetListItem(fForm.Name)

    Set dictFieldIn = New clsField
    dictFieldIn.setFeldname = f_Feldname
    dictFieldIn.setFeldType = f_Feldtyp
    If Not IsNull(f_Feld) Then
        dictFieldIn.setFeldWert = f_Feld.Value
    Else
        dictFieldIn.setFeldWert = ""
    End If
    ls_Feld = Split(f_Feldname, "#")
    dictFieldIn.setTabellenname = ls_Feld(0)
    f_clsFormDict.AddListItem = dictFieldIn

    Set dictFieldIn = Nothing

End Sub


Public Sub UpdatetFormTab(ByRef fForm As Form, fCriterium As String)
On Error Resume Next
'Datenübergabe zum Aktualliesieren der Tabellendaten übergeben wird:
'f_Form -> Formularbezeichnung
'f_Field -> Feldname hinter dem der Index eingetragen ist
'f_ID -> Index auf den gesuchten Datensatz in der hinterlegten Tabelle


    Dim ldb As DAO.Database
    Dim l_dictFieldlist As clsDictFieldList
    Dim lsa_FieldName() As String
    Dim ls1 As String
    Dim ls2 As String
    Dim ls As String
    Dim varKey2 As Variant
    Dim lcl_clsField As clsField

    Dim lAlterWert As Variant
    Dim lTabellenname As String
    Dim lText As String
    Dim lID() As String
    Dim lFeldname As String

    'Vorgangsnummer extrahieren für die Speichung der Daten Vorgangsbezogen in tblProtokoll / 20251203-ROS
    lID = Split(fCriterium, "=")


    Set ldb = openAccDBConnection("openDB")
    'Für aktuelles Formular sein Dictionary für die Verwaltung der Felder erzeugen
    Set l_dictFieldlist = dictFormularRef.GetListItem(fForm.Name)
    If l_dictFieldlist.GetListCount > 0 Then
        ls = ""
        'Zusammensetztung des sql-Strings ls1 "Feldname" zum Updaten der Abfrage
        For Each varKey2 In l_dictFieldlist.GetDictObject
            Set lcl_clsField = l_dictFieldlist.GetListItem(varKey2)
            If lcl_clsField.getFeldname = "Mastertab#Kategorie-Status" Then
'                Debug.Print "Inhaltlicher Breakpoint"
            End If
'            Debug.Print "  --> Field: " & lcl_clsField.getFeldname & " - FeldType: " & lcl_clsField.getFeldType & " - FeldValue: " & lcl_clsField.getFeldWert & " - Tabele: " & lcl_clsField.getTabellenname
            lsa_FieldName = Split(lcl_clsField.getFeldname, "#")
            ls1 = "[" & lsa_FieldName(1) & "]"
            ls2 = ""
'############################################################################################# 20251203-ROS
'Auslesen des Wert vor der Übernahme der Änderung zur Speicherung in tblProtokoll
            lTabellenname = "[" & lsa_FieldName(0) & "]"
            lAlterWert = Nz(DLookup(ls1, lTabellenname, fCriterium))
            lTabellenname = lsa_FieldName(0)
            lFeldname = lsa_FieldName(1)
'###########################################################################################################
            Select Case lcl_clsField.getFeldType
    '            Case dbBoolean
                Case 106
                    If lcl_clsField.getFeldWert Then
                        ls2 = "True"
                    Else
                        ls2 = "False"
                    End If
                Case 109, 111, 107
                    lcl_clsField.setFeldWert = Replace(lcl_clsField.getFeldWert, "'", "''")
                    ls2 = "'" & lcl_clsField.getFeldWert & "'"
            End Select
            'Zusammensetztung des sql-String ls für das Update Feldname ls1 und Feldwert ls2
            If ls2 <> "" Then ls = ls & ls1 & " = " & ls2 & ", "
'#################################################################################################### 20251203-ROS
'geänderte Werte werden ausgelesen und tblProtokoll abgespeichert
            lText = lTabellenname & " # " & lFeldname & " # " & lAlterWert & " # " & Replace(ls2, "'", "")
            ldb.Execute ("Insert into tblProtokoll (Vorgang_ID, Datum, KZZ, Aenderung, Tabelle, Feld, Alter_Wert, Neuer_Wert) values " & _
            "( " & CLng(lID(1)) & ",'" & Now() & "', '" & get_pKuerzel & "', '" & lText & "','" & lTabellenname & "','" & lFeldname & "','" & lAlterWert & "','" & Replace(ls2, "'", "") & "');")
'##################################################################################################################
        Next
        ls = Left(ls, Len(ls) - 2)

        'Datenbank Update unter Verwendung des sql-String ls
        ldb.Execute ("Update [" & fForm.getRecordSource() & "] set " & ls & " WHERE " & fCriterium & ";")

        dictFormularRef.RemoveFormListItem (fForm.Name)
    End If
    Set l_dictFieldlist = Nothing

    Forms(fForm.Name).Requery

End Sub



Private Function QueryExistsADO(QueryName As String) As Boolean

    On Error Resume Next
    QueryExistsADO = IsObject(CurrentData.AllQueries(QueryName))

End Function


Public Sub AddChangeFieldFormDictionary(tblForm As String, tblField As String, tblFormID As Long)
'Die Prozedur erzeugt entweder ein neues Listenelement oder
'überschreibt ein schon bestehendes Fieldelemet im Field-Dictionary des aktuellen Formular

    Dim clt As Control
    Dim Wert As String
    Dim strSourceTable As String
    Dim dictFieldOut As clsField
    Dim sFeldkey As String
    Dim Field As DAO.Field
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset


    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from " & tblForm & "  where [" & tblField & "] = " & tblFormID & ";")


    If rs.RecordCount Then
        For Each Field In rs.Fields
                sFeldkey = Field.Name
                If IsNull(Field.Value) Then
                    Wert = ""
                Else
                    Wert = Field.Value
                End If
                'Call addControlsList(sFeldkey, Field.Type, Wert, tblForm)
                Set dictFieldOut = dictFieldlist.GetListItem(sFeldkey)
'                Debug.Print dictFieldOut.getFeldname, dictFieldOut.getFeldType, dictFieldOut.getFeldWert, dictFieldOut.getTabellenname
                Set dictFieldOut = Nothing
        Next Field
    End If
    Set rs = Nothing

End Sub
Public Sub FormOpenDictionary(fForm As Form)
'Übergabe Formular f_Form zu erzeugen der Dictionary und INtegration in den FormularContainer

    Dim clt As Control
    Dim Wert As String
    Dim strSourceTable As String
    Dim dictFieldOut As clsField
    Dim sFeldkey As String
    Dim tblForm As String

    Dim l_clsDictFieldList As clsDictFieldList
    'Dictonary für die Formularverwaltung erzeugen. Das Dict verwaltet für jedes Formular in Abhängigkeite seine FormName (Key) ein Dictionary, das alle Felder beinhaltet
    If dictFormularRef Is Nothing Then Set dictFormularRef = New clsDictFormulareRef

    'Für aktuelles Formular sein Dictionary für die Verwaltung der Felder erzeugen
    Set l_clsDictFieldList = New clsDictFieldList
    'Das Form Dictionary in die Verwaltung der Form Dict's einfügen
    dictFormularRef.AddListItem fForm.Name, l_clsDictFieldList
'    Debug.Print "Anzahl Listelemente von dictFormularRef: " & dictFormularRef.GetListCount

    '!!!!!????? wird mit nothing nicht auch der Speicher der Referenz komplett freigegeben und somit die Ref in der Verwaltung
    Set l_clsDictFieldList = Nothing

End Sub

Public Sub FormCloseDictionary(ByRef fForm As Form)
'Eleminiert das FieldList Element des Formulars aus der Form List

    Dim clt As Control
    Dim Wert As String
    Dim strSourceTable As String
    Dim dictFieldOut As clsField
    Dim sFeldkey As String
    Dim tblForm As String

    If dictFormularRef Is Nothing Then
'        MsgBox ("Error: Form Dictionary nicht referenziert beim Close-Form!")
        MsgBox getTC_FCFromCode("TC4280"), vbInformation, getTC_FCFromCode("TC4279")
        Exit Sub
    End If

    If Not dictFormularRef.GetListItem(fForm.Name) Is Nothing Then
        dictFormularRef.RemoveFormListItem (fForm.Name)
        dictFormularRef.RemoveListItem (fForm.Name)
    Else
'        MsgBox ("Error: Für Formular gibt es kein FieldList Eintrag!")
        MsgBox getTC_FCFromCode("TC4278"), vbInformation, getTC_FCFromCode("TC4279")
    End If


End Sub


Public Function getChangeCountCurrForm(fForm As Form) As Integer
On Error GoTo err_getChangeCountCurrForm
'Gibt die Anzahl geänderter Felder in einem geöffneten Formular zurück

    Dim l_dictFieldlist As clsDictFieldList

    'Init der Formular Änderungsliste.
    Set l_dictFieldlist = dictFormularRef.GetListItem(fForm.Name)
    getChangeCountCurrForm = l_dictFieldlist.GetListCount

    Exit Function
err_getChangeCountCurrForm:
    If Err = 91 Then Exit Function

End Function

Public Function RemoveAllFormList(fForm As Form) As Integer
'Gibt die Anzahl geänderter Felder in einem geöffneten Formular zurück

    Dim l_dictFieldlist As clsDictFieldList

    'Init der Formular Änderungsliste.
    Set l_dictFieldlist = dictFormularRef.GetListItem(fForm.Name)
    l_dictFieldlist.clsRemoveAll

End Function

Public Sub PrintDebugAllFormFieldValues()

    Dim lcl_clsFieldList As clsDictFieldList
    Dim lcl_clsField As clsField
    Dim varKey1 As Variant
    Dim varKey2 As Variant

    If Not dictFormularRef Is Nothing Then
        For Each varKey1 In dictFormularRef.GetDictObject
            Set lcl_clsFieldList = dictFormularRef.GetListItem(varKey1)
'            Debug.Print "Formular: " & varKey1
            For Each varKey2 In lcl_clsFieldList.GetDictObject
                Set lcl_clsField = lcl_clsFieldList.GetListItem(varKey2)
'                Debug.Print "  --> Field: " & lcl_clsField.getFeldname & " - FeldType: " & lcl_clsField.getFeldType & " - FeldValue: " & lcl_clsField.getFeldWert & " - Tabele: " & lcl_clsField.getTabellenname
            Next
        Next
    End If
End Sub
