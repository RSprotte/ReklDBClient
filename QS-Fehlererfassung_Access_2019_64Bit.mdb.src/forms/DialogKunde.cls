Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Abbrechen_Click()
    DoCmd.Close acForm, "DialogKunde"
End Sub

Private Sub Form_Open(Cancel As Integer)

    If psDialogKunde = "BearbeitenKunde" Then
        If IsLoaded("KundenUebersicht") Then
            With Forms("KundenUebersicht")
                Me.SKundenNr = .Controls("KndNr_Avista")
                Me.SFirmierung = .Controls("Firma_Name")
                Me.SZuordnung = .Controls("Zurodnung_ID")
            End With
        End If
        Me.Refresh
    End If

End Sub

Private Sub Speichern_Click()
'###### Änderung ######
'Die Änderung muss jetzt nicht in der Tabelle Kunden sondern in tblFirma durchgeführt werden.

On Error GoTo err_Speichern_Click


    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset


    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from tblFirma where [KndNr_Avista] = '" & Me.SKundenNr & "';")
    If rs.RecordCount Then
        ldb.Execute ("update tblFirma set [Firma_Name] = '" & Me.SFirmierung & "', [Zurodnung_ID] = " & Me.SZuordnung & " where [KndNr_Avista] = '" & Me.SKundenNr & "';")
    Else
        ldb.Execute ("Insert into tblFirma ([Firma_Name],[KndNr_Avista], [Zurodnung_ID]) values ('" & Me.SFirmierung & "','" & Me.SKundenNr & "'," & Me.SZuordnung & ");")
    End If
     ' Set ldb = openAccDBConnection("closeDB")
    DoCmd.Close acForm, "DialogKunde"
    Set rs = Nothing
    Exit Sub

err_Speichern_Click:
    MsgBox Error
    Resume Next

End Sub
