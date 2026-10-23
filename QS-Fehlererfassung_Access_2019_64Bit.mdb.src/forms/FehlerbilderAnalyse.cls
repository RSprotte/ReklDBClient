Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub FehlerbilderAnalyse_Click()

    DoCmd.OpenForm "FehlerbilderAnalyse", acNormal
    With Forms("FehlerbilderAnalyse")
        .Controls("ListTaetigkeiten") = Me.ID
        .Controls("ListTaetigkeiten").Requery
        .Requery
        .Refresh
        .Controls("ListTaetigkeiten").Locked = True
        .Controls("DefTaetigkeit").Locked = True
        .UF01.Controls("Beschreibung").Locked = True
        .Controls("Übernehmen").Enabled = False
        Forms("Fehlerbilder").AllowAdditions = False
    End With


End Sub

Private Sub Neu_Click()
        DoCmd.OpenForm "Dialog_FehlerbildAnalyse"
        DialogFehlerbildAnalyse = "Neu"
        DefTaetigkeitID = Me.ListFehlerbilder

End Sub
Private Sub Schließen_Click()
     DoCmd.Close
     If CurrentProject.AllForms("Fehlererfassung - Register").IsLoaded Then
        Forms("Fehlererfassung - Register").Requery
        Forms("Fehlererfassung - Register").Refresh
     End If
End Sub
