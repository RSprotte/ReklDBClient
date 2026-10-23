Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub txtInfo_DblClick(Cancel As Integer)
    If Forms("Fehlererfassung - Register").Controls("Mastertab#Kategorie-Status") = "Abgeschlossen" Then
        DoCmd.OpenForm ("frmInfoLief"), , , "[Lieferant-ID] = " & Me.[Lieferant-ID], acFormReadOnly
    Else
        DoCmd.OpenForm ("frmInfoLief"), , , "[Lieferant-ID] = " & Me.[Lieferant-ID], acFormEdit
    End If
End Sub
