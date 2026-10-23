Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_DblClick(Cancel As Integer)
    Dim lsForm As String

    p_ReklNr = Me.ID

        NewReklamationE "Bearbeiten", p_ReklNr 'get_p_ReklNr()
        lsForm = "Fehlererfassung - Register"

'    If CurrentProject.AllForms("Fehlererfassung - Register").IsLoaded = True Then
'        Forms("Fehlererfassung - Register")!Kombinationsfeld6297.Visible = False
'        Forms("Fehlererfassung - Register")!txtFirma.Visible = True
'    End If

end2_Form_DblClick:

End Sub
