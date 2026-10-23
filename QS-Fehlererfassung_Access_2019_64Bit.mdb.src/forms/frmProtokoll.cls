Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub cboFeld_AfterUpdate()

    If Me.FilterOn = True Then
        Me.FilterOn = False
        Me.Filter = "[Feld] = '" & Me.cboFeld & "' and [Tabelle] = '" & Me.cobTabelle & "'"
    Else
        Me.Filter = "[Feld] = '" & Me.cboFeld & "'"
    End If
    Me.FilterOn = True

End Sub

Private Sub cmbFilterLoeschen_Click()

    Me.cobTabelle = ""
    Me.cboFeld = ""
    Me.FilterOn = False
    Me.cmbFilterLoeschen.SetFocus

End Sub

Private Sub cobTabelle_AfterUpdate()

    If Me.FilterOn = True Then
        Me.FilterOn = False
        Me.Filter = "[Feld] = '" & Me.cboFeld & "' and [Tabelle] = '" & Me.cobTabelle & "'"
    Else
        Me.Filter = "[Tabelle] = '" & Me.cobTabelle & "'"
    End If
    Me.FilterOn = True

End Sub
