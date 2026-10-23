Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub Alle_AfterUpdate()

    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")

    If Me.Alle Then
        If Me.Filter = "" Then
            ldb.Execute ("UPDATE [Mastertab-DSBerechtigungTemp] " & _
                    "SET [Löschen] = True;")
        Else
            ldb.Execute ("UPDATE [Mastertab-DSBerechtigungTemp] " & _
                    "SET [Löschen] = True WHERE ([Kürzel] = '" & Me.KFKürzel & "');")
        End If
    Else
         If Me.Filter = "" Then
            ldb.Execute ("UPDATE [Mastertab-DSBerechtigungTemp] " & _
                    "SET [Löschen] = False;")
        Else
            ldb.Execute ("UPDATE [Mastertab-DSBerechtigungTemp] " & _
                    "SET [Löschen] = False WHERE ([Kürzel] = '" & Me.KFKürzel & "');")
        End If
    End If
    Me.Requery
    Me.Refresh
     ' Set ldb = openAccDBConnection("closeDB")

End Sub

Private Sub Form_Open(Cancel As Integer)

    Me.KFKürzel = Null
    Me.Filter = ""

End Sub

Private Sub Freigeben_Click()
On Error GoTo Err_Freigeben_Click

    Dim ldb As DAO.Database
    Dim ls As String

    DoCmd.Save
    Set ldb = openAccDBConnection("openDB")

    If Not Me.Filter = "" Then
        ldb.Execute ("UPDATE [Mastertab-DSBerechtigung] INNER JOIN [Mastertab-DSBerechtigungTemp] ON [Mastertab-DSBerechtigung].ID = [Mastertab-DSBerechtigungTemp].ID " & _
                "SET [Mastertab-DSBerechtigung].Zustand = 'Frei' WHERE ([Mastertab-DSBerechtigungTemp].Löschen=True AND [Mastertab-DSBerechtigungTemp].Kürzel = '" & Me.KFKürzel & "');")
    Else
        ldb.Execute ("UPDATE [Mastertab-DSBerechtigung] INNER JOIN [Mastertab-DSBerechtigungTemp] ON [Mastertab-DSBerechtigung].ID = [Mastertab-DSBerechtigungTemp].ID " & _
                "SET [Mastertab-DSBerechtigung].Zustand = 'Frei' WHERE ((([Mastertab-DSBerechtigungTemp].Löschen)=True));")
    End If
     ' Set ldb = openAccDBConnection("closeDB")

    DoCmd.Close
    Exit Sub
Exit_Freigeben_Click:
     ' Set ldb = openAccDBConnection("closeDB")

    DoCmd.Close
    Exit Sub

Err_Freigeben_Click:
    MsgBox Err.Description
    Resume Exit_Freigeben_Click


End Sub

Private Sub KFKürzel_AfterUpdate()

    If Not IsNull(Me.KFKürzel) Then
        Me.Filter = "[Kürzel] = '" & Me.KFKürzel & "'"
        Me.FilterOn = True
    Else
        Me.Filter = ""
        Me.FilterOn = True
    End If
    Me.Requery
    Me.Refresh

End Sub

Private Sub Löschen_Click()

    Me.Requery
    Me.Refresh

End Sub
