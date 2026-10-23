Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cboBasisText_AfterUpdate()

    If IsNull(Me.cboBasisText) Then
        Me.FilterOn = False
    Else
        Me.FilterOn = False
        Me.Filter = "[BasisText] = '" & Me.cboBasisText & "'"
        Me.FilterOn = True
    End If
    Me.Refresh

End Sub

Private Sub cboName_AfterUpdate()

    If IsNull(Me.cboName) Then
        Me.FilterOn = False
    Else
        Me.FilterOn = False
        Me.Filter = "[OjectID_Name] = '" & Me.cboName & "'"
        Me.FilterOn = True
    End If
    Me.Refresh

End Sub

Private Sub cboTextID_AfterUpdate()

    If IsNull(Me.cboTextID) Then
        Me.FilterOn = False
    Else
        Me.FilterOn = False
        Me.Filter = "[TextID] = '" & Me.cboTextID & "'"
        Me.FilterOn = True
    End If
    Me.Refresh

End Sub

Private Sub cboType_AfterUpdate()

    If IsNull(Me.cboType) Then
        Me.FilterOn = False
    Else
        Me.FilterOn = False
        Me.Filter = "[TypeOfObject] = '" & Me.cboType & "'"
        Me.FilterOn = True
    End If
    Me.Refresh

End Sub

Private Sub Form_Close()

    FormKoorSave Me

End Sub

Private Sub Form_Load()

    FormOpenKoor Me

End Sub
