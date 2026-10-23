Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()

    FormKoorSave Me

End Sub

Private Sub Form_Load()

    FormOpenKoor Me
    Me.Name_EMail = 9
    Call Name_EMail_AfterUpdate

End Sub

Private Sub Name_EMail_AfterUpdate()

    Me.Filter = "[ID] = " & Me.Name_EMail.Column(0) & ""
    Me.FilterOn = True

End Sub
