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
    Call FormEinUndAusBlenden(Me)

End Sub
