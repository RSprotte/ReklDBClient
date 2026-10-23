Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()

    CurrentDb.Execute ("delete * from tblLagerTemp;")
    FormKoorSave Me

End Sub


Private Sub Form_Load()

    FormOpenKoor Me

End Sub
