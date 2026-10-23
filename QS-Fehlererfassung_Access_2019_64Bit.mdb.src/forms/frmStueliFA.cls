Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public lexcelprov As clsExcelProvider
Public adbprov As clsAvistaProvider

Private Sub Form_Close()

    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")
    ldb.Execute ("delete * from tblTempStueli")
    FormKoorSave Me

End Sub

Private Sub Form_Load()

    Call FormEinUndAusBlenden(Me)
    FormOpenKoor Me

End Sub
