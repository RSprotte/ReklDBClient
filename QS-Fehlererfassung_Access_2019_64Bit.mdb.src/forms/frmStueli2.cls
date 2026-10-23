Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public lexcelprov As clsExcelProvider
Public adbprov As clsAvistaProvider

Private Sub cmbTVNArt_Click()

    TVN_Abrufen Me.Artikel, Me.Var

End Sub

Private Sub cmbTVNArt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

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
