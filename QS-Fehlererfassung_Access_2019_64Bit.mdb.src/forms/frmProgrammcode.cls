Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'20260803-ROS
Private Sub Form_Close()

    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")

    ldb.Execute "DELETE * FROM tblTempProgrammcode;", dbFailOnError

    FormKoorSave Me

End Sub

Private Sub Form_DblClick(Cancel As Integer)

    SpringeZuProzedur Me.Modul, Me.ProzedurName

End Sub

Private Sub Form_Load()

    FormOpenKoor Me
    Me.txtgesucht.Value = Me.OpenArgs

End Sub
