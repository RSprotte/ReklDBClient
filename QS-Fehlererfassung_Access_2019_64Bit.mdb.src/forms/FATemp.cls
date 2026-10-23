Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_DblClick(Cancel As Integer)


Forms("FehlererfassungIntern").Controls("FA") = Me.aint_nr
Forms("FehlererfassungIntern").Controls("FALosgroesse") = Me.aint_bedarf_menge


With Forms("FehlererfassungIntern")
    .Requery
    .Refresh
End With

Me.SetFocus
DoCmd.Close

End Sub
