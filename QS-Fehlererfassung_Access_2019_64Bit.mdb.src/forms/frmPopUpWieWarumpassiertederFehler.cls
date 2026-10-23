Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)

    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long


    lgPosition = getPointerPos(XPos, YPos)
    FormOpenKoorMousPointer Me, XPos, YPos, 520, 250, False

End Sub
