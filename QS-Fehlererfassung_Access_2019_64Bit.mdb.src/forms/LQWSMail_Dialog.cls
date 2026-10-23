Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_Load()

    Dim objMouseOverForm As clsMouseoverForm
    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

'    FormOpenKoor Me

    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long

    lgPosition = getPointerPos(XPos, YPos)
    FormOpenKoorMousPointer Me, XPos, YPos, 909, 197

    Call FormEinUndAusBlenden(Me)

End Sub
Private Sub Schließen_Click()

If IsLoaded("AdminLQFehlerortbestimmung") Then
    Forms("AdminLQFehlerortbestimmung").Requery
    Forms("AdminLQFehlerortbestimmung").Refresh
End If

DoCmd.Close

End Sub



Private Sub SelLQWS_AfterUpdate()

    Me.Filter = "[ID] = " & Me.SelLQWS
    Me.FilterOn = True
    Me.Requery
    Me.Refresh

End Sub
