VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMouseoverControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim WithEvents m_cmd As CommandButton
Attribute m_cmd.VB_VarHelpID = -1
Dim m_Parent As clsMouseoverForm

Public Property Set Parent(obj As clsMouseoverForm)
    Set m_Parent = obj
End Property

Public Property Set CommandButton(cmd As CommandButton)
    Set m_cmd = cmd
    With m_cmd
        .OnMouseMove = "[Event Procedure]"
    End With
End Property

Private Sub m_cmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lHandle As Long

    With m_Parent
        lHandle = LoadCursor(0, IDC_HAND)
        If (lHandle > 0) Then SetCursor lHandle
    End With
End Sub
