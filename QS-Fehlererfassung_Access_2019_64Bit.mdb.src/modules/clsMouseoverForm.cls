VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMouseoverForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim WithEvents m_Form As Form
Attribute m_Form.VB_VarHelpID = -1
Dim colControls As Collection
Dim WithEvents frmDetail As Section
Attribute frmDetail.VB_VarHelpID = -1


Public Property Set Form(frm As Form)
    Dim objMouseOverControl As clsMouseoverControl
    Dim ctl As Control
    Set m_Form = frm
    Set frmDetail = m_Form.Section(0)
    With m_Form
        .OnMouseMove = "[Event Procedure]"
    End With
    With frmDetail
        .OnMouseMove = "[Event Procedure]"
    End With
    Set colControls = New Collection
    For Each ctl In m_Form.Controls
        Select Case ctl.ControlType
            Case acCommandButton
                Set objMouseOverControl = New clsMouseoverControl
                With objMouseOverControl
                    Set .CommandButton = ctl
                    Set .Parent = Me
                End With
                colControls.Add objMouseOverControl
        End Select
    Next ctl
End Property
