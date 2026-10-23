VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMailProviderOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Event Handler für das Mailitem um das Event "Senden" abzufangen und daraus Folgeaktionen abzuleiten
Public WithEvents olMailItem1 As Outlook.MailItem
Attribute olMailItem1.VB_VarHelpID = -1

Dim olApp As Outlook.Application  'Outlook Obbject
'Dim olMailItem1 As Outlook.MailItem
Dim msVon As String
Dim msAn As String
Dim msCC As String
Dim msBCC As String
Dim msBetreff As String
Dim msBody As String
Dim msSaveAs As String
Dim msKuerzel As String

'
Dim mbReadOnly As Boolean

'Refrenz des steuerenden Formulars
Dim mfrmParent As Form


Private Sub class_Initialize()
On Error GoTo err_class_Initialize

    Set olApp = CreateObject("Outlook.Application")
    'Set olMailItem1 = olApp.CreateItem(olMailItem)
    msVon = ""
    msAn = ""
    msCC = ""
    msBCC = ""
    msBetreff = ""
    msBody = ""
    msKuerzel = ""
    msSaveAs = ""
    mbReadOnly = False

'    Set olApp = CreateObject("Outlook.Application")
    Exit Sub

err_class_Initialize:

    If Err = 91 Then
        Resume Next
    End If

End Sub

Private Sub class_Terminate()

    If Not olApp Is Nothing Then
        Set olApp = Nothing
    End If

End Sub

'Member methods
Public Property Get getVon() As String
    getVon = msVon
End Property

Public Property Let setVon(ByVal fsVon As String)
    msVon = fsVon
End Property

Public Property Get getAn() As String
    getAn = msAn
End Property

Public Property Let setAn(ByVal fsAn As String)
    msAn = fsAn
End Property

Public Property Get getCC() As String
    getCC = msCC
End Property

Public Property Let setCC(ByVal fsCC As String)
    msCC = fsCC
End Property

Public Property Get getBCC() As String
    getBCC = msBCC
End Property

Public Property Let setBCC(ByVal fsBCC As String)
    msBCC = fsBCC
End Property

Public Property Get getBetreff() As String
    getBetreff = msBetreff
End Property

Public Property Let setBetreff(ByVal fsBetreff As String)
    msBetreff = fsBetreff
End Property

Public Property Get getBody() As String
    getBody = msBody
End Property

Public Property Let setBody(ByVal fsBody As String)
    msBody = fsBody
End Property

Public Property Get getSaveAs() As String
    getSaveAs = msSaveAs
End Property

Public Property Let setSaveAs(ByVal fsSaveAs As String)
    msSaveAs = fsSaveAs
End Property
Public Property Get getKuerzel() As String
    getKuerzel = msKuerzel
End Property

Public Property Let setKuerzel(ByVal fsKuerzel As String)
    msKuerzel = fsKuerzel
End Property
Public Property Get getfrmParent() As Form
    Set getfrmParent = mfrmParent
End Property

Public Property Let setfrmParent(ByVal ffrmForm As Form)
    Set mfrmParent = ffrmForm
End Property
Public Property Get getReadOnly() As Boolean
    getReadOnly = mbReadOnly
End Property

Public Property Let setReadOnly(ByVal fbReadOnly As Boolean)
    mbReadOnly = fbReadOnly
End Property

'methods
Public Function getOLProvInstanz() As Outlook.Application

    If olApp Is Nothing Then
        Set olApp = New Outlook.Application
    End If
    Set getOLProvInstanz = olApp

End Function
Public Function getolMailItem1() As Outlook.MailItem

    If olMailItem1 Is Nothing Then
        Set olMailItem1 = olApp.CreateItem(olMailItem)
    End If
    Set getolMailItem1 = olMailItem1

End Function


Public Sub fillMailItem()

    Dim loMailItem As Outlook.MailItem

    If olMailItem1 Is Nothing Then
        Set olMailItem1 = getolMailItem1()
    End If
    With olMailItem1
        '.SenderEmailAddress = msVon
        .SentOnBehalfOfName = msVon
        .To = msAn
        .CC = msCC
        .BCC = msBCC
        .subject = msBetreff
        .HTMLBody = msBody
    End With
    'Set loMailItem = Nothing

End Sub

Public Sub displayMailItem()

    If Not olMailItem1 Is Nothing Then
        olMailItem1.Display
    End If

End Sub

Public Sub sendMailItem()
On Error GoTo err_sendMailItem

    If Not olMailItem1 Is Nothing Then
        olMailItem1.send
    End If

    Exit Sub
err_sendMailItem:

    If Err Then
        Resume Next
    End If

End Sub

Public Sub SaveAsMailItem(fsSaveAs As String, Optional fsKuerzel As String = "")
On Error GoTo err_SaveAsMailItem

    If Not olMailItem1 Is Nothing Then
        If msBetreff <> olMailItem1.subject Then
            msBetreff = olMailItem1.subject
        End If
        msBetreff = SonderzeichenEntfernen(msBetreff)
        If fsSaveAs <> "" Then
            If fsKuerzel = "" Then
                olMailItem1.SaveAs fsSaveAs & format(Date, "yyyymmdd") & "-" & msKuerzel & "_" & msBetreff & ".msg", olMSG
            Else
                olMailItem1.SaveAs fsSaveAs & format(Date, "yyyymmdd") & "-" & fsKuerzel & "_" & msBetreff & ".msg", olMSG
            End If
        End If
'        If fsSaveAs <> "" Then olMailItem1.SaveAs fsSaveAs & "test.msg", olMSG
    End If
    Exit Sub

err_SaveAsMailItem:

    If Err <> 0 Then
        MsgBox "Err: " & Err & " - " & Error
        Resume Next
    End If


End Sub

Public Sub AddItemMailItem(fsAnhang As String)
On Error GoTo err_AddItemMailItem

    If olMailItem1 Is Nothing Then
        Set olMailItem1 = getolMailItem1()
    End If
    olMailItem1.Attachments.Add (fsAnhang)
    Exit Sub

err_AddItemMailItem:

    If Err <> 0 Then
        MsgBox "Err: " & Err & " - " & Error
        Resume Next
    End If




End Sub

Public Sub OpenItemFromFile(fsFile As String)

'        Set olMailItem1 = olApp.CreateItemFromTemplate(fsFile)
        Set olMailItem1 = olApp.Session.OpenSharedItem(fsFile)
        displayMailItem
        'Wenn sich ein steuerndes Formular eingetragen hat, die CallBack Funktion aufrufen
        If Not mfrmParent Is Nothing Then
            mfrmParent.callBackOutlookProvOpenItem
        End If


End Sub



'### Methoden des Event Handlers #####

Private Sub olMailItem1_Send(Cancel As Boolean)
    'Mail sichern wenn Member msSaveAs vorhanden ist

    Cancel = False

    If msSaveAs <> "" Then
        SaveAsMailItem msSaveAs
    End If
    'Wenn sich ein steuerndes Formular eingetragen hat, die CallBack Funktion aufrufen
    If Not mfrmParent Is Nothing Then
        mfrmParent.callBackOutlookProv
    End If

    If mbReadOnly Then Cancel = True

End Sub
