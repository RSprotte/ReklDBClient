Attribute VB_Name = "ProtokollierungAllgemein"
Option Compare Database
Option Explicit

Private Sub fprintf(fFrom As String, fOut As String)

Dim fso As New FileSystemObject
'the file we're going to write to
Dim ts As TextStream

If get_LoggingClient() = "Aktiv" Or get_LoggingClientAll() = "Aktiv" Then
    'open this file to write to it
    'Set ts = fso.CreateTextFile("ReklDB_LogFile.txt", False)
    Set ts = fso.OpenTextFile("ReklDB_LogFile.txt", ForAppending, True)

    'write out a couple of lines
    ts.Write (Now & ": " & fFrom & " - ")
    ts.WriteLine (fOut)
    ts.Close
End If


End Sub

Public Sub MsgBoxRSP(ByVal fFrom As String, ByVal fOut As String, ByVal fTitle As String)

fprintf fFrom, fOut
MsgBox fOut, , fTitle


End Sub
Public Sub fprintLogRSP(ByVal fFrom As String, ByVal fOut As String, Optional ByRef fsMail As String = "")

fprintf fFrom, fOut

If fsMail <> "" Then
    'Mail senden an Service
    Dim lclsMailProv As clsMailProviderOutlook
    Set lclsMailProv = New clsMailProviderOutlook
    'Mailinhalte füllen
    lclsMailProv.setAn = getTC_FCFromCode("TC0510")        'service@de.lq-group.com
    lclsMailProv.setVon = lclsMailProv.getAn
    lclsMailProv.setBetreff = "Fehler - Von: " & fFrom
    lclsMailProv.setBody = "Fehler: " & fOut
    lclsMailProv.fillMailItem

    lclsMailProv.sendMailItem

End If

End Sub
