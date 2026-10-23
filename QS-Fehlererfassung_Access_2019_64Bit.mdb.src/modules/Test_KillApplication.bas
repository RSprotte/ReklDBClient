Attribute VB_Name = "Test_KillApplication"
Option Compare Database
Option Explicit



#If VBA7 Or Win64 Then
' alle benötigten API's
Private Declare PtrSafe Function FindWindow Lib "user32" Alias _
    "FindWindowA" (ByVal lpClassName As String, _
                   ByVal lpWindowName As String) As Long

Private Declare PtrSafe Function GetWindowThreadProcessId Lib _
    "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Private Declare PtrSafe Function CloseHandle Lib _
    "kernel32" (ByVal hObject As Long) As Long

Private Declare PtrSafe Function OpenProcess Lib _
    "kernel32" (ByVal dwDesiredAccess As Long, _
                ByVal bInheritHandle As Long, _
                ByVal dwProcessId As Long) As Long

Private Declare PtrSafe Function TerminateProcess Lib _
    "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
#Else
' alle benötigten API's
Private Declare Function FindWindow Lib "user32" Alias _
    "FindWindowA" (ByVal lpClassName As String, _
                   ByVal lpWindowName As String) As Long

Private Declare Function GetWindowThreadProcessId Lib _
    "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Private Declare Function CloseHandle Lib _
    "kernel32" (ByVal hObject As Long) As Long

Private Declare Function OpenProcess Lib _
    "kernel32" (ByVal dwDesiredAccess As Long, _
                ByVal bInheritHandle As Long, _
                ByVal dwProcessId As Long) As Long

Private Declare Function TerminateProcess Lib _
    "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
#End If



Public lTaskID As Long
Const PROCESS_TERMINATE = &H1
Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const GW_HWNDNEXT = 2
Const GW_HWNDPREV = 3
Const GW_OWNER = 4
Const GW_CHILD = 5
Const GW_MAX = 5
Const GWL_STYLE = (-16)
Const WS_VISIBLE = &H10000000
Const WS_BORDER = &H800000

Public Function GetWindowTaskID(ByVal hwnd As Long) As Long
    Dim lResult As Long
    Dim lTemp As Long

    lResult = GetWindowThreadProcessId(hwnd, lTemp)
    GetWindowTaskID = lTemp
End Function

Public Sub TerminateTaskID(lTaskID As Long)
    Dim hTask As Long
    Dim lResult As Long

    hTask = OpenProcess(PROCESS_TERMINATE, 0&, lTaskID)
    lResult = TerminateProcess(hTask, 1&)
    lResult = CloseHandle(hTask)
    fprintLogRSP "TerminateTaskID", "ProzessId = " & lTaskID & " hTask: " & hTask & " lResult of CloseHandle: " & lResult

End Sub

Public Function mytest(fwb As Excel.Workbook)
    Dim Anwendungsname As String

    Anwendungsname = "phpDesigner 8" ' Datei xxx.txt mit Editor geöffnet
    TerminateTaskID GetWindowTaskID(fwb.Application.hwnd)
    TerminateTaskID GetWindowTaskID(fwb.Application.hwnd)
End Function
