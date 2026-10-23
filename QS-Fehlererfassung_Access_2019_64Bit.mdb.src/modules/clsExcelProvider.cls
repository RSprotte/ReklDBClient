VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long


Private wb As Excel.Workbook
Private objExcel As Excel.Application

Private Sub class_Initialize()

    Set objExcel = CreateObject("Excel.Application")

End Sub


Private Sub class_Terminate()

    objExcel.Quit
    Set objExcel = Nothing

End Sub
Public Function openExcelInstance(fs_Excelwb As String, Optional fs_copy As String = "") As Excel.Workbook

    Dim fs As Object

    Set fs = CreateObject("Scripting.FileSystemObject")

    If fs.FileExists(fs_Excelwb) = True Then
        If fs_copy <> "" Then
            fs.CopyFile fs_Excelwb, fs_copy, False
            If objExcel Is Nothing Then Set objExcel = CreateObject("Excel.Application")
            Set wb = objExcel.Application.Workbooks.Open(fs_copy)
        Else
            Set wb = Excel.Application.Workbooks.Open(fs_Excelwb)
        End If
        Set openExcelInstance = wb
    End If
    Set fs = Nothing
End Function
Public Function openExistExcelInstance(fs_Excelwb As String) As Excel.Workbook

    Dim fs As Object
    Dim WS As Excel.Worksheet
    Dim objExcel As Object

    Set fs = CreateObject("Scripting.FileSystemObject")

    If fs.FileExists(fs_Excelwb) = True Then
        Set wb = Excel.Application.Workbooks.Open(fs_Excelwb)
        Set openExistExcelInstance = wb
    Else
        Set openExistExcelInstance = Nothing
    End If

    Set fs = Nothing
End Function

Public Sub closeExcelInstance()
On Error GoTo Err_closeExcelInstance:

    Dim lwbProcID As Long

    If Not wb Is Nothing Then
        wb.Save
        wb.Close SaveChanges:=False
        Set wb = Nothing
    End If

    Exit Sub
Err_closeExcelInstance:
    If Err = -2147417848 Then
        Resume Next
    End If


End Sub

Public Function getExcelInstance() As Excel.Application
On Error GoTo Err_getExcelInstance:

    Dim lwbProcID As Long

    If Not objExcel Is Nothing Then
        Set getExcelInstance = objExcel
    End If

    Exit Function
Err_getExcelInstance:
    If Err = -2147417848 Then
        Resume Next
    End If


End Function
