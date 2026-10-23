Attribute VB_Name = "FileUndVerzeichnisFunktionen"
Option Compare Database

Public Function getFileExists(ByVal f_File As String) As Boolean

    Dim fs As New FileSystemObject
    getFileExists = False
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(f_File) Then getFileExists = True

    Set fs = Nothing

End Function

Public Function getDirectoryExists(ByVal f_Dir As String) As Boolean

    Dim fs As New FileSystemObject
    Dim F

    getDirectoryExists = False
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set F = fs.GetFolder(f_Dir)
    If Not F Is Nothing Then getDirectoryExists = True

    Set fs = Nothing


End Function

Public Function getFileListFromDirectory(ByVal f_Dir As String, ByRef f_FileArr() As String) As Boolean

    Dim fs As New FileSystemObject
    Dim fo
    Dim fi As File
    Dim iFilecount As Integer

    getFileListFromDirectory = False
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set fo = fs.GetFolder(f_Dir)
    ReDim Preserve f_FileArr(0)
    iFilecount = 0
    If Not fo Is Nothing Then
        For Each fi In fo.Files
            iFilecount = iFilecount + 1
            ReDim Preserve f_FileArr(iFilecount)
            f_FileArr(iFilecount - 1) = fi.Name
        Next
    End If

    If iFilecount > 0 Then getFileListFromDirectory = True

    Set fi = Nothing
    Set fo = Nothing
    Set fs = Nothing


End Function

Public Function readTextFile(ByVal fFile As String) As String
'fFile = Pfad + Filename

    Dim sPfadDatei As String
    Dim fs As Object


    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(fFile) = True Then
        Set A = fs.OpenTextFile(fFile)
        readTextFile = A.ReadAll
        A.Close
    Else
        readTextFile = ""
    End If
    Set fs = Nothing

End Function
Public Function writeTextFile(fFile As String, ByVal fText As String, Optional ByVal fActivity As String = "Attempt") As String
'fFile = Pfad + Filename

    Dim sPfadDatei As String
    Dim fs As Object


    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(fFile) = True Then
        If fActivity = "Overwrite" Then
            Set A = fs.OpenTextFile(fFile, 2)
        Else
            Set A = fs.OpenTextFile(fFile, 8)
        End If
    Else
        fs.CreateTextFile (fFile)
        Set A = fs.OpenTextFile(fFile, 8)
    End If
    A.Write (fText)
    A.Close
    Set fs = Nothing

End Function
