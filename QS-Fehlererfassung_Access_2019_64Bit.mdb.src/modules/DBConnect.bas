Attribute VB_Name = "DBConnect"
Option Compare Database

Private mDB As DAO.Database

Public Property Get CurrentDBC() As DAO.Database
    If mDB Is Nothing Then
        Set mDB = CurrentDb
    End If
    Set CurrentDBC = mDB
End Property
