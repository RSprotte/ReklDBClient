VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDictFormulareRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
'Die Klasse/Objekt clsDictFormularRef verwaltetfür alle aktiven Formulare die Fieldlisten
'Key: Formularname (eindeutig)
'Item: clsFieldList des jeweiligen aktiven Formular

Private RefdictFormList As Object


Public Sub class_Initialize()

    If RefdictFormList Is Nothing Then Set RefdictFormList = CreateObject("scripting.dictionary")
    RefdictFormList.RemoveAll

End Sub


Private Sub Finalize()

    RefdictFormList.RemoveAll
    Set RefdictFormList = Nothing

End Sub


Public Sub AddListItem(ByVal fKey As String, ByRef fItem As clsDictFieldList)


    If RefdictFormList.Exists(fKey) Then RefdictFormList.Remove (fKey)
    RefdictFormList.Add fKey, fItem

End Sub
Public Sub RemoveFormListItem(ByVal fKey As String)
'löscht alle Field Items in der FieldList
'löscht selber nicht das ListItem Element

    Dim lcls_dictFieldList As clsDictFieldList

    If RefdictFormList.Exists(fKey) Then
        Set lcls_dictFieldList = GetListItem(fKey)
        lcls_dictFieldList.clsRemoveAll
    End If

End Sub
Public Sub RemoveListItem(ByVal fKey As String)
'löscht das eigentliche FieldList Item

    If RefdictFormList.Exists(fKey) Then RefdictFormList.Remove (fKey)

End Sub

Public Property Get GetListItem(fKey As Variant) As clsDictFieldList

    If Not RefdictFormList.Exists(fKey) Then
        Set GetListItem = Nothing
    Else
        Set GetListItem = RefdictFormList.Item(fKey)
    End If

End Property

Public Property Get GetListCount() As Integer

    GetListCount = RefdictFormList.Count

End Property

Public Property Get GetListItemAsIndex(fIndex As Integer) As clsDictFieldList
    Dim lcls As clsDictFieldList
    'Set clsDictFieldList = RefdictFormList.Item(fIndex)
    Set clsDictFieldList = RefdictFormList.Item("Fehlererfassung - Register")
    Set GetListItemAsIndex = RefdictFormList.Item(fIndex)

End Property

Public Property Get GetDictObject() As Object

    Set GetDictObject = RefdictFormList

End Property
