VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDictFieldList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
'Das Objekt DictFieldList beinhaltet eine Dictionary (Liste of Objects) die zu einem Formular alle geänderten Felder, die über die zugeordneten Steuerlemente des Formulars
'verändert wurden. D.h es sind die geänderten Felder und deren Werte eingetragen. Alle Felder die zur Laufzeit nicht geändert werden sind nicht Bestandteil der Dictionary/List
'Mit der Referenzierung des clsDictFieldList wird ein Dictionary-Object (RefDictFieldList) instanziiert.
'Key List-Element: Als Key für jedes Listeelement (Field-Object) wird der Tabellenname + "#" + Feldname des Tabellenfeld genommen, um eine eindeutigen Key zu haben.
'List-Item: ein clsFiled-Object
'Änderungen von Listelementen wird dadurch erwirkt das das bestehende Elemnt gelöscht und neu angelegt wird.


Private RefdictFieldlist As Object


Public Sub class_Initialize()

    If RefdictFieldlist Is Nothing Then Set RefdictFieldlist = CreateObject("scripting.dictionary")
    RefdictFieldlist.RemoveAll

End Sub


Private Sub Finalize()

    RefdictFieldlist.RemoveAll
    Set RefdictFieldlist = Nothing

End Sub


Public Property Let AddListItem(ByRef fListItem As clsField)


    If RefdictFieldlist.Exists(fListItem.getTabellenname & "#" & fListItem.getFeldname) Then
        RefdictFieldlist.Remove (fListItem.getTabellenname & "#" & fListItem.getFeldname)
    End If
    RefdictFieldlist.Add fListItem.getTabellenname & "#" & fListItem.getFeldname, fListItem

End Property

Public Property Get GetListItem(ByVal fItemKey As Variant) As clsField


    If Not RefdictFieldlist.Exists(fItemKey) Then
        Set GetListItem = Nothing
    Else
        Set GetListItem = RefdictFieldlist.Item(fItemKey)
    End If

End Property

Public Sub clsRemoveAll()

        RefdictFieldlist.RemoveAll

End Sub
Public Sub clsRemoveItemAsKey(ByVal fKey As String)

        RefdictFieldlist.Remove (fKey)

End Sub


Public Function GetListCount() As Integer

    GetListCount = RefdictFieldlist.Count

End Function

Public Property Get GetKeyFromIndex(ByVal fIndex As Integer)

    GetKeyFromIndex = RefdictFieldlist.Keys()(fIndex)

End Property


Public Property Get GetListCount2() As Integer

    GetListCount2 = UBound(RefdictFormList.Keys)

End Property

Public Property Get GetListItemAsIndex(fIndex As Integer) As clsField

        Set GetListItemAsIndex = RefdictFieldlist.Item(fIndex)

End Property

Public Property Get GetDictObject() As Object

    Set GetDictObject = RefdictFieldlist

End Property
