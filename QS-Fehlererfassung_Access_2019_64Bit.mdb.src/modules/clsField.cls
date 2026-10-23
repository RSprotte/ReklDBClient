VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private Tabellenname As String  'Tab-Name
Private Feldname As String      'Name des Tabellenfeld
Private Feldtyp As String       'Type Control of Form
Private FeldWert As String      'Current Value of Control - Change Value


Public Property Get getFeldname() As String
'Gibt den aktuellen Tabellen Feldname des aktuellen referenzierten FieldObject zurück
    getFeldname = Feldname
End Property
Public Property Get getFeldType() As String
'Gibt den aktuellen Type des Steuerelemnts des aktiven Formulars zurück
    getFeldType = Feldtyp
End Property
Public Property Get getFeldWert() As String
'Gibt den aktuellen Wert des refrenzoerten Feldes (Steuerlementes) zurück
    getFeldWert = FeldWert
End Property
Public Property Get getTabellenname() As String
'Gibt die zugeordnete Tabelle des Feldes bzw. Controls  zurück
    getTabellenname = Tabellenname
End Property
Public Property Let setFeldname(ByVal fsValue As String)
    Feldname = fsValue
End Property
Public Property Let setFeldType(ByVal fsValue As String)
    Feldtyp = fsValue
End Property
Public Property Let setFeldWert(ByVal fsValue As String)
    FeldWert = fsValue
End Property
Public Property Let setTabellenname(ByVal fsValue As String)
    Tabellenname = fsValue
End Property

Public Sub class_Initialize()

    Feldname = ""
    Feldtyp = ""
    FeldWert = ""
    Tabellenname = ""

End Sub
