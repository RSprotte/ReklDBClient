Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CobHtmlCode_Click()

    Dim sText As String
    Dim txt2html As String
        ' Text-Datei zeilenweise einlesen
        ' und Sonderzeichen, wie Umlaute und spitze
        ' Klammern durch HTML-Steuercodes ersetzen

        sText = Me.txtFreitext
        sText = Replace(sText, "&", "&amp;")
        sText = Replace(sText, """", "&quot;")
        sText = Replace(sText, "<", "&lt;")
        sText = Replace(sText, ">", "&gt;")
        sText = Replace(sText, vbCrLf, "<br>")
        sText = Replace(sText, "  ", " &nbsp;")

        txt2html = sText
        Me.txtHTMLText = txt2html

        Me.Requery
        Me.Refresh
End Sub
