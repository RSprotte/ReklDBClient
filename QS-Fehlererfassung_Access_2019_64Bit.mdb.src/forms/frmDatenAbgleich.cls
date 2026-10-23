Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub Bis_LostFocus()

'    Call setFilter

End Sub

Private Sub cboBearbeitet_Click()

    Dim curAmntDue As Currency, lngBlack As Long
 Dim lngRed As Long, lngYellow As Long, lngWhite As Long

 lngRed = RGB(255, 0, 0)
 lngBlack = RGB(0, 0, 0)
 lngYellow = RGB(255, 255, 0)
 lngWhite = RGB(255, 255, 255)
' me.ID.BorderColor = lngWhite
' me.ID.ForeColor = lngWhite
' me.ID.BackColor = lngWhite
 Me.ID.BackColor = lngWhite



End Sub

Private Sub cmbKatBe1_AfterUpdate()

    Call setFilter

End Sub

Private Sub cmbKatBe2_AfterUpdate()

    Call setFilter

End Sub

Private Sub cmbKatProblem_AfterUpdate()

    Call setFilter

End Sub

Private Sub cmbVerursacher_AfterUpdate()

    Call setFilter

End Sub

Private Sub cmbFilterSetzten_Click()

'    Dim ls_Filter As String
'
'    ls_Filter = ""
'
'
'    If Not ((Me.Von = "")) Then
'        ls_DateVon = Format(Me.Von, "mm\/dd\/yyyy")
'        ls_DateBis = Format(Me.Bis, "mm\/dd\/yyyy")
'        'lb_ADate = True
'        If Me.Bis = "" Or IsNull(Me.Bis) Then
'            Me.Bis = Date
'            Me.Bis.Requery
'        End If
'        ls_Filter = ls_Filter & "([Auftragsdatum] between #Von# and #Bis#) AND "
'        ls_Filter = Replace(ls_Filter, "Von", ls_DateVon)
'        ls_Filter = Replace(ls_Filter, "Bis", ls_DateBis)
'    End If
'
'
'
'    If Me.cmbVerursacher <> "" Then
'        ls_Filter = ls_Filter & "(Anzeige = '" & Me.cmbVerursacher & "') AND "
'    End If
'
'    If Me.cmbKatBe1 <> "" Then                                                 '20191009-ROS
'        ls_Filter = ls_Filter & "([Kategorie-Bearbeitung I] = '" & Me.cmbKatBe1 & "') AND "
'    End If
'
'    If Me.cmbKatBe2 <> "" Then
'        ls_Filter = ls_Filter & "([Kategorie-Bearbeitung II] = '" & Me.cmbKatBe2 & "') AND "
'    End If
'
'    If Me.cmbKatProblem <> "" Then
'        ls_Filter = ls_Filter & "([Kategorie-Problem] = '" & Me.cmbKatProblem & "')"
'    End If
'
'    If Right(ls_Filter, 5) = " AND " Then
'        Me.Filter = Left(ls_Filter, Len(ls_Filter) - 5)
'    Else
'        Me.Filter = ls_Filter
'    End If
'
'
'    Me.FilterOn = True
'    Me.Requery

    Call setFilter

End Sub

Private Sub cmbFilterLöschen_Click()

    Me.cmbVerursacher = ""
    Me.cmbKatBe1 = ""
    Me.cmbKatBe2 = ""
    Me.cmbKatProblem = ""
    Me.FilterOn = False
    Me.Requery
    Me.Refresh

End Sub

Private Sub Form_DblClick(Cancel As Integer)
    Dim lsForm As String

    p_ReklNr = Me.ID
'    If Not IsNull(Me.ErfasserIntern) Then
'        OpenReklamationI
'        lsForm = "Fehlererfassung - Register"
'    Else
        NewReklamationE "Bearbeiten", get_p_ReklNr()
        lsForm = "Fehlererfassung - Register"
'    End If


end2_Form_DblClick:
End Sub
Private Sub setFilter()

    Dim ls_Filter As String

    ls_Filter = ""

    Me.FilterOn = False

    If Not ((Me.Von = "")) Then
        ls_DateVon = format(Me.Von, "mm\/dd\/yyyy")
        ls_DateBis = format(Me.Bis, "mm\/dd\/yyyy")
        'lb_ADate = True
        If Me.Bis = "" Or IsNull(Me.Bis) Then
            Me.Bis = Date
            Me.Bis.Requery
        End If
        ls_Filter = ls_Filter & "([Auftragsdatum] between #Von# and #Bis#) AND "
        ls_Filter = Replace(ls_Filter, "Von", ls_DateVon)
        ls_Filter = Replace(ls_Filter, "Bis", ls_DateBis)
    End If



    If Me.cmbVerursacher <> "" Then
        ls_Filter = ls_Filter & "(Anzeige = '" & Me.cmbVerursacher & "') AND "
    End If

    If Me.cmbKatBe1 <> "" Then                                                 '20191009-ROS
        ls_Filter = ls_Filter & "([Kategorie-Bearbeitung I] = '" & Me.cmbKatBe1 & "') AND "
    End If

    If Me.cmbKatBe2 <> "" Then
        ls_Filter = ls_Filter & "([Kategorie-Bearbeitung II] = '" & Me.cmbKatBe2 & "') AND "
    End If

    If Me.cmbKatProblem <> "" Then
        ls_Filter = ls_Filter & "([Kategorie-Problem] = '" & Me.cmbKatProblem & "')"
    End If

    If Right(ls_Filter, 5) = " AND " Then
        Me.Filter = Left(ls_Filter, Len(ls_Filter) - 5)
    Else
        Me.Filter = ls_Filter
    End If


    Me.FilterOn = True
    Me.Requery



End Sub


Private Sub InterneDaten_ROS_AfterUpdate()

    Me.Requery
    Me.Refresh

End Sub
