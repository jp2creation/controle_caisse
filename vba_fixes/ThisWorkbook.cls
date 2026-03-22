Option Explicit

Private Sub Workbook_Open()
    MajCouleurOngletsP3
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If TypeOf Sh Is Worksheet Then
        MajCouleurOngletFeuilleP3 Sh
    End If
End Sub
