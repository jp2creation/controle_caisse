Option Explicit

Private Const STATUS_CELL As String = "P3"
Private Const TEMPLATE_NAME As String = "MODELE_JOUR"

Private Function EstFeuilleCaisse(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    EstFeuilleCaisse = (ws.Name = TEMPLATE_NAME) Or EstFeuilleJourCaisse(ws)
End Function

Private Function EstFeuilleJourCaisse(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    EstFeuilleJourCaisse = (Len(ws.Name) = 8 And IsNumeric(ws.Name))
End Function

Private Function DateFeuilleCaisse(ByVal ws As Worksheet) As Date
    If Not EstFeuilleJourCaisse(ws) Then Exit Function

    On Error Resume Next
    DateFeuilleCaisse = DateSerial(CInt(Right$(ws.Name, 4)), CInt(Mid$(ws.Name, 3, 2)), CInt(Left$(ws.Name, 2)))

    If Err.Number <> 0 Then
        DateFeuilleCaisse = 0
    ElseIf Format$(DateFeuilleCaisse, "ddmmyyyy") <> ws.Name Then
        DateFeuilleCaisse = 0
    End If

    Err.Clear
    On Error GoTo 0
End Function

Private Function FeuilleSansActivite(ByVal ws As Worksheet) As Boolean
    FeuilleSansActivite = (Application.CountA(ws.Range("F11:L40")) = 0) _
                       And (Application.CountA(ws.Range("O11:P40")) = 0)
End Function

Private Function ValeurNumerique(ByVal valeur As Variant) As Double
    If IsError(valeur) Or IsEmpty(valeur) Then Exit Function

    If VarType(valeur) = vbString Then
        If Trim$(CStr(valeur)) = vbNullString Then Exit Function
    End If

    If IsNumeric(valeur) Then
        ValeurNumerique = CDbl(valeur)
    End If
End Function

Private Function ProblemePrincipal(ByVal ws As Worksheet, ByVal statutCellule As String) As String
    If UCase$(Trim$(statutCellule)) = "OK" Then
        ProblemePrincipal = "RAS"
        Exit Function
    End If

    If ValeurNumerique(ws.Range("M3").Value) <> 0 Then
        ProblemePrincipal = "Ecart de saisie"
        Exit Function
    End If

    If ValeurNumerique(ws.Range("M60").Value) <> 0 Then
        ProblemePrincipal = "Ecart de caisse"
        Exit Function
    End If

    If Application.CountIf(ws.Range("M11:M40"), "ERREUR") > 0 Then
        ProblemePrincipal = "Lignes facture"
        Exit Function
    End If

    If Application.CountIf(ws.Range("N47:N50"), "<>OK") > 0 Then
        ProblemePrincipal = "Controle banque"
        Exit Function
    End If

    ProblemePrincipal = "A verifier"
End Function

Private Function StatutReelCaisse(ByVal ws As Worksheet) As String
    Dim statutCellule As String
    Dim probleme As String
    Dim dateFeuille As Date

    statutCellule = UCase$(Trim$(CStr(ws.Range(STATUS_CELL).Value)))

    If statutCellule = "OK" Or statutCellule = "ERREUR" Then
        StatutReelCaisse = statutCellule
        Exit Function
    End If

    probleme = ProblemePrincipal(ws, statutCellule)

    If FeuilleSansActivite(ws) Then
        If EstFeuilleJourCaisse(ws) Then
            dateFeuille = DateFeuilleCaisse(ws)

            If dateFeuille <> 0 And dateFeuille < Date And probleme <> "A verifier" Then
                StatutReelCaisse = "ERREUR"
            End If
        End If

        Exit Function
    End If

    If probleme = "RAS" Then
        StatutReelCaisse = "OK"
    Else
        StatutReelCaisse = "ERREUR"
    End If
End Function

Public Sub MajCouleurOngletFeuilleP3(ByVal ws As Worksheet)
    Dim statut As String

    If ws Is Nothing Then Exit Sub
    If Not EstFeuilleCaisse(ws) Then Exit Sub

    On Error GoTo Fin

    statut = StatutReelCaisse(ws)

    Select Case statut
        Case "OK"
            ws.Tab.Color = RGB(0, 176, 80)
        Case "ERREUR"
            ws.Tab.Color = RGB(192, 0, 0)
        Case Else
            ws.Tab.ColorIndex = xlColorIndexNone
    End Select

Fin:
End Sub

Public Sub MajCouleurOngletsP3()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        MajCouleurOngletFeuilleP3 ws
    Next ws
End Sub
