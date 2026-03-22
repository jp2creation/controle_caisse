Option Explicit

Private Const TEMPLATE_NAME As String = "MODELE_JOUR"
Private Const CONFIG_SHEET_NAME As String = "CONFIG"
Private Const FIRST_COUNT_ROW As Long = 47
Private Const LAST_COUNT_ROW As Long = 61
Private Const PREVIOUS_DAY_COLUMN As String = "E"
Private Const DAY_CASH_COLUMN As String = "F"
Private Const BANK_DEPOSIT_COLUMN As String = "G"

Private Function FeuilleExiste(ByVal nomFeuille As String) As Boolean
    On Error Resume Next
    FeuilleExiste = Not ThisWorkbook.Worksheets(nomFeuille) Is Nothing
    On Error GoTo 0
End Function

Private Function DerniereDateCaisse() As Date
    Dim ws As Worksheet
    Dim nomFeuille As String
    Dim dateFeuille As Date
    Dim meilleureDate As Date

    meilleureDate = 0

    For Each ws In ThisWorkbook.Worksheets
        nomFeuille = Trim$(ws.Name)

        If Len(nomFeuille) = 8 And IsNumeric(nomFeuille) Then
            On Error Resume Next
            dateFeuille = DateSerial(CInt(Right$(nomFeuille, 4)), CInt(Mid$(nomFeuille, 3, 2)), CInt(Left$(nomFeuille, 2)))

            If Err.Number = 0 Then
                If Format$(dateFeuille, "ddmmyyyy") = nomFeuille Then
                    If dateFeuille > meilleureDate Then
                        meilleureDate = dateFeuille
                    End If
                End If
            End If

            Err.Clear
            On Error GoTo 0
        End If
    Next ws

    DerniereDateCaisse = meilleureDate
End Function

Private Function DateSuivanteCaisse() As Date
    Dim datePrecedente As Date

    datePrecedente = DerniereDateCaisse()

    If datePrecedente = 0 Then
        DateSuivanteCaisse = Date
    Else
        DateSuivanteCaisse = datePrecedente + 1
    End If
End Function

Private Function DemanderDatePersonnalisee(ByVal dateProposee As Date) As Date
    Dim saisie As String

    Do
        saisie = InputBox( _
            "Saisis la date voulue (jj/mm/aaaa)." & vbCrLf & _
            "Laisse vide puis OK pour annuler.", _
            "Modifier la date", _
            Format$(dateProposee, "dd/mm/yyyy"))

        If saisie = vbNullString Then
            DemanderDatePersonnalisee = 0
            Exit Function
        End If

        If IsDate(saisie) Then
            DemanderDatePersonnalisee = CDate(saisie)
            Exit Function
        End If

        MsgBox "Date invalide. Exemple : 14/03/2026", vbExclamation
    Loop
End Function

Public Function EstFeuilleJourCaisse(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    EstFeuilleJourCaisse = (DateFeuilleCaisse(ws.Name) <> 0)
End Function

Private Function DateFeuilleCaisse(ByVal nomFeuille As String) As Date
    On Error GoTo Fin

    If Len(nomFeuille) <> 8 Or Not IsNumeric(nomFeuille) Then Exit Function

    DateFeuilleCaisse = DateSerial(CInt(Right$(nomFeuille, 4)), CInt(Mid$(nomFeuille, 3, 2)), CInt(Left$(nomFeuille, 2)))
    If Format$(DateFeuilleCaisse, "ddmmyyyy") <> nomFeuille Then
        DateFeuilleCaisse = 0
    End If

Fin:
End Function

Private Sub RemplirCaisseVeille(ByVal wsPrev As Worksheet, ByVal wsNext As Worksheet)
    Dim cible As Range
    Dim etaitProtegee As Boolean
    Dim ligne As Long

    Set cible = wsNext.Range(PREVIOUS_DAY_COLUMN & FIRST_COUNT_ROW & ":" & PREVIOUS_DAY_COLUMN & LAST_COUNT_ROW)
    etaitProtegee = wsNext.ProtectContents Or wsNext.ProtectDrawingObjects Or wsNext.ProtectScenarios

    On Error GoTo Sortie

    If etaitProtegee Then
        wsNext.Unprotect
    End If

    For ligne = FIRST_COUNT_ROW To LAST_COUNT_ROW
        wsNext.Range(PREVIOUS_DAY_COLUMN & ligne).Formula = _
            "='" & wsPrev.Name & "'!" & DAY_CASH_COLUMN & ligne & _
            "-'" & wsPrev.Name & "'!" & BANK_DEPOSIT_COLUMN & ligne
    Next ligne

Sortie:
    If etaitProtegee Then
        On Error Resume Next
        wsNext.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        On Error GoTo 0
    End If
End Sub

Private Function FeuillesJourTriees() As Collection
    Dim feuilles As New Collection
    Dim ws As Worksheet
    Dim i As Long
    Dim inseree As Boolean

    For Each ws In ThisWorkbook.Worksheets
        If EstFeuilleJourCaisse(ws) Then
            inseree = False

            For i = 1 To feuilles.Count
                If DateFeuilleCaisse(ws.Name) < DateFeuilleCaisse(feuilles(i).Name) Then
                    feuilles.Add ws, Before:=i
                    inseree = True
                    Exit For
                End If
            Next i

            If Not inseree Then
                feuilles.Add ws
            End If
        End If
    Next ws

    Set FeuillesJourTriees = feuilles
End Function

Public Sub SynchroniserReportsCaisses(Optional ByVal nomFeuilleSource As String = vbNullString)
    Dim feuilles As Collection
    Dim i As Long
    Dim indexDepart As Long
    Dim previousEnableEvents As Boolean
    Dim previousScreenUpdating As Boolean

    previousEnableEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating

    On Error GoTo Sortie

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set feuilles = FeuillesJourTriees()
    If feuilles.Count < 2 Then GoTo Sortie

    indexDepart = 2

    If Len(nomFeuilleSource) > 0 Then
        For i = 1 To feuilles.Count
            If feuilles(i).Name = nomFeuilleSource Then
                indexDepart = i + 1
                Exit For
            End If
        Next i
    End If

    If indexDepart > feuilles.Count Then GoTo Sortie

    For i = indexDepart To feuilles.Count
        RemplirCaisseVeille feuilles(i - 1), feuilles(i)
        feuilles(i).Calculate
    Next i

Sortie:
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEnableEvents
End Sub

Public Sub ActualiserCaisses(Optional ByVal nomFeuilleSource As String = vbNullString)
    SynchroniserReportsCaisses nomFeuilleSource
    Application.Calculate
    MajCouleurOngletsP3
End Sub

Public Sub NouveauJourCaisse()
    Dim wsTemplate As Worksheet
    Dim wsNew As Worksheet
    Dim dNew As Date
    Dim newName As String
    Dim rep As VbMsgBoxResult
    Dim previousCalculation As XlCalculation
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim previousDisplayAlerts As Boolean

    previousCalculation = Application.Calculation
    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    previousDisplayAlerts = Application.DisplayAlerts

    On Error GoTo GestionErreur

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    If Not FeuilleExiste(TEMPLATE_NAME) Then
        MsgBox "La feuille modele '" & TEMPLATE_NAME & "' est introuvable.", vbCritical
        GoTo SortiePropre
    End If

    If Not FeuilleExiste(CONFIG_SHEET_NAME) Then
        MsgBox "La feuille '" & CONFIG_SHEET_NAME & "' est introuvable.", vbCritical
        GoTo SortiePropre
    End If

    dNew = DateSuivanteCaisse()

    rep = MsgBox( _
        "Creer la nouvelle feuille pour le " & Format$(dNew, "dd/mm/yyyy") & " ?" & vbCrLf & vbCrLf & _
        "Oui = utiliser cette date" & vbCrLf & _
        "Non = choisir une autre date" & vbCrLf & _
        "Annuler = abandonner", _
        vbYesNoCancel + vbQuestion, _
        "Nouveau jour de caisse")

    If rep = vbCancel Then
        GoTo SortiePropre
    End If

    If rep = vbNo Then
        dNew = DemanderDatePersonnalisee(dNew)
        If dNew = 0 Then
            GoTo SortiePropre
        End If
    End If

    newName = Format$(dNew, "ddmmyyyy")

    If FeuilleExiste(newName) Then
        MsgBox "La feuille " & newName & " existe deja.", vbExclamation
        GoTo SortiePropre
    End If

    Set wsTemplate = ThisWorkbook.Worksheets(TEMPLATE_NAME)
    wsTemplate.Copy After:=ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    Set wsNew = ActiveSheet

    On Error Resume Next
    wsNew.Name = newName
    If Err.Number <> 0 Then
        MsgBox "Impossible de renommer la nouvelle feuille en " & newName & ".", vbCritical
        Err.Clear
        On Error GoTo GestionErreur
        GoTo SortiePropre
    End If
    On Error GoTo GestionErreur

    On Error Resume Next
    wsNew.Unprotect
    On Error GoTo GestionErreur

    wsNew.Range("D3").Value = dNew

    wsNew.Range("C11:L40").ClearContents
    wsNew.Range("O11:Q40").ClearContents
    wsNew.Range("M47:M49").ClearContents
    wsNew.Range("F47:G61").ClearContents
    wsNew.Range("F47:G61").Value = 0
    wsNew.Range("E47:E61").ClearContents
    wsNew.Range("E47:E61").Value = 0

    wsNew.Calculate
    ActualiserCaisses

    On Error Resume Next
    wsNew.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    On Error GoTo GestionErreur

    wsNew.Activate
    wsNew.Range("C11").Select

    MsgBox "Nouvelle caisse : " & newName, vbInformation, "Caisse"

SortiePropre:
    Application.Calculation = previousCalculation
    Application.DisplayAlerts = previousDisplayAlerts
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

GestionErreur:
    MsgBox "Erreur pendant la creation du nouveau jour : " & Err.Description, vbCritical, "Caisse"
    Resume SortiePropre
End Sub
