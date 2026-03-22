Option Explicit

Public Sub AjouterJustificatifSortie()
    Dim ws As Worksheet
    Dim cibleLigne As Long
    Dim nomClient As String
    Dim montant As Variant
    Dim fd As FileDialog
    Dim sourceFichier As String
    Dim dossierBase As String
    Dim dossierJour As String
    Dim nomFichier As String
    Dim destination As String

    On Error GoTo GestionErreur

    Set ws = ActiveSheet

    If Not EstFeuilleJourCaisse(ws) Then
        MsgBox "Ouvre d'abord une feuille de caisse datee (jjmmaaaa).", vbExclamation, "Justificatif"
        Exit Sub
    End If

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Enregistre d'abord le classeur avant d'ajouter un justificatif.", vbExclamation, "Justificatif"
        Exit Sub
    End If

    cibleLigne = ActiveCell.Row
    If cibleLigne < 11 Or cibleLigne > 40 Then
        MsgBox "Selectionne une cellule sur une ligne de SORTIES (lignes 11 a 40).", vbExclamation, "Justificatif"
        Exit Sub
    End If

    nomClient = Trim$(CStr(ws.Cells(cibleLigne, "O").Value))
    montant = ws.Cells(cibleLigne, "P").Value

    If nomClient = vbNullString And (IsEmpty(montant) Or montant = vbNullString) Then
        MsgBox "La ligne de sortie est vide. Selectionne une ligne renseignee.", vbExclamation, "Justificatif"
        Exit Sub
    End If

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "Choisir un scan/photo de facture de sortie"
        .Filters.Clear
        .Filters.Add "Fichiers acceptes", "*.pdf;*.jpg;*.jpeg;*.png;*.webp"

        If .Show <> -1 Then
            Exit Sub
        End If

        sourceFichier = .SelectedItems(1)
    End With

    dossierBase = ThisWorkbook.Path & Application.PathSeparator & "Justificatifs_Sorties"
    dossierJour = dossierBase & Application.PathSeparator & ws.Name

    EnsureFolder dossierBase
    EnsureFolder dossierJour

    nomFichier = ConstruireNomFichierUnique(dossierJour, sourceFichier, cibleLigne)
    destination = dossierJour & Application.PathSeparator & nomFichier

    FileCopy sourceFichier, destination

    ws.Cells(10, "Q").Value = "JUSTIFICATIF"
    ws.Columns("Q").ColumnWidth = 18

    On Error Resume Next
    ws.Cells(cibleLigne, "Q").Hyperlinks.Delete
    On Error GoTo GestionErreur

    ws.Hyperlinks.Add Anchor:=ws.Cells(cibleLigne, "Q"), Address:=destination, TextToDisplay:="Ouvrir"

    MsgBox "Justificatif ajoute pour la ligne " & cibleLigne & ".", vbInformation, "Justificatif"
    Exit Sub

GestionErreur:
    MsgBox "Impossible d'ajouter le justificatif : " & Err.Description, vbExclamation, "Justificatif"
End Sub

Private Function EstFeuilleJourCaisse(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    EstFeuilleJourCaisse = (Len(ws.Name) = 8 And IsNumeric(ws.Name))
End Function

Private Function ConstruireNomFichierUnique(ByVal dossierJour As String, ByVal sourceFichier As String, ByVal cibleLigne As Long) As String
    Dim extension As String
    Dim baseNom As String
    Dim nomFichier As String
    Dim index As Long

    If InStrRev(sourceFichier, ".") > 0 Then
        extension = Mid$(sourceFichier, InStrRev(sourceFichier, "."))
    End If

    baseNom = Format$(Now, "yyyymmdd_hhnnss") & "_L" & Format$(cibleLigne, "00")
    nomFichier = baseNom & extension
    index = 1

    Do While Len(Dir$(dossierJour & Application.PathSeparator & nomFichier)) > 0
        nomFichier = baseNom & "_" & index & extension
        index = index + 1
    Loop

    ConstruireNomFichierUnique = nomFichier
End Function

Private Sub EnsureFolder(ByVal chemin As String)
    Dim parentPath As String
    Dim separatorPos As Long

    If Len(chemin) = 0 Then Exit Sub
    If Len(Dir$(chemin, vbDirectory)) <> 0 Then Exit Sub

    separatorPos = InStrRev(chemin, Application.PathSeparator)
    If separatorPos > 0 Then
        parentPath = Left$(chemin, separatorPos - 1)
        If Len(parentPath) > 0 And Len(Dir$(parentPath, vbDirectory)) = 0 Then
            EnsureFolder parentPath
        End If
    End If

    MkDir chemin
End Sub
