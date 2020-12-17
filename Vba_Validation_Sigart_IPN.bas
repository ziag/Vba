Attribute VB_Name = "Vba_Validation_Sigart_IPN"
Option Explicit
'***************************************************************************************************************************************
'
'Code pour exécuter le rapport de validation SIGART/IPN
'
'***************************************************************************************************************************************
Public Sub TrouveDifference_SIGART_IPN()

    
    Call EffaceValidation
    Call FormatRapportIPN

    '***************Vérification a partir du SigartID
    Call TriColonneAvecEntete("B")
    Call FindVarianceByColLettre("B", vbYellow)

    '***************Vérification a partir du IPN
    Call TriColonneAvecEntete("A")
    Call FindVarianceByColLettre("A", vbGreen)



    Call EffaceCouleurColonne("I")
    
    '************trouve les orphelins****************
    Call GetUniqueAndCount("A", vbRed)
    Call GetUniqueAndCount("B", vbRed)
    Call CouleurRow("A", "B")
    Call EffaceRegleColonne("A")
    Call EffaceRegleColonne("B")
    
    Call FocusSurPremiereLigne
    
End Sub

Private Sub TriColonneAvecEntete(colonne As String)
On Error Resume Next
Dim lastRow As Long

    lastRow = ActiveSheet.Range(colonne & Rows.Count).End(xlUp).Row

    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add2 _
        Key:=Range(colonne & "1:" & colonne & lastRow), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
        
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
  
End Sub

Private Sub FormatRapportIPN()
   
    If ActiveSheet.AutoFilterMode Then
        'si le filtre existe on fait rien
    Else
        Range("A1").Select
        Selection.AutoFilter
    End If
    
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    
    
End Sub
 
Private Sub GetUniqueAndCount(colonne As String, couleur As String)
Dim lastRow As Long
Dim Rng As Range
Dim rCell As Range
Dim v As Variant

    lastRow = ActiveSheet.Range(colonne & Rows.Count).End(xlUp).Row - 1
    
    Set Rng = Range(colonne & "2:" & colonne & lastRow)
   
    Rng.FormatConditions.Delete
    
    Dim uniqueVals As UniqueValues
    Set uniqueVals = Rng.FormatConditions.AddUniqueValues
    uniqueVals.DupeUnique = xlUnique
    uniqueVals.Interior.Color = couleur
    
End Sub

Private Sub FindVarianceByColLettre(colLettre As String, couleur As String)
Dim i As Integer
Dim j As Integer
Dim colNbr As Long
 
    i = 0
    j = 0
    colNbr = ColLettre2Chiffre(colLettre)
    
    For i = 1 To Range("A1").End(xlDown).Row - 1
        j = i + 1
    
        If Cells(i, colNbr).Value = Cells(j, colNbr).Value Then
            
            Rows(i & ":" & j).ColumnDifferences(Range("A" & j)).Offset(1, 0).Select
            
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = couleur
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        
        End If
    
    Next i

End Sub

Private Sub CouleurRow(colUnLettre As String, colDeuxLettre As String)
Dim colUn As Long
Dim colDeux As Long
Dim i As Integer
Dim LastCol As Long
Dim colLettre As String
Dim estMemeCouleur As Boolean
    
    i = 0
    colUn = ColLettre2Chiffre(colUnLettre)
    colDeux = ColLettre2Chiffre(colDeuxLettre)
    
    LastCol = ActiveSheet.Range("A1").CurrentRegion.Columns.Count
    colLettre = Split(Cells(1, LastCol).Address, "$")(1)


    For i = 1 To Range("A1").End(xlDown).Row - 1
    
        estMemeCouleur = ComparerCouleur(Range("A" & i), Range("B" & i))
        
        If estMemeCouleur = True Then
        
            Range("A" & i & ":" & colLettre & i).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = vbRed
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
    
    Next i
    
End Sub

Private Function ColLettre2Chiffre(colLettre As String)

   ColLettre2Chiffre = Range(colLettre & 1).Column

End Function

Public Sub EffaceValidation()
Attribute EffaceValidation.VB_ProcData.VB_Invoke_Func = "e\n14"
On Error Resume Next

    If ActiveSheet.AutoFilterMode Then
        Selection.AutoFilter = False
    End If
    
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Cells.Select
    
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Cells.FormatConditions.Delete
    
    Call FocusSurPremiereLigne
    
End Sub

Private Sub EffaceCouleurColonne(colonne As String)

    Columns(colonne & ":" & colonne).Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("A1").Select
    
End Sub

Private Function ComparerCouleur(rColor As Range, rRange As Range) As Boolean
Dim vResult As Boolean
    
    vResult = False
 
    If rColor.DisplayFormat.Interior.Color = vbRed And _
        rRange.DisplayFormat.Interior.Color = rColor.DisplayFormat.Interior.Color _
         Then
            vResult = True
    End If
    
    ComparerCouleur = vResult
    
End Function

Private Sub EffaceRegleColonne(colLettre As String)

    Columns(colLettre).Select
    Selection.FormatConditions.Delete
    
End Sub

Private Sub FocusSurPremiereLigne()

    Range("A2").Select
    Range("1:1").Select
    
End Sub

 
'
'Private Function ColChiffre2Lettre(colNombre As Long)
'
'    ColChiffre2Lettre = Split(Cells(1, ColumnNumber).Address, "$")(1)
'
'End Function
