Attribute VB_Name = "Vba_Envoi_Courriels"
Option Explicit
'***************************************************************************************************************************************
'
'Code pour vérifier liste d'envoi courriel UDA
'
'***************************************************************************************************************************************
Public Sub VerifListeCourrielUDA()
Dim vue As Integer
Dim exclusion  As Integer
Dim listMbr As Integer
Dim listStag As Integer
Dim listSansEmail As Integer
Dim listExclusNonTrouve As Integer
Dim msg As String


    vue = countRowFeuille("Data") - 1 'sans entete
    exclusion = CountUnique("Exclus", "D")
    listMbr = countRowFeuille("ListeDesMembres") - 1 'sans entete
    listStag = countRowFeuille("ListeDesStagiaires") - 1 'sans entete
    listSansEmail = countRowFeuille("ListeSansCourriel") - 1 'sans entete
    listExclusNonTrouve = countRowFeuille("ListeDesExclusNonTrouvé") - 1 'sans entete
        
    msg = vue & " - (" & exclusion & " + " & listMbr & " + " & listStag & " + " & listSansEmail & " - " & listExclusNonTrouve & ") = " & _
          vue - (exclusion + listMbr + listStag + listSansEmail - listExclusNonTrouve)
                                         
    MsgBox (msg)
End Sub

Private Function CountUnique(nomFeuille As String, colLettre As String) As Long  'dataRange As Range
Dim lastRow As Long
Dim Rng As Range
Dim CheckCell
Dim Counter As Double

    Counter = 0
    Worksheets(nomFeuille).Activate
    lastRow = ActiveSheet.Range(colLettre & Rows.Count).End(xlUp).Row
    Set Rng = ActiveSheet.Range(colLettre & "2:" & colLettre & lastRow)
    
    For Each CheckCell In Rng.Cells
        Counter = Counter + (1 / (WorksheetFunction.CountIf(Rng, CheckCell.Value)))
    Next
    
    CountUnique = Counter
End Function

Private Function countRowFeuille(nomFeuille As String) As Long
Dim Wb As Workbook
Dim Ws As Worksheet
        
    With Sheets(nomFeuille)
        countRowFeuille = .Range("A" & .Rows.Count).End(xlUp).Row
    End With

End Function


