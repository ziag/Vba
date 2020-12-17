Attribute VB_Name = "Vba_Macro_Generique"
Option Explicit

Public Sub FormatRapport() 'Touche de raccourci du clavier: Ctrl+Shift+W
Attribute FormatRapport.VB_ProcData.VB_Invoke_Func = "W\n14"
 
    Selection.AutoFilter
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    Cells.Select
    Cells.EntireColumn.AutoFit
    Rows("1:1").Select
    
End Sub

Sub PointPourVirgule()  'Touche de raccourci du clavier: Ctrl+Shift+E

    Cells.Replace What:=".", Replacement:=",", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=True
        
End Sub

Sub FormatClasseur()
Attribute FormatClasseur.VB_ProcData.VB_Invoke_Func = "E\n14"
Dim Wk As Workbook
Dim i As Integer

    Set Wk = ActiveWorkbook
    
    For i = 1 To Wk.Sheets.Count
        
        Worksheets(Wk.Sheets(i).Name).Activate
        Call FormatRapport(Wk.Sheets(i).Name)
    Next i
    
    Worksheets(1).Activate

End Sub

'Sub FindValueInColFromRange(colRange As String, colToCheck As String)
'
'    Dim Cell As Range
'    Columns("A:A").Select
'    Set Cell = Selection.Find(What:="celda", After:=ActiveCell, LookIn:=xlFormulas, _
'            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
'            MatchCase:=False, SearchFormat:=False)
'
'    If Cell Is Nothing Then
'        'do it something
'
'    Else
'        'do it another thing
'    End If
'End Sub

Private Sub DupliquerFeuille()
    Dim sh As Worksheet, sh2 As Worksheet
    Set sh = ActiveSheet
    For Each sh2 In Worksheets
        If sh2.Name = [B3] Then
            Sheets([B3].Value).Activate
            If MsgBox("Feuille " & [B3] & " existante. La supprimer ?", vbCritical + vbYesNo) = vbNo Then
                sh.Activate
                Exit Sub
            Else
                Application.DisplayAlerts = False
                Sheets([B3].Value).Delete
                Application.DisplayAlerts = True
                Exit For
            End If
        End If
    Next sh2
    sh.Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = [B3]
    sh.Activate
End Sub

Private Sub CreateSheetsFromAList()
Attribute CreateSheetsFromAList.VB_ProcData.VB_Invoke_Func = "l\n14"
    Dim MyCell As Range, MyRange As Range
    
    Set MyRange = Sheets("Data").Range("H2")
    Set MyRange = Range(MyRange, MyRange.End(xlDown))

    For Each MyCell In MyRange
        Sheets.Add After:=Sheets(Sheets.Count) 'creates a new worksheet
        Sheets(Sheets.Count).Name = MyCell.Value ' renames the new worksheet
    Next MyCell
End Sub

Sub LoopRange()

    Dim RngAncien As Range
    Dim RngNouv As Range
    Dim RngValidation As Range
    Dim sh As Worksheet
    
    Dim rAncien As Variant
    Dim rNouv As Variant
    Dim i As Integer
    
    Set sh = ActiveSheet
    
    Dim lastRow_1 As Long
    Dim lastRow_2 As Long
    lastRow_1 = sh.Range("A" & Rows.Count).End(xlUp).Row
    lastRow_2 = sh.Range("D" & Rows.Count).End(xlUp).Row
   
    Set RngNouv = sh.Range("A2:A" & lastRow_1)
    Set RngAncien = sh.Range("D2:D" & lastRow_2)
    Set RngValidation = sh.Range("B2:B" & lastRow_2)
    Debug.Print Time
    
    For Each rNouv In RngNouv.Cells
        i = 0
        For Each rAncien In RngAncien.Cells
            i = i + 1
            If rAncien.Value <> "" And rNouv.Value <> "" Then
                If rAncien.Value = rNouv.Value Then
                    RngValidation.Cells(i, 1).Value = 1
                ElseIf RngValidation.Cells(i, 1).Value = "" Then
                    RngValidation.Cells(i, 1).Value = 0
                End If
            Else
                Exit For
            End If
       Next rAncien
    Next rNouv
    
    Debug.Print Time
End Sub
