Attribute VB_Name = "Vba_Rapport_SX_Mandats_Claims"
Option Explicit
'***************************************************************************************************************************************
'
'Code pour exécuter le rapport SoundExchange Mandates et Claims
'
'***************************************************************************************************************************************

Private Const SVR_CON = "Provider=MSOLEDBSQL;Server=ARTISTI-SQL;Database=Sigart_rapport;Trusted_Connection=yes;DataTypeCompatibility=80"
Private Const NBR_DE_ID = 101 'Doit etre plus grand que un(1) vu que le premiere ligne est le nom des colonnes :-).
'***************************************************************************************************************************************
'***************************************************************************************************************************************
'***************************************************************************************************************************************
Public Type Sigart
     sigartID As Variant
     AvecDisco As Boolean
     inClaim As Boolean
 End Type
    
Public Sub Rapport_SX_Mandate_Claim()
Dim IdRequete As Variant
Dim IdClaim As Variant

 Dim Claim() As Sigart
 Dim Mandate() As Sigart
    
    'vérifie si les feuilles existent si oui delete
    Call EffaceFeuille("Claim")
    Call EffaceFeuille("Mandate")
        
    Call Set_PaysExclus("united states") 'met à jour les exclus de représentation au USA

    Claim = GetIdSig("A")
    Call Get_ClaimSig(Claim)
        
    Mandate = GetIdSig("L")
    Call MergeStruct(Mandate, Claim)
    
     
   Call Get_MandateSig(Mandate)
    
End Sub

Private Sub FindAExclure(couleur As String)  'colLettre1 As String, colLettre2 As String,
Dim i As Integer
Dim j As Integer
Dim colNbr1 As Long
Dim colNbr2 As Long
 
    i = 0
    j = 0
    colNbr1 = 3 ' ColLettre2Chiffre(colLettre1)
    colNbr2 = 4 ' ColLettre2Chiffre(colLettre2)
    
    For i = 1 To Range("A1").End(xlDown).Row  ' - 1
        j = i + 1
    
        If Cells(i, colNbr1).Value = False And Cells(i, colNbr2).Value = False Then
            Cells(i, 1).Value = ""
            Range("A" & i & ":Q" & i).Select
        
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = couleur
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With

        End If
    
    Next i
        
    Rows("1:1").Select

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

Private Sub MergeStruct(ByRef Mandate() As Sigart, Claim() As Sigart)
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim tmpMandate As Integer
Dim topLimite As Integer

    tmpMandate = UBound(Mandate)
    
    For k = 1 To tmpMandate
        Mandate(k).inClaim = True
    Next k
    
    For i = 1 To UBound(Claim)
        For j = 1 To tmpMandate
        
            If Claim(i).sigartID = Mandate(j).sigartID Then
                 Mandate(j).AvecDisco = Claim(i).AvecDisco
                 Exit For
            ElseIf j = tmpMandate Then
                 Debug.Print ("M " & Claim(i).sigartID)
                 topLimite = UBound(Mandate) + 1
                 ReDim Preserve Mandate(topLimite)
                 Mandate(topLimite) = Claim(i)
            End If
            
        Next j
    Next i
    
    
    
End Sub


'********************************Procedure*******************************************************
Private Sub Affiche_RecordSet(rs As ADODB.Recordset, NomPage As String)
Dim Wb As Workbook
Dim Ws As Worksheet
Dim i As Integer
Dim Exists As Boolean
Dim DerniereLigne As Long
        
   If rs.EOF = False Then
       Set Wb = ActiveWorkbook
       Exists = WorksheetExists(NomPage)
       
       If Exists = False Then
           Set Ws = Wb.Worksheets.Add
           
           Ws.Name = NomPage
           Ws.Activate
           Ws.Select
          
           For i = 0 To rs.Fields.Count - 1
               Ws.Cells(1, i + 1).Value = rs.Fields(i).Name
           Next i
           
       Else
           Set Ws = Wb.Sheets(NomPage)
           
       End If
    
       DerniereLigne = (Ws.Cells(Ws.Rows.Count, "A").End(xlUp).Row) + 1
       Ws.Range("A" & DerniereLigne).CopyFromRecordset rs
    End If
    
End Sub

Private Sub AjoutClaimDisco(Sigart() As Sigart)
Dim i As Integer
Dim j As Integer
Dim limite As Long
     
    limite = GetNbrLigne
    
    For i = 1 To limite
        For j = 0 To UBound(Sigart)
           If Cells(i, 8).Value = Sigart(j).sigartID Then
                Cells(i, 3).Value = Sigart(j).AvecDisco
                Cells(i, 4).Value = Sigart(j).inClaim
                Exit For
           End If
           
        Next j
    Next i
       
End Sub
Private Sub AjouteCouleur(couleur As String, colDebut As Integer, colFin As Integer)
Dim i As Integer

    For i = colDebut To colFin
        Columns(i).Interior.Color = couleur
    Next i
    
End Sub

Private Sub DeleteRows(ColDoublonUn As Integer, Optional ColDoublonDeux As Integer)
Dim MyRange As Range
Dim lastRow As Long
Dim LastCol As Long
Dim colLettre As String
    
    lastRow = GetNbrLigne()
    lastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    LastCol = ActiveSheet.Range("A1").CurrentRegion.Columns.Count
    colLettre = Split(Cells(1, LastCol).Address, "$")(1)
     
    Set MyRange = ActiveSheet.Range("A1:" & colLettre & lastRow)
    
    If ColDoublonDeux <> 0 Then
        MyRange.RemoveDuplicates Columns:=Array(ColDoublonUn, ColDoublonDeux), Header:=xlYes
    Else
        MyRange.RemoveDuplicates Columns:=ColDoublonUn, Header:=xlYes
    End If
    
End Sub

Private Sub EffaceFeuille(NomPage As String)
Dim Exists As Boolean
Dim Wb As Workbook
Set Wb = ActiveWorkbook
    
    Exists = WorksheetExists(NomPage)
    
    If Exists = True Then
        Application.DisplayAlerts = False
        Wb.Worksheets(NomPage).Delete
        Application.DisplayAlerts = True
    End If
        
End Sub

Private Sub FormatRapportSX()

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

'Private Sub Get_Claim(TabloSigartId As Variant)
'Dim NumColDoublon_1 As Integer
'Dim NumColDoublon_2 As Integer
'
'    If IsArray(TabloSigartId) Then
'        Call GetData(TabloSigartId, "stp_Get_Sx_Claim", "Claim")
'        NumColDoublon_1 = GetColIndexFromColName("RECORDING-LOCAL-ID-CLAIMING-SOCIETY")
'        NumColDoublon_2 = GetColIndexFromColName("RIGHT-HOLDER-LOCAL-ID-CLAIMING-SOCIETY")
'        Call DeleteRows(NumColDoublon_1, NumColDoublon_2)
'        Call AjouteCouleur(vbRed, 1, 4)
'        Call FormatRapportSX
'    End If
'
'End Sub

Private Sub Get_ClaimSig(ByRef sig() As Sigart)
Dim cn As ADODB.Connection
Dim NumColDoublon_1 As Integer
Dim NumColDoublon_2 As Integer
Dim i As Integer
Dim sigartID As Double
Dim rs As Recordset

    If UBound(sig) > 0 Then
        
        Set cn = New ADODB.Connection
        cn.ConnectionString = SVR_CON
        cn.Open
        
        For i = 1 To UBound(sig)  ' 0= la ligne des noms de colones.
            If IsNumeric(sig(i).sigartID) And sig(i).sigartID <> "" Then
                sigartID = sig(i).sigartID
                
                'Call GetDataSig(sig(), "stp_Get_Sx_Claim", "Claim")
                Call GetDataSig(cn, sigartID, "stp_Get_Sx_Claim", rs)
              
                If rs.EOF = False Then
                    sig(i).AvecDisco = True
                    sig(i).inClaim = True
                End If
                
                Call Affiche_RecordSet(rs, "Claim")
                rs.Close
        
            End If
        Next
     
     cn.Close
     
  
    
     NumColDoublon_1 = GetColIndexFromColName("RECORDING-LOCAL-ID-CLAIMING-SOCIETY")
     NumColDoublon_2 = GetColIndexFromColName("RIGHT-HOLDER-LOCAL-ID-CLAIMING-SOCIETY")
     Call DeleteRows(NumColDoublon_1, NumColDoublon_2)
     Call AjouteCouleur(vbRed, 1, 4)
     Call FormatRapportSX
    
    End If
    
End Sub

'Private Sub GetData(TabloSigartId As Variant, Sp As String, NomPage As String)
'Dim Cn As ADODB.Connection
'Dim Cmd As ADODB.Command
'Dim rs As ADODB.Recordset
'Dim i As Integer
'
'
'    Set Cn = New ADODB.Connection
'    Cn.ConnectionString = SVR_CON
'    Cn.Open
'
'    Set Cmd = New ADODB.Command
'    Cmd.ActiveConnection = Cn
'    Cmd.CommandType = adCmdStoredProc
'    Cmd.CommandText = Sp
'
'    For i = 1 To UBound(TabloSigartId)  ' 0= la ligne des noms de colones.
'        'If IsNumeric(TabloSigartId(i, 1)) And TabloSigartId(i, 1) <> "" Then
'        If IsNumeric(TabloSigartId(i)) And TabloSigartId(i) <> "" Then
'            Cmd.Parameters.Refresh
'            'Cmd.Parameters("@sigartID").Value = TabloSigartId(i, 1)
'            Cmd.Parameters("@sigartID").Value = TabloSigartId(i)
'            Set rs = Cmd.Execute
'            Call Affiche_RecordSet(rs, NomPage)
'            rs.Close
'        End If
'    Next i
'
'    Cn.Close
'
'End Sub

'Private Sub GetDataSig(ByRef sig() As Sigart, Sp As String, NomPage As String)

Private Sub GetDataSig(cn As ADODB.Connection, sigartID As Double, Sp As String, ByRef rs As Recordset)
Dim Cmd As ADODB.Command
'Dim rs As ADODB.Recordset
'Dim i As Integer

    Set Cmd = New ADODB.Command
    Cmd.ActiveConnection = cn
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = Sp

'    For i = 1 To UBound(sig)  ' 0= la ligne des noms de colones.
'        If IsNumeric(sig(i).SigartID) And sig(i).SigartID <> "" Then
    Cmd.Parameters.Refresh
            'Cmd.Parameters("@sigartID").Value = sig(i).SigartID
    Cmd.Parameters("@sigartID").Value = sigartID
    Set rs = Cmd.Execute
            
            
'            If rs.EOF = False Then
'                sig(i).AvecDisco = True
'            End If
'            Call Affiche_RecordSet(rs, NomPage)
'            rs.Close
'        End If
'    Next i
    
'    Set GetDataSig = rs
'    Cn.Close
    
End Sub
'Private Sub Get_Mandate(TabloSigartId As Variant)
'Dim NumColDoublon As Integer
'
'    If IsArray(TabloSigartId) Then
'        Call GetData(TabloSigartId, "stp_get_SX_Mandate", "Mandate")
'        NumColDoublon = GetColIndexFromColName("Performer_Local_ID")
'        Call DeleteRows(NumColDoublon)
'        Call AjouteCouleur(vbRed, 1, 4)
'        Call FormatRapportSX
'    End If
'
'End Sub

Private Sub Get_MandateSig(ByRef sig() As Sigart)
Dim cn As ADODB.Connection
Dim NumColDoublon As Integer
Dim i As Integer
Dim sigartID As Double
Dim rs As Recordset

    If UBound(sig) > 0 Then
        
        Set cn = New ADODB.Connection
        cn.ConnectionString = SVR_CON
        cn.Open
        
        For i = UBound(sig) To 1 Step -1   ' 0= la ligne des noms de colones.
            If IsNumeric(sig(i).sigartID) And sig(i).sigartID <> "" Then
                sigartID = sig(i).sigartID
                
                Call GetDataSig(cn, sigartID, "stp_get_SX_Mandate", rs)
                Call Affiche_RecordSet(rs, "Mandate")
                rs.Close
        
            End If
        Next
        
        cn.Close
    End If
    
    NumColDoublon = GetColIndexFromColName("Performer_Local_ID")
    Call DeleteRows(NumColDoublon)
    Call AjouteCouleur(vbRed, 1, 4)
    Call AjoutClaimDisco(sig())
    Call FindAExclure(vbYellow)
    Call FormatRapportSX
   
    
    
    
'Dim NumColDoublon As Integer
'Dim rs As Recordset
'
'    If UBound(sig) > 1 Then
'        For i = 1 To UBound(sig)  ' 0= la ligne des noms de colones.
'            If IsNumeric(sig(i).sigartID) And sig(i).sigartID <> "" Then
'
'                'Call GetDataSig(sig(), "stp_get_SX_Mandate", "Mandate")
'                rs = GetDataSig(sig(i).sigartID, "stp_get_SX_Mandate")
'
'
'
'                If rs.EOF = False Then sig(i).AvecDisco = True
''            End If
'            Call Affiche_RecordSet(rs, NomPage)
'            rs.Close
'
'            NumColDoublon = GetColIndexFromColName("Performer_Local_ID")
'            Call DeleteRows(NumColDoublon)
'            Call AjouteCouleur(vbRed, 1, 4)
'            'Call FormatRapportSX
'
'            End If
'        Next
'    End If
    
End Sub

Private Sub Set_PaysExclus(paysEN As String)
Dim cn As ADODB.Connection
Dim Cmd As ADODB.Command

    Set cn = New ADODB.Connection
    cn.ConnectionString = SVR_CON
    cn.Open

    Set Cmd = New ADODB.Command
    Cmd.ActiveConnection = cn
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "stp_FeedTerritoireExclus"
    
    Cmd.Parameters.Refresh
    Cmd.Parameters("@paysEN").Value = paysEN
    Cmd.Execute

    cn.Close

End Sub

'********************* function **********************
Function EnleverDoublon(Tablo As Variant) As Variant
Dim dic As Object
Dim arrItem As Variant

    If IsArray(Tablo) Then
        
        Set dic = CreateObject("Scripting.Dictionary")
        
        For Each arrItem In Tablo
            If Not dic.Exists(arrItem) Then
                dic.Add arrItem, arrItem
            End If
        Next
        EnleverDoublon = dic.Keys
    Else
        EnleverDoublon = Tablo
    End If
    
End Function

Private Function GetColIndexFromColName(nomColonne As String) As Integer
Dim strSearch As String
Dim aCell As Range
 
    Set aCell = ActiveSheet.Rows(1).Find(What:=nomColonne, LookIn:=xlValues, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
   
     GetColIndexFromColName = aCell.Column
     
End Function

Private Function GetNbrLigne() As Long
On Error Resume Next

    GetNbrLigne = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row

On Error GoTo 0
End Function

'Private Function GetId(colonne As String) As Variant()
'Dim lastRow As Long
'Dim Tablo As Variant
'
'    lastRow = GetNbrLigne()
'    Tablo = Range(colonne & "1:" & colonne & lastRow)
'
'    GetId = EnleverDoublon(Tablo)
'
'End Function

Private Function GetIdSig(colonne As String) As Sigart()      'feuil As String,  ByRef s() As Sigart
Dim lastRow As Long
Dim TabloRange As Variant
Dim TabloSansDoublon As Variant
Dim SigartInfo() As Sigart
Dim i As Integer

    lastRow = GetNbrLigne()
    TabloRange = Range(colonne & "1:" & colonne & lastRow)
    TabloSansDoublon = EnleverDoublon(TabloRange)
        
    If IsArray(TabloRange) Then
        ReDim Preserve SigartInfo(UBound(TabloSansDoublon))
        
        For i = 0 To UBound(TabloSansDoublon)
            SigartInfo(i).sigartID = TabloSansDoublon(i)
        Next
    End If
    
    GetIdSig = SigartInfo()
    
End Function

'Private Function ColLettre2Chiffre(colLettre As String)
'
'   ColLettre2Chiffre = Range(colLettre & 1).Column
'
'End Function
 

Private Function WorksheetExists(WorksheetName As String) As Boolean

On Error Resume Next
    WorksheetExists = (ActiveWorkbook.Sheets(WorksheetName).Name <> "")
On Error GoTo 0

End Function
