Attribute VB_Name = "Vba_Reclamations_SSD"
Option Explicit
'***************************************************************************************************************************************
'
'Code pour envoyer des réclamations
'
'***************************************************************************************************************************************

Sub Send_Reclamation()
Attribute Send_Reclamation.VB_ProcData.VB_Invoke_Func = "M\n14"
 
    Dim OutApp As Object
    Dim OutMail As Object
    Dim Cell As Range
 
    Application.ScreenUpdating = True
    Set OutApp = CreateObject("Outlook.Application")
 
    On Error GoTo cleanup
    For Each Cell In Columns("A").Cells.SpecialCells(xlCellTypeConstants)
        If Cell.Value Like "?*@?*.?*" And _
            LCase(Cells(Cell.Row, "F").Value) = "oui" Then
 
            Set OutMail = OutApp.CreateItem(0)
            On Error Resume Next
            With OutMail
                .To = Cell.Value
                .Subject = Cells(Cell.Row, "B").Value '  "Payslip"
                .Body = "Bonjour " & Cells(Cell.Row, "A").Value _
                      & vbNewLine & vbNewLine & _
                        "Voir pièce jointe. " & _
                    vbNewLine & vbNewLine & _
                        "Au revoir."
 
                .Attachments.Add (Cells(Cell.Row, "E").Value) ' ("C:\000 Payslips\" & Cells(cell.Row, "A").Value & ".pdf")
                '.Display  'à remplacer par
                .Send
                
            End With
            On Error GoTo 0
            Set OutMail = Nothing
        End If
    Next Cell
 
cleanup:
    Set OutApp = Nothing
    Application.ScreenUpdating = True
End Sub
