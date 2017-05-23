

Function CalcComm(SaleAmount As Double, saleDate As Date) As Double
    
    Dim CommRate As Double

    
    If SaleAmount < 10000 Then
        CommRate = 0.03
    ElseIf SaleAmount < 25000 Then
        CommRate = 0.04
    Else
        CommRate = 0.06
    End If

    If (Month(saleDate) = 12) Or _
        (Month(saleDate) = 1) Or _
        (Month(saleDate) = 2) Then
        
        CommRate = CommRate * 1.5
    End If
    
    CalcComm = SaleAmount * CommRate
        
End Function


Sub SubFillComm_Range()

   Dim i As Integer
   For i = 5 To 1004
        Range("E" & i).Value = CalcComm(Range("D" & i).Value, Range("C" & i).Value)
    
   Next i

End Sub

Sub SubFillComm_cells()

   Dim i As Integer
   For i = 5 To 1004
        Cells(i, 5).Value = CalcComm(Cells(i, 4).Value, Cells(i, 3).Value)
        
   Next i

End Sub

Function ShowQuarter(saleDate As Date) As String

    Dim SaleMonth As Integer
    
    SaleMonth = Month(saleDate)
    
    If SaleMonth < 4 Then
        ShowQuarter = "Q1"
    ElseIf SaleMonth < 7 Then
        ShowQuarter = "Q2"
    ElseIf SaleMonth < 10 Then
        ShowQuarter = "Q3"
    Else
        ShowQuarter = "Q4"
    End If
    
    
End Function

