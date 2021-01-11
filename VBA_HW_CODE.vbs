Sub YearStock()

    For Each ws In Worksheets
    
        Dim NxtStk As Integer
        Dim lstRow2 As Long
        Dim Year As Long
        Dim TotalStk As Long
        Dim OpenVal As Double
        Dim OpenVal2 As Double
        Dim CloseVal As Double
        Dim TempVal As Double
        
        TotalStk = 0
        NxtStk = 2
        TempVal = 0
        lstRow2 = ws.Cells(Rows.Count, 1).End(xlUp).Row
        lstCol2 = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        ResRow = ws.Cells(Rows.Count, 1).End(xlUp).Row + 1
        ResCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ws.Cells(1, lstCol2 + 2).Value = "Year"
        ws.Columns("I").ColumnWidth = 10
        ws.Cells(1, lstCol2 + 3).Value = "Ticket"
        ws.Columns("J").ColumnWidth = 10
        ws.Cells(1, lstCol2 + 4).Value = "Yearly_Change"
        ws.Columns("K").ColumnWidth = 15
        ws.Columns("K:K").NumberFormat = "0.00"
        ws.Cells(1, lstCol2 + 5).Value = "Percent Change"
        ws.Columns("L").ColumnWidth = 15
        ws.Columns("L:L").NumberFormat = "0.00%"
        ws.Cells(1, lstCol2 + 6).Value = "Total Stock Vol (M)"
        ws.Columns("M").ColumnWidth = 20
        ws.Columns("M:M").NumberFormat = "#,##0"
        
        
        For i = 2 To lstRow2
        
            If i = 2 Then
                
                TempVal = ws.Cells(2, 3).Value
                
                TmpIni = 2
            
            End If
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Stock = ws.Cells(i, 1).Value
                
                Year = Left((ws.Cells(i, 2).Value), 4)
                
                RngIni2 = ws.Cells(i + 1, 1).Row
                
                If TmpIni > 0 Then
                    
                    RngIni = TmpIni
                    
                    TotalStk = (Application.Sum(Range(ws.Cells(RngIni, lstCol2), ws.Cells(i, lstCol2)))) / 1000
                    
                    ws.Cells(NxtStk, lstCol2 + 6).Value = TotalStk
                    
                    TotalStk = 0
                    
                    TmpIni = 0
                    
                    RngIni = RngIni2
            
                Else
                    
                    TotalStk = (Application.Sum(Range(ws.Cells(RngIni, lstCol2), ws.Cells(i, lstCol2)))) / 1000
                    
                    ws.Cells(NxtStk, lstCol2 + 6).Value = TotalStk
                    
                    TotalStk = 0
                    
                    RngIni = RngIni2
                
                End If
                
                ws.Range("I" & NxtStk).Value = Year
                
                ws.Range("J" & NxtStk).Value = Stock
                
                
                
                OpenVal2 = ws.Cells(i + 1, 3).Value
                        
                CloseVal = ws.Cells(i, 6).Value
                
  
            
                If TempVal > 0 Then
                
                    OpenVal = TempVal
                    
                    ws.Range("K" & NxtStk).Value = CloseVal - OpenVal
                    
                    If ws.Range("K" & NxtStk).Value < 0 Then
                
                        ws.Range("K" & NxtStk).Interior.ColorIndex = 3
                    
                    ElseIf ws.Range("K" & NxtStk).Value > 0 Then
                    
                        ws.Range("K" & NxtStk).Interior.ColorIndex = 4
                        
                    End If
                    
                    If CloseVal = 0 Or OpenVal = 0 Then
                    
                        ws.Range("L" & NxtStk).Value = 0
                    
                    Else
                    
                        ws.Range("L" & NxtStk).Value = (CloseVal / OpenVal) - 1
                    
                    End If
                    
                    TempVal = 0
                    
                    OpenVal = OpenVal2
                
                Else
                
                    ws.Range("K" & NxtStk).Value = CloseVal - OpenVal
                    
                    If ws.Range("K" & NxtStk).Value < 0 Then
                
                        ws.Range("K" & NxtStk).Interior.ColorIndex = 3
                    
                    ElseIf ws.Range("K" & NxtStk).Value > 0 Then
                    
                        ws.Range("K" & NxtStk).Interior.ColorIndex = 4
                        
                    End If
                    
                    
                    If CloseVal = 0 Or OpenVal = 0 Then
                    
                        ws.Range("L" & NxtStk).Value = 0
                    
                    Else
                    
                        ws.Range("L" & NxtStk).Value = (CloseVal / OpenVal) - 1
                    
                    End If
                    
                    OpenVal = OpenVal2
                
                End If
                    
                
                NxtStk = NxtStk + 1
                
                
            End If
            
            
        Next i
        
    Next ws
    
    For Each ws In Worksheets
    
        Dim MaxVol As Long
        Dim Ticket As String
            
        Ticket = ""
        MaxVol = 0
        MaxChng = 0
        MinChng = 0
        FnlCol2 = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        FnlRow2 = ws.Cells(Rows.Count, FnlCol2).End(xlUp).Row
        
        For j = 2 To FnlRow2
             
            If ws.Cells(j, FnlCol2).Value > MaxVol Then
            
                MaxVol = ws.Cells(j, FnlCol2).Value
                
                Ticket = ws.Cells(j, FnlCol2 - 3).Value
                
            End If
            
            If ws.Cells(j, FnlCol2 - 1).Value > MaxChng Then
            
                MaxChng = ws.Cells(j, FnlCol2 - 1).Value
                
                Ticket2 = ws.Cells(j, FnlCol2 - 3).Value
                
            End If
            
            If ws.Cells(j, FnlCol2 - 1).Value < MinChng Then
            
                MinChng = ws.Cells(j, FnlCol2 - 1).Value
                
                Ticket3 = ws.Cells(j, FnlCol2 - 3).Value
                
            End If
                        
        Next j
        
        ws.Cells(2, FnlCol2 + 3).Value = "Ticket"
        ws.Cells(2, FnlCol2 + 4).Value = "Value"
        
        ws.Columns("O").ColumnWidth = 20
        ws.Cells(3, FnlCol2 + 2).Value = "Greatest Total Vol"
        ws.Cells(4, FnlCol2 + 2).Value = "Greatest % Increase"
        ws.Cells(5, FnlCol2 + 2).Value = "Greatest % Decrease"
        
        ws.Cells(3, FnlCol2 + 3).Value = Ticket
        ws.Cells(3, FnlCol2 + 4).Value = MaxVol
        ws.Cells(3, FnlCol2 + 4).NumberFormat = "#,##0"
        
        ws.Cells(4, FnlCol2 + 3).Value = Ticket2
        ws.Cells(4, FnlCol2 + 4).Value = MaxChng
        ws.Cells(4, FnlCol2 + 4).NumberFormat = "0.00%"
        
        ws.Cells(5, FnlCol2 + 3).Value = Ticket3
        ws.Cells(5, FnlCol2 + 4).Value = MinChng
        ws.Cells(5, FnlCol2 + 4).NumberFormat = "0.00%"
    
    Next ws

End Sub
