
Sub StockData()
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        ' Last Row Counting
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' Heading
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        ' Variables 
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim ticker_name As String
        Dim percent_change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
        ' Initial Open Price
        open_price = Cells(2, Column + 2).Value
        
        ' Initializing Loop
        For i = 2 To LastRow
         ' Ticker Letter Check
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                ' Ticker Variables
                ticker_var = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = ticker_var
                ' Define Close Price
                close_price = Cells(i, Column + 5).Value
                ' Calculate Yearly Change
                yearly_change = close_price - open_price
                Cells(Row, Column + 9).Value = yearly_change
                ' Calculate Percent Change
                If (open_price = 0 And close_price = 0) Then
                    percent_change = 0
                ElseIf (open_price = 0 And close_price <> 0) Then
                    percent_change = 1
                Else
                    percent_change = yearly_change / open_price
                    Cells(Row, Column + 10).Value = percent_change
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                ' Total Stock Volume
                total_volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = total_volume
                ' Creating new row for table
                Row = Row + 1
                ' Recalc Open Price
                open_price = Cells(i + 1, Column + 2)
                ' Volume Total Calculations
                Volume = 0
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If
        Next i
        
        ' Worksheet Yearly Change
        YCLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        ' RG Coloring
        For j = 2 To YCLastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        ' Determine % Increase 
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
        For Z = 2 To YCLastRow
            If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                Cells(2, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                Cells(3, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                Cells(4, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(Z, Column + 11).Value
            End If
        Next Z
        
    Next WS
        
End Sub
