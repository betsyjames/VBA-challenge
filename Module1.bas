Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()

    'Loop through all sheets
     For Each ws In ActiveWorkbook.Worksheets
     ws.Activate

        'Set variable for ticker
        Dim Ticker As String

        'Set variable for total volume
        Dim Total_Volume As Double
        Total_Volume = 0

        'Set varible for Open_Price,Close_Price,Yearly Change , Percentage Change
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        

        'Keep track of the row to display summary values
        Dim Summary_Row As Double
        Summary_Row = 2

        'Find the Last Row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Set Inital Open_Price
        Open_Price = Cells(2, 3).Value
        
        'Set Headings for Summary Display
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
    
    
        'Loop through all rows
        For I = 2 To lastrow
            'Check if we are still within the same ticker, if it is not
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
     
                'Set the ticker
                 Ticker = Cells(I, 1).Value
     
                'Save the Opening Price
                 Close_Price = Cells(I, 6).Value
     
                'Add the Volume total
                 Total_Volume = Total_Volume + Cells(I, 7).Value
                 
                'Calculate Yearly_Change
                 Yearly_Change = Close_Price - Open_Price
                 
                 'Calculate Percentage_Change
                 If (Open_Price = 0 Or Close_Price = 0) Then
                    Percent_Change = 0
                 ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = Close_Price
                 Else
                    Percent_Change = Yearly_Change / Open_Price
                 End If
                 
                 'Print the ticker in column
                  Range("I" & Summary_Row).Value = Ticker
                 'Print the Yearly Change in column
                  Range("J" & Summary_Row).Value = Yearly_Change
                 'Print the Percent Change
                  Range("K" & Summary_Row).Value = Percent_Change
                 'Format the percent value to include 2 decimals and percent sign
                  Range("K" & Summary_Row).NumberFormat = "0.00%"
                 'Print the total Volume
                  Range("L" & Summary_Row).Value = Total_Volume
     
                'Add one to the summary table
                 Summary_Row = Summary_Row + 1
     
                'Reset the Open Price
                 Open_Price = Cells(I + 1, 3)
                'Reset the Volumn Total
                 Total_Volume = 0
     
            Else
   
                Total_Volume = Total_Volume + Cells(I, 7).Value
     
            End If
     
        Next I
        
        ' Find the Last Row of Yearly Change per WS
        Yearlastrow = ws.Cells(Rows.Count, "J").End(xlUp).Row
        ' Set the Cell Colors
        For j = 2 To Yearlastrow
            If (Cells(j, "J").Value > 0 Or Cells(j, "J").Value = 0) Then
                Cells(j, "J").Interior.ColorIndex = 4
            ElseIf Cells(j, "J").Value < 0 Then
                Cells(j, "J").Interior.ColorIndex = 3
            End If
        Next j
        
        'Bonus
        'Set Headings for Greatest % Increase,Decrease Total Volume, Ticker and Value
        
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        'Loop through each row of summary
        For k = 2 To Yearlastrow
            'Find the max value in column Percent Change
            If (Cells(k, "K").Value = WorksheetFunction.Max(Range("K2:K" & Yearlastrow))) Then 'Need to fix range
                'Print the Ticker
                 Cells(2, "P").Value = Cells(k, "I").Value
                 Cells(2, "Q").Value = Cells(k, "K").Value
                 Cells(2, "Q").NumberFormat = "0.00%"
                 
            'Find the min value in column Percent Change
            ElseIf (Cells(k, "K").Value = WorksheetFunction.Min(Range("K2:K" & Yearlastrow))) Then 'Need to fix range
                'Print the Ticker
                 Cells(3, "P").Value = Cells(k, "I").Value
                 Cells(3, "Q").Value = Cells(k, "K").Value
                 Cells(3, "Q").NumberFormat = "0.00%"
                 
            'Find the max volume from column Total Stock Volume
            ElseIf (Cells(k, "L").Value = WorksheetFunction.Max(Range("L2:L" & Yearlastrow))) Then 'Need to fix range
                'Print the Ticker
                 Cells(4, "P").Value = Cells(k, "I").Value
                 Cells(4, "Q").Value = Cells(k, "L").Value
            End If
         
        
        Next k
        
     Next ws
     
End Sub

