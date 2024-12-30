Attribute VB_Name = "Module1"
Sub StockAnalysis()
    ' Set Variables
    Dim ws As Worksheet
    Dim i As Long
    
    ' Set variables to hold WorksheetName as a String, LastRow as a Long, Ticker as String
    Dim WorksheetName As String
    Dim LastRow As Long
    Dim ticker As String
                   
    '----------------------------
    ' LOOP THROUGH ALL WORKSHEETS
    '-----------------------------
    For Each ws In ThisWorkbook.Worksheets
    
        ' Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        ' Add Headers to columns
        ws.Cells(1, 11).Value = "Ticker"  ' Column K for ticker
        ws.Cells(1, 12).Value = "Quarterly Change"  ' Column L for Quarterly Change
        ws.Cells(1, 13).Value = "Percentage Change" ' Column M for Percentage Change
        ws.Cells(1, 14).Value = "Total Stock Volume" ' Column N for Total Stock Volume
        ws.Cells(1, 17).Value = "Ticker"  ' Header for displaying Ticker results in Column Q
        ws.Cells(1, 18).Value = "Value" ' Header for displaying Value results in Column R
        ws.Cells(2, 16).Value = "Greatest % Increase" 'Header for displaying Greatest % Increase in Row 2, Column P
        ws.Cells(3, 16).Value = "Greatest % Decrease" ' Header for displaying Greatest % Decrease in Row 3, Column P
        ws.Cells(4, 16).Value = " Greatest Total Volume"  ' Header for displaying Greatest Total Volume in Row 4, Column P
        
        ' Set an initial variable for holding the total stock volume of the stock
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
        
        ' Set variables for previous closing price and quarterly change
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim Quarterly_Change As Double
        Dim Percentage_Change As Double
                     
        ' Set Variables to hold the greatest values
        Dim Greatest_Percentage_Increase As Double
        Dim Greatest_Percentage_Decrease As Double
        Dim Greatest_Total_Volume As Double
        Dim Greatest_Ticker_Increase As String
        Dim Greatest_Ticker_Decrease As String
        Dim Greatest_Ticker_Volume As String
        
        ' Initialize greatest values
        Greatest_Percentage_Increase = 0
        Greatest_Percentage_Decrease = 0
        Greatest_Total_Volume = 0
        
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Grab the WorksheetName
        WorksheetName = ws.Name
    
        ' Loop through all ticker symbols
        For i = 2 To LastRow
            
            ' Check if we are still within the same ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Set the Ticker Count
                ticker = ws.Cells(i, 1).Value
            
                ' Find the opening price for the ticker in the current worksheet
                OpeningPrice = Application.WorksheetFunction.VLookup(ticker, ws.Range("A:C"), 3, False) ' opening price is in column C
                
                ' Find the closing price
                ClosingPrice = ws.Cells(i, "F").Value ' closing price is in column F
                
                ' Calculate Quarterly Change
                Quarterly_Change = ClosingPrice - OpeningPrice
                
                ' Calculate Percentage Change
                If OpeningPrice <> 0 Then
                    Percentage_Change = (Quarterly_Change / OpeningPrice) * 100
                Else
                    Percentage_Change = 0 ' No percentage change for the first entry
                End If

                ' Add to the Total_Stock_Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value ' The stock volumes display in column G

                ' Print the ticker symbol in the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = ticker
             
                ' Print the Quarterly Change to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Quarterly_Change
                
                ' Print the Percentage Change to the Summary Table
                ws.Range("M" & Summary_Table_Row).Value = Percentage_Change & "%"
                
                ' Print the Total_Stock_Volume to the Summary Table
                ws.Range("N" & Summary_Table_Row).Value = Total_Stock_Volume

                ' Highlight Quarterly Change in green if positive and red if negative
                If Quarterly_Change > 0 Then
                    ws.Range("L" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Range("L" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0) ' Red
                End If
                               
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Calculate and store Greatest Total
                If Total_Stock_Volume > Greatest_Total_Volume Then
                    Greatest_Total_Volume = Total_Stock_Volume
                    Greatest_Ticker_Volume = ticker
                End If
            
                'Reset the Total_Stock_Volume count
                Total_Stock_Volume = 0
 
            ' If the cell immediately following a row is the same ticker symbol then
            Else
                ' Add to the Total_Stock_Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
                ' Calculate and store Greatest % Increase
                If Percentage_Change > Greatest_Percentage_Increase Then
                    Greatest_Percentage_Increase = Percentage_Change
                    Greatest_Ticker_Increase = ticker
                End If
            
                ' Calculate and store Greatest % Decrease
                If Percentage_Change < Greatest_Percentage_Decrease Then
                    Greatest_Percentage_Decrease = Percentage_Change
                    Greatest_Ticker_Decrease = ticker
                End If

            End If
        
        Next i
            
        ' Displaying the calculation results in respective rows and columns as define in the earler scripts
        ws.Cells(2, 17).Value = Greatest_Ticker_Increase
        ws.Cells(2, 18).Value = Format(Greatest_Percentage_Increase, "0.00") & "%"
        ws.Cells(3, 17).Value = Greatest_Ticker_Decrease
        ws.Cells(3, 18).Value = Format(Greatest_Percentage_Decrease, "0.00") & "%"
        ws.Cells(4, 17).Value = Greatest_Ticker_Volume
        ws.Cells(4, 18).Value = Format(Greatest_Total_Volume, "0")
            
        'Autofit the columns
        ws.Columns.AutoFit
        
    Next ws

End Sub



