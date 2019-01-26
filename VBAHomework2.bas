Attribute VB_Name = "Module1"
Sub check_each_sheet()

    'Loop through all sheets
    For Each Stock_Year In Worksheets()
    
        Stock_Year.Select
        set_header
        seperate_stock_ticker
        conditional_format
        spotlight_values
    
    Next Stock_Year

End Sub

Sub set_header()
 
    'Set up titles in header I=ticker, J= yearly change, K = percent change, L = total stock volume
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    

End Sub

Sub conditional_format()

    'Fit Columns to Header
    Range("I:L").EntireColumn.AutoFit 'Columns I to L
    
    
    'Variable to determine last row of summary table
    LastRow = Range("I1", Range("I1").End(xlDown)).Rows.Count
      
    ' Loop through Summary Table
    For ticker = 2 To LastRow
        yearly_change = Cells(ticker, "J").Value
        
        'Set Percent style for collum 11 ("K")
        Cells(ticker, 11).Style = "Percent"
        
        'Set color for Yearly Change based of positive or negative
        If yearly_change < 0 Then
            'RED for negative yearly_change values
            Cells(ticker, "J").Interior.ColorIndex = 3
        Else
            'GREEN for non negetive yearly_change values
            Cells(ticker, "J").Interior.ColorIndex = 4
        
        End If
        
    Next ticker
            
End Sub

Sub spotlight_values()
'Your solution will also be able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".
'Columns O,P, Q
    
    'Set up table
    Cells(2, "O").Value = "Greatest % Increase"
    Cells(3, "O").Value = "Greatest % Decrease"
    Cells(4, "O").Value = "Greatest Total Volume"
    Cells(1, "P").Value = "Ticker"
    Cells(1, "Q").Value = "Value"
    
    'get initial values for variables
    high_percent = Range("J2").Value
    high_ticker = Range("I2").Value
    low_percent = Range("J2").Value
    low_ticker = Range("I2").Value
    volume_ticker = Range("I2").Value
    high_volume = Range("L2").Value
    
    'Variable to determine last row of summary table
    LastRow = Range("I1", Range("I1").End(xlDown)).Rows.Count
      
    ' Loop through Summary Table
    For ticker = 3 To LastRow
        'Get greatest percentage increase
        If Cells(ticker, "J").Value > high_percent Then
            high_percent = Cells(ticker, "J").Value
            high_ticker = Cells(ticker, "I").Value
        End If
        'Get greatest percentage decrease
        If Cells(ticker, "J").Value < low_percent Then
            low_percent = Cells(ticker, "J").Value
            low_ticker = Cells(ticker, "I").Value
        End If
        'Get largest volume traded
        If Cells(ticker, "L").Value > high_volume Then
            high_volume = Cells(ticker, "L").Value
            volume_ticker = Cells(ticker, "I").Value
        End If
    
    Next ticker
    
    Range("P2").Value = high_ticker
    Range("Q2").Value = high_percent
    
    Range("P3").Value = low_ticker
    Range("Q3").Value = low_percent
    
    Range("P4").Value = volume_ticker
    Range("Q4").Value = high_volume
    
    'Fit Columns
    Range("O:Q").EntireColumn.AutoFit 'Columns I to L


End Sub


Sub seperate_stock_ticker()

    ' Set an initial variables
    Dim Stock_Symbol As String
    
    ' Set an initial variable for holding the total volume for stock
    Dim Volume_Total As LongLong
    Volume_Total = 0
    
    ' Keep track of the location for each stock(by ticker) in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'Row count for current stock ticker being counted
    Dim Row_Count As Integer
    Row_Count = 0
    
    'Variable to determine last row on each tab
    'LastRow = Cells(Row.Count, 1).End(xIUp).Row
    LastRow = Range("A1", Range("A1").End(xlDown)).Rows.Count
      
    ' Loop through all trading information
    'Loop to iterate over table
    For I = 2 To LastRow

        ' Check if we are still within the same stock symbol
        'When stock symbol changes information is printed to summary table
        'New Stock Symbol is set
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

            ' Set the Stock name and Close_Price
            Stock_Symbol = Cells(I, 1).Value
            Close_Price = Cells(I, 6).Value

            'Calculate Yearly_Change (new-old) and (new - old/old) of price
            yearly_change = (Close_Price - Open_Price)
            'Error handeling if statment
            If Open_Price <> 0 Then
                Percent_Change = (Close_Price - Open_Price) / Open_Price
            Else
                Percent_Change = (Close_Price - Open_Price) * 100
            End If
            ' Add to the Volume Total
            Volume_Total = Volume_Total + Cells(I, 7).Value

            ' Print the Stock Symbol in the Summary Table
            Range("I" & Summary_Table_Row).Value = Stock_Symbol
            
            ' Print the Yearly_Change in the Summary Table
            Range("J" & Summary_Table_Row).Value = yearly_change

            ' Print the Percent_Change in the Summary Table
            Range("K" & Summary_Table_Row).Value = Percent_Change
    
            ' Print the Volume_Total to the Summary Table
            Range("L" & Summary_Table_Row).Value = Volume_Total

            ' Interate 1 to summary table row
            Summary_Table_Row = Summary_Table_Row + 1
          
            ' Reset the Volume_Total and Row_Count
            Volume_Total = 0
            Row_Count = 0
    ' If the cell immediately following a row is the same ticker symbol...
        Else
            If Row_Count = 0 Then
                'Here you grad the opening price
                Open_Price = Cells(I, 3).Value
                
            End If

        ' Add to the Volume Total
        Volume_Total = Volume_Total + Cells(I, 7).Value
        
        'Iterate Row counter
        Row_Count = Row_Count + 1

        
        End If

    Next I

End Sub
