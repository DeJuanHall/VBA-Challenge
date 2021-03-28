Attribute VB_Name = "Module2"
Sub StockSummarySingleWS()


    'Create and declare variables
Dim i As Long
Dim Ticker_Symbol As String
Dim Yearly_Open_Value As Double
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0
Dim Yearly_Change As Double
Yearly_Change = 0
Dim Percent_Change As Double
Dim Stock_table As Long
Stock_table = 2
Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    
    'Loop all Ticker Symbols
For i = 2 To LastRow
Yearly_Open_Value = Cells(Stock_table, 3).Value
    
    'Add if statement
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'set values 'Add Ticker, Yearly Change, Percent Change and formatting to Summary

        Range("I1").Value = "Ticker_Symbol"
        Ticker_Symbol = Cells(i, 1).Value
        Range("I" & Stock_table).Value = Ticker_Symbol
        
        Range("J1").Value = "Yearly_Change"
        Yearly_Change = Yearly_Change + (Cells(i, 6).Value - Yearly_Open_Value)
        Range("J" & Stock_table).Value = Yearly_Change
        If Range("J" & Stock_table).Value > 0 Then Range("J" & Stock_table).Interior.ColorIndex = 4
        If Range("J" & Stock_table).Value < 0 Then Range("J" & Stock_table).Interior.ColorIndex = 3
        
        Range("k1").Value = "Percent_Change"
        Percent_Change = (Yearly_Change / Yearly_Open_Value)
        Range("K" & Stock_table).Value = Percent_Change
        Range("K" & Stock_table).NumberFormat = "0.00%"
        
        Range("L1").Value = "Total_Stock_Volume"
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        Range("L" & Stock_table).Value = Total_Stock_Volume

    
    'Reset
        Stock_table = Stock_table + 1
        Yearly_Change = 0
        Total_Stock_Volume = 0
        Open_Price = Cells(Stock_table, 3).Value
    
    'look for next same ticker in next row
    Else
    
    'Add to ticker total
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    End If
Next i



End Sub

