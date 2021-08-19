Attribute VB_Name = "Module1"
Sub Multi_Stock_Data_Analysis()

Dim WS_Count As Integer
Dim I As Long
Dim Ticker As String
Dim Year_Opening_Price As Double
Dim Year_Closing_Price As Double
Dim Year_Price_Change As Double
Dim Year_Percent_Change As Double
Dim Total_Stock_Volume As LongLong
Dim Summary_Table_Row As Integer
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Last_Row As Long
Dim Get_Once As Boolean
Dim WS As Worksheet

'Loop through each of the sheets
For Each WS In Worksheets

'Insert Summary Table Headers
WS.Cells(1, 9).Value = "Ticker"
WS.Cells(1, 10).Value = "Yearly Change"
WS.Cells(1, 11).Value = "Year Percent Price"
WS.Cells(1, 12).Value = "Total Stock Volume"

' Keep track of the last row in each sheet
Last_Row = WS.Range("A" & Rows.Count).End(xlUp).Row

' Keep track of the location for each stock ticker symbol in the Summary Table
Summary_Table_Row = 2
'Initialise the Total_Stock_Volume
Total_Stock_Volume = 0
  ' Part 1: Loop through all stocks and create the Summary Table
  For I = 2 To Last_Row

' Check if we are still within the same stock ticker symbol, if it is not...
    If WS.Cells(I, 1).Value <> WS.Cells(I + 1, 1).Value Then

' Set the Stock Ticker Symbol
      Ticker = WS.Cells(I, 1).Value
' Print the Stock Ticker Symbol in the Summary Table
      WS.Range("I" & Summary_Table_Row).Value = Ticker
      
      'Set the Closing Price
      Year_Closing_Price = WS.Cells(I, 6).Value
      
      ' Set the Yearly Change
      Year_Price_Change = Year_Closing_Price - Year_Opening_Price
      
      ' Print the Stock Price Difference in the Summary Table
      WS.Range("J" & Summary_Table_Row).Value = Year_Price_Change
          
         ' Add Conditional Formatting to the the Stock Price Difference
      If WS.Range("J" & Summary_Table_Row).Value >= 0 Then
         WS.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
      Else
         WS.Range("J" & Summary_Table_Row).Interior.Color = vbRed
      End If
      
      'Increase the Total_Stock_Volume by the final amount
      Total_Stock_Volume = Total_Stock_Volume + WS.Cells(I, 7).Value
      
      'Print the Total Stock Volume in the Summary Table
      WS.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
      
      'Reset the Total Stock Volume
      Total_Stock_Volume = 0
      
      'Set the Percent Change
      If (Year_Opening_Price = 0) Then
      Year_Percent_Change = 0
      Else
      Year_Percent_Change = Year_Price_Change / Year_Opening_Price
      End If
      
      'Print the Percent Change in the Summary Table
      WS.Range("K" & Summary_Table_Row).Value = Year_Percent_Change
      
      ' Reformat the Percent Change in the Summary Table
      WS.Range("K" & Summary_Table_Row) = Format(WS.Range("K" & Summary_Table_Row).Value, "Percent")
      
      ' Add one to the Summary Table row
      Summary_Table_Row = Summary_Table_Row + 1
    
    Get_Once = False
    ' If the cell immediately following a row is the same ticker symbol...
    Else
      If Get_Once = False Then
         ' Only get a Stock's Opening Price once
         Year_Opening_Price = WS.Cells(I, 3).Value
         Get_Once = True
      End If
      
      'Increase the Total_Stock_Volume
      Total_Stock_Volume = Total_Stock_Volume + WS.Cells(I, 7).Value
      

    End If

  Next I
  
 'Part 2: Loop through the Summary Table to create a second Summary Table
 'Insert column labels
 WS.Range("P1").Value = "Ticker"
 WS.Range("Q1").Value = "Value"
 ' Insert row labels
 WS.Range("O2").Value = "Greatest % Increase"
 WS.Range("O3").Value = "Greatest % Decrease"
 WS.Range("O4").Value = "Greatest Total Volume"
 
 'Get the second Summary Table Range
  Year_Percent_Price_Range = WS.Range(WS.Cells(2, 11), WS.Cells((Summary_Table_Row - 1), 11))
  Total_Stock_Volume_Range = WS.Range(WS.Cells(2, 12), WS.Cells((Summary_Table_Row - 1), 12))
  
 'Find the the Greatest Increase in value in the Summary Table
 Greatest_Increase = WS.Application.WorksheetFunction.Max(Year_Percent_Price_Range)
 WS.Range("Q2").Value = Greatest_Increase
 WS.Range("Q2") = Format(WS.Range("Q2").Value, "Percent")
 
'Find the ticker symbol of the stock with the Greatest Increase in the Summary Table
 Greatest_Increase_Row_Number = WS.Application.Match(Greatest_Increase, Year_Percent_Price_Range, 0) + 1
 WS.Range("P2").Value = WS.Range("I" & Greatest_Increase_Row_Number)

 'Find the Greatest Decrease
 Greatest_Decrease = WS.Application.WorksheetFunction.Min(Year_Percent_Price_Range)
 WS.Range("Q3").Value = Greatest_Decrease
 WS.Range("Q3") = Format(WS.Range("Q3").Value, "Percent")
 
 'Find the ticker of the stock with the Greatest Increase in the Summary Table
 Greatest_Decrease_Row_Number = WS.Application.Match(Greatest_Decrease, Year_Percent_Price_Range, 0) + 1
 WS.Range("P3").Value = WS.Range("I" & Greatest_Decrease_Row_Number)

 'Find the Greatest Total Volume in the Summary Table
 Greatest_Total_Volume = WS.Application.WorksheetFunction.Max(Total_Stock_Volume_Range)
 WS.Range("Q4").Value = Greatest_Total_Volume
 WS.Range("Q4").NumberFormat = "0.0000E+00"

'Find the total volume of the stock with the greatest increase
 Greatest_Total_Volume_Row_Number = WS.Application.Match(Greatest_Total_Volume, Total_Stock_Volume_Range, 0) + 1
 WS.Range("P4").Value = WS.Range("I" & Greatest_Total_Volume_Row_Number)


'Readjust the width of all the columns
WS.Columns("I:Q").AutoFit

Next WS


End Sub
