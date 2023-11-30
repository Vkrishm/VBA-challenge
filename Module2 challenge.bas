Attribute VB_Name = "Module1"

Sub Stock_Data_Analysis()

'Set variable for Ticker
Dim Ticker As String

'Set variable for total stock volume
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0
  
'Set variable for year open
Dim Year_Open As Double

'Set variable for year close
Dim Year_Close As Double

'Set variable for yearly close
Dim Yearly_Change As Double

'Set variable for percentage_change
Dim Percentage_Change As Double

'Set variable to set up a row to start to print
Dim Summary_Table_Row As Integer

'Define variable of the worksheet
Dim ws As Worksheet


For Each ws In Worksheets

'For printing on corresponding row
Summary_Table_Row = 2

'Used in calculating Yearly change and Percentage change
Start = 2

    'Assign column header

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ' Loop through all Stock transactions
    Lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row


  For i = 2 To Lastrow

    'Check if we are still within the same Ticker, if it is not...
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    'Set the Ticker name
      Ticker = ws.Cells(i, 1).Value

    'Add to the Total stock volume
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

    'Print the Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker

    'Print the Total Stock volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

    'Reset the Total stock volume
      Total_Stock_Volume = 0
      
    'Calculate the yearly and percentage change
      Yearly_Change = (ws.Cells(i, 6).Value - ws.Cells(Start, 3).Value)
      Percentage_Change = Yearly_Change / ws.Cells(Start, 3).Value
    
    'Start of the next ticker
      Start = i + 1
     
    'Print the results
     ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
     ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
    
    'Change Column J to percentage format
     ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

    'Add one to the summary table row
     Summary_Table_Row = Summary_Table_Row + 1
      
    'If the cell immediately following a row is the same Ticker...
     Else

    'Add to the Total stock volume
     Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

    End If

  Next i
  
    'Conditional formatting columns colors

    'Find the Last row for column J
     jLastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
    
        For j = 2 To jLastRow

   'if value is > or < zero
     If ws.Cells(j, 10) > 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
        Else
        ws.Cells(j, 10).Interior.ColorIndex = 3
        End If

        Next j

  'Bonus table
  
  'Assign column and Row header
  ws.Range("O2").Value = "Greatest % Increase"
  ws.Range("O3").Value = "Greatest % Decrease"
  ws.Range("O4").Value = "Greatest Total Volume"
  ws.Range("P1").Value = "Ticker"
  ws.Range("Q1").Value = "Value"
  
  'Find the Last row in K
   kLastRow = ws.Cells(Rows.Count, "K").End(xlUp).Row
   
  'Finding the highest increase  and decrease in percentage percentage
    
    'Declaring the initial values
    ws.Range("Q2").Value = 0
    ws.Range("Q3").Value = 0
    ws.Range("Q4").Value = 0

    For i = 2 To kLastRow
    'Finding greatest percentage increase
    If ws.Cells(i, 11).Value > 0 And ws.Cells(i, 11).Value > ws.Range("Q2").Value Then
    ws.Range("Q2").Value = ws.Cells(i, 11).Value
    ws.Range("P2").Value = ws.Cells(i, 9).Value

End If

    'Finding greatest percentage decrease
    If ws.Cells(i, 11).Value < 0 And ws.Cells(i, 11).Value < ws.Range("Q3").Value Then
    ws.Range("Q3").Value = ws.Cells(i, 11).Value
    ws.Range("P3").Value = ws.Cells(i, 9).Value

End If
    'Finding greatest total stock volume
    If ws.Cells(i, 12).Value > ws.Range("Q4").Value Then
    ws.Range("P4").Value = ws.Cells(i, 9).Value
    ws.Range("Q4").Value = ws.Cells(i, 12).Value
End If


    'Change Q2 and Q3 to percentage format
     ws.Range("Q2:Q3").NumberFormat = "0.00%"
   
  Next
  
    'Auto adjust column width
    ws.Columns("A:Q").AutoFit

  Next ws

End Sub











