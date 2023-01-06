Sub StockReview()

For Each ws In Worksheets

    'Defining last row
    Dim lastrow As Long
     lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
    'Defining Summary_Table_Row
    Dim Summary_Table_Row As Integer
     Summary_Table_Row = 2
     
    'Defining Ticker
    Dim Ticker As String
    ws.Range("I1,P1") = "Ticker"
    
    'Defining Yearly_Change
    Dim Yearly_Change As Double
    ws.Range("J1") = "Yearly Change"
    
    'Defining Percent_Change
    Dim Percent_Change As Double
    ws.Range("K1") = "Percent Change"
    
    'defining Total_Stock_Volume
    Dim Total_Stock_Volume As Double
    ws.Range("L1") = "Total Stock Volume"
    
    'defining Open_Price
    Dim Open_Price As Double
    'set Open_Price
    Open_Price = ws.Range("C2").Value
    
             'Add Functionality
    'Defining Greatest_Increase
    Dim Greatest_Increase As Double
    Greatest_Increase = ws.Cells(2, 11).Value
    ws.Range("O2") = "Greatest % Increase"
    
    'Defining Greatest_Decrease
    Dim Greatest_Decrease As Double
    Greatest_Decrease = ws.Cells(2, 11).Value
    ws.Range("O3") = "Greatest % Decrease"
    
    'Define Greatest_Total_Volume
    ws.Range("O4") = "Greatest Total Volume"
    Dim Greatest_Total_Volume As Double
    Greatest_Total_Volume = 0
    
    'column name
    ws.Range("Q1") = "Value"
    
    'Define Greatest_Increase_Ticker
    Dim Greatest_Increase_Ticker As String
    
    'Define Greatest_Decrease_Ticker
    Dim Greates_Increase_Ticker As String
    
    
            'Displaying Data
    
    For i = 2 To lastrow
    
      If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
         
         'set the Ticker name
         Ticker = ws.Cells(i, 1).Value
         'set Yearly_Change
         Yearly_Change = ws.Cells(i, 6) - Open_Price
         'set Percent Change
         Percent_Change = (ws.Cells(i, 6) / Open_Price) - 1
         'add last row to the Total_Stock_Volume count
         Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
         
         
         'print the Ticker name in the result column
         ws.Range("I" & Summary_Table_Row).Value = Ticker
         'print Yearly_Cange in the result column
         ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
         'set number format for Yearly_Change
         ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
         'print Percent_Change
         ws.Range("K" & Summary_Table_Row).Value = Percent_Change
         'set number format for Percent_Change
         ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
         'print Total Stock Volume
         ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
         
         'add one to the Summary_Table_Row
         Summary_Table_Row = Summary_Table_Row + 1
         'Reset Open_Price
         Open_Price = ws.Cells(i + 1, 3).Value
         'reset Total_Stock_Volume
         Total_Stock_Volume = 0
      
      Else
         'Count Total_Stock_Volume
         Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
         
      End If
    Next i
      
      For i = 2 To lastrow
      
      'apply color
      If ws.Cells(i, 10).Value <= 0 Then
         ws.Cells(i, 10).Interior.ColorIndex = 3
      Else
         ws.Cells(i, 10).Interior.ColorIndex = 4
      End If
      
         'count Greatest % Increase
      If ws.Cells(i, 11).Value > Greatest_Increase Then
         Greatest_Increase = ws.Cells(i, 11).Value
         'find Greatest_Increase_Ticker value
         Greatest_Increase_Ticker = ws.Cells(i, 9).Value
      End If
         
         'count Greatest % Decrease
      If ws.Cells(i, 11).Value < Greatest_Decrease Then
         Greatest_Decrease = ws.Cells(i, 11).Value
         'find Greatest_Decrease_Ticker value
         Greatest_Decrease_Ticker = ws.Cells(i, 9).Value
      End If
      
         'count Greatest_Total_Volume
      If ws.Cells(i, 12).Value > Greatest_Total_Volume Then
         Greatest_Total_Volume = ws.Cells(i, 12).Value
         'find Greatest_Total_Volume_Ticker
         Greatest_Total_Volume_Ticker = ws.Cells(i, 9).Value
      End If
         
    Next i
    
         'print Greatest % Increase
         ws.Range("Q2") = Greatest_Increase
         'format cell
         ws.Range("Q2").NumberFormat = "0.00%"
         'print Greatest_Increase_Ticker
         ws.Range("P2") = Greatest_Increase_Ticker
         'print Greatest % Decrease
         ws.Range("Q3") = Greatest_Decrease
         'format cell
         ws.Range("Q3").NumberFormat = "0.00%"
         'print Greatest_Decrease_Ticker
         ws.Range("P3") = Greatest_Decrease_Ticker
         'print Greatest_Total_Volume
         ws.Range("Q4") = Greatest_Total_Volume
         'print Greatest_Total_Volume_Ticker
         ws.Range("P4") = Greatest_Total_Volume_Ticker
         
    With ws.Cells.EntireColumn.AutoFit
         
    End With
    
Next ws



End Sub
