Sub VBAChallenge()

    ' Declare Current as a worksheet object variable.
    Dim ws As Worksheet

    ' Loop through all of the worksheets in the active workbook.
    For Each ws In Worksheets

    ' Set an initial variable for holding the ticker
    Dim Ticker As String

    'Create new column headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greates Total Volume"

    ' Set an initial variable for holding the volume total
    Dim Vol_Total As Double
    Vol_Total = 0

    ' Keep track of the location for summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through all tickers
    For i = 2 To LastRow

    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker
      Ticker = ws.Cells(i, 1).Value

      ' Add to the Volume Total
      Vol_Total = Vol_Total + ws.Cells(i, 7).Value

      ' Print the Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the Vol to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Vol_Total

      ' Add one to the summary table row for next ticker
      Summary_Table_Row = Summary_Table_Row + 1

      ' Reset the Vol Total
      Vol_Total = 0

    ' If the cell immediately following a row is the same Ticker...
    Else

      ' Add to the Ticker Total
      Vol_Total = Vol_Total + ws.Cells(i, 7).Value

    End If

    Next i

        ' Find start and end rows with unique Ticker
    Dim TickStartRow As Long
    Dim TickEndRow As Long
    Summary_Table_Row = 2
    LastRow1 = ws.Cells(Rows.Count, 9).End(xlUp).Row

        ' Loop through all tickers
    For i = 2 To LastRow1
    
      'Find start and end rows
      TickStartRow = ws.Range("A:A").Find(what:=ws.Cells(i, 9), after:=ws.Cells(1, 1), LookAt:=xlWhole).Row
      TickEndRow = ws.Range("A:A").Find(what:=ws.Cells(i, 9), after:=ws.Cells(1, 1), SearchDirection:=xlPrevious, LookAt:=xlWhole).Row

      ' Print the Year change in the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = ws.Range("F" & TickEndRow).Value - ws.Range("C" & TickStartRow).Value
      
          ' Set color formatting to the cells
      If ws.Range("J" & Summary_Table_Row).Value > 0 Then
      ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      
      ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
      ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      
      End If
      
      ' Print the Percent change in the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = (((ws.Range("F" & TickEndRow).Value - ws.Range("C" & TickStartRow).Value)) / ws.Range("C" & TickStartRow).Value)
      
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      
      'Add one to the summary table row for next ticker
      Summary_Table_Row = Summary_Table_Row + 1

    Next i
    
    
    'Bonus
    
    Dim P_cell As Range
    Dim P_range As Range

    Set P_range = ws.Range("K2:K" & Summary_Table_Row)
        
    lowestnumber = ws.Application.WorksheetFunction.Min(P_range)
    highestnumber = ws.Application.WorksheetFunction.Max(P_range)
        
        For Each P_cell In P_range
        
         If P_cell.Value = lowestnumber Then
        
         ws.Cells(3, 17).Value = lowestnumber
         ws.Cells(3, 17).NumberFormat = "0.00%"
        
         ElseIf P_cell.Value = highestnumber Then
        
         ws.Cells(2, 17).Value = highestnumber
         ws.Cells(2, 17).NumberFormat = "0.00%"
        
         End If
        Next P_cell
        
        For K = 2 To LastRow1

         If ws.Cells(K, 11).Value = highestnumber Then

         ws.Cells(2, 16).Value = ws.Cells(K, 9).Value

         ElseIf ws.Cells(K, 11).Value = lowestnumber Then

         ws.Cells(3, 16).Value = ws.Cells(K, 9).Value

         End If

        Next K
        
    Dim T_cell As Range
    Dim T_range As Range

    Set T_range = ws.Range("L2:L" & Summary_Table_Row)
        
    highestnumberT = ws.Application.WorksheetFunction.Max(T_range)
        
        For Each T_cell In T_range
        
         If T_cell.Value = highestnumberT Then
        
         ws.Cells(4, 17).Value = highestnumberT
        
         End If

        Next T_cell
        
        For L = 2 To LastRow1

         If ws.Cells(L, 12).Value = highestnumberT Then

         ws.Cells(4, 16).Value = ws.Cells(L, 9).Value

         End If

        Next L

    Next ws

End Sub