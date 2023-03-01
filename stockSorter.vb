Sub stockSorter()


'Set initial variable for holding ticker name
Dim Ticker_Name As String

'Set initial variable for holding total ticker volume
Dim Ticker_Total As Double
Ticker_Total = 0


'Keep track of location for each Ticker name in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

 Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly_change"
    Cells(1, 12).Value = "Total Stock Vol"
    Cells(1, 11).Value = "Yearly_percentage"
    
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Yearly_change"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    
    
 
    
    

Dim Yearly_percent As Double



'Loop through all tickers
For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row

         If year_open = 0 Then

          year_open = Cells(i, 3).Value
      End If

    'Check if still same ticker name, if it isn't..
    If Cells(i - 1, 1) = Cells(i, 1) And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        year_close = Cells(i, 6).Value
        yearly_change = year_close - year_open
        Yearly_percent = (year_close - year_open) / year_open * 100
    
        'Set ticker name
        Ticker_Name = Cells(i, 1).Value
    
        'Add to ticker volume total
        Ticker_Total = Ticker_Total + Cells(i, 7).Value

        Range("I" & Summary_Table_Row).Value = Ticker_Name
        
        Range("K" & Summary_Table_Row).Value = Yearly_percent
        
        Range("j" & Summary_Table_Row).Value = yearly_change

        Range("L" & Summary_Table_Row).Value = Ticker_Total
    
        Summary_Table_Row = Summary_Table_Row + 1
    
        Ticker_Total = 0
        
        year_open = 0


  
    
    Else
   

    
        Ticker_Total = Ticker_Total + Cells(i, 7).Value

        End If

    
Next i



Dim f As Long, r1 As Range

   For f = 2 To Cells(Rows.Count, 10).End(xlUp).Row
      Set r1 = Range("J" & f)
      
      If r1.Value > 0 Then r1.Interior.Color = vbGreen
      If r1.Value < 0 Then r1.Interior.Color = vbRed
      If r1.Value = 0 Then r1.Interior.Color = vbYellow
   Next f

    Cells(2, 17) = Application.WorksheetFunction.Max(Range("k:k"))
    Cells(3, 17) = Application.WorksheetFunction.Min(Range("k:k"))
    Cells(4, 17) = Application.WorksheetFunction.Max(Range("l:l"))
    
      Dim e As Long, j As Long
    Dim lastRowQ As Long, lastRowK As Long
    Dim searchValue As Double
    Dim ticker As String
    
    lastRowQ = Cells(Rows.Count, "Q").End(xlUp).Row
    lastRowK = Cells(Rows.Count, "K").End(xlUp).Row
    
    For e = 2 To lastRowQ
        searchValue = Cells(e, "Q").Value
        ticker = ""
        
        ' Check if current row is Q4 and value is in scientific notation
        If e = 4 And Left(CStr(searchValue), 1) = "1" And Len(CStr(searchValue)) > 15 Then
            searchValue = CDbl(searchValue)
        End If
        
        For j = 2 To lastRowK
            If searchValue = Cells(j, "K").Value Or searchValue = Cells(j, "L").Value Then
                ticker = Cells(j, "I").Value
                Exit For
            End If
        Next j
        
        If ticker = "" Then
            Cells(e, "P").Value = "Ticker Not Found"
        Else
            Cells(e, "P").Value = ticker
        End If
    Next e


End Sub

