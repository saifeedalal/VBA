Option Explicit

Sub StockAnalysis()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    Dim ws As Worksheet
    Dim Ticker As String
    Dim Total_Stock_Volume As Double
    Dim Summary_Table_Row As Integer
    Dim LastRow As Double
    Dim lastTickerRow As Double
    Dim rangeOpenVal As Range
    Dim rangeCloseVal As Range
    Dim rangeUniqueTicker As Range
    Dim yearlyChange As Double
    Dim count1 As Long
    Dim count2 As Long
    Dim count3 As Long
    Dim firstTickerRow As Boolean
    Dim maxPercentIncrease As Double
    Dim minPercentIncrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentIncreaseRow As Double
    Dim minPercentIncreaseRow As Double
    Dim maxTotalVolumeRow As Double
        
    For Each ws In Worksheets
    
        ws.Activate

        Total_Stock_Volume = 0
        count1 = 0
        count2 = 0
        count3 = 0
        
        Set rangeOpenVal = ws.Range("R2")
        Set rangeCloseVal = ws.Range("S2")

        ' Keep track of the location for each row/line in the summary table
        Summary_Table_Row = 2
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total VOlume"
        
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        firstTickerRow = True
        
        For count1 = 2 To LastRow

            ' Check if we are still within the same value, if it is not...
            If ws.Cells(count1 + 1, 1).Value <> ws.Cells(count1, 1).Value Then

                ' Set the Ticker Value
                Ticker = ws.Cells(count1, 1).Value

                ' Add to Total_Stock_Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(count1, 7).Value

                ' Print Ticker in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker

                ' Print the Total_Stock_Volume to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                ' Set the yearClose value for the ticket in the temporary range object
                rangeCloseVal(Summary_Table_Row - 1).Value = ws.Cells(count1, 6).Value
                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' set to true so that we fetch yearOpen value for next row
                firstTickerRow = True
                    
                ' Reset Total_Stock_Volume
                Total_Stock_Volume = 0
                
            Else

                ' Add to the Total Stock Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(count1, 7).Value
                
                'Select yearOpen Val for first row of the ticker
                If (firstTickerRow = True And ws.Cells(count1, 3).Value <> 0) Then
                    rangeOpenVal(Summary_Table_Row - 1).Value = ws.Cells(count1, 3).Value
                    firstTickerRow = False
                End If

            End If
                
        Next count1
              
        
        'count of unique tickers
        lastTickerRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
        'Iterating for each unique tracker to calculate yearly change and percent change
        For count2 = 2 To lastTickerRow
        
            yearlyChange = rangeCloseVal(count2 - 1) - rangeOpenVal(count2 - 1)
            ws.Range("J" & count2).Value = yearlyChange
             
            ' Additional safegurad to avoid divde by zero error
            If (rangeOpenVal(count2 - 1) > 0) Then
             
                ws.Range("K" & count2).Value = (yearlyChange / rangeOpenVal(count2 - 1))
                'tempRange(count2-1).Value = (yearlyChange / rangeOpenVal(count2 - 1))
            Else
                
                ws.Range("K" & count2).Value = 0
                Debug.Print "Zero open value for " & ws.Range("I" & count2).Value
                
            End If
            
            ' Color formatting
            If (yearlyChange >= 0) Then
        
                ws.Range("J" & count2).Interior.ColorIndex = 4
        
            Else
        
                ws.Range("J" & count2).Interior.ColorIndex = 3
        
            End If
                    
        Next count2

        'Setting the % format
        ws.Range("K2:K" & lastTickerRow).NumberFormat = "0.00%"
        
        'Finding and setting maxPercentIncrease and minPercentIncrease
        maxPercentIncrease = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastTickerRow))
        minPercentIncrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastTickerRow))
        
        ws.Range("Q2").Value = maxPercentIncrease
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = minPercentIncrease
        ws.Range("Q3").NumberFormat = "0.00%"
        
        'Creating a temporary copy column for vlookup purpose
        ws.Range("K:K").Copy Destination:=ws.Range("H:H")
        
        'Finding the ticker name for maxPercentIncrease and minPercentIncrease
        ws.Range("P2").Value = Application.WorksheetFunction.VLookup(maxPercentIncrease, ws.Range("H2:I" & lastTickerRow), 2, False)
        ws.Range("P3").Value = Application.WorksheetFunction.VLookup(minPercentIncrease, ws.Range("H2:I" & lastTickerRow), 2, False)
        
        'Clear contents of temp column
        ws.Columns(8).EntireColumn.Clear
        
        'Finding and setting maxTotalVolume
        maxTotalVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastTickerRow))
        ws.Range("Q4").Value = maxTotalVolume
        
        'Creating a temporary copy column for vlookup purpose
        ws.Range("L:L").Copy Destination:=ws.Range("H:H")
        
        'Finding the ticker name for maxTotalVolume
        ws.Range("P4").Value = Application.WorksheetFunction.VLookup(maxTotalVolume, ws.Range("H2:I" & lastTickerRow), 2, False)
        
        'Deleting/Clearing all temp columns
        ws.Columns(8).EntireColumn.Clear
        ws.Columns(18).EntireColumn.Delete
        ws.Columns(18).EntireColumn.Delete
        
        ws.Columns.AutoFit
            
    Next ws

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
End Sub




