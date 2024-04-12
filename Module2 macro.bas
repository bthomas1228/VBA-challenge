Attribute VB_Name = "Module1"
Sub ticker()
For Each ws In Worksheets

' Create columns and add headers for Ticker, Yearly Change, Percent Change and Total Stock Volume
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Loop through worksheet to output ticker symbol, yearly change, percent change and total volume for each stock
    
      
    'Initiate variable to index row to deposit ticker, yearly change, percent change and volume info
    Dim rowcount As Integer
    rowcount = 2
    
    'Initiate variable to count rows as the loop progresses (rows of same stock)
    Dim counter As Integer
        counter = 0
    
    'Initiate variables to define the info to capture
    Dim ticker As String
    Dim yearly As Double
    Dim percent As Double
    Dim totalvolume As Double
    Dim volume As Double
        volume = 0
 
        
    'Create for loop to retrieve the information and place it into the appropriate cells in the correct format

    For Row = 2 To LastRow
        If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
            ticker = ws.Cells(Row, 1).Value
            yearly = (ws.Cells(Row, 6).Value - ws.Range("C" & Row - counter).Value)
            percent = (yearly / ws.Range("C" & Row - counter).Value)
            totalvolume = volume + ws.Cells(Row, 7).Value
                                
            ws.Range("I" & rowcount).Value = ticker
            ws.Range("J" & rowcount).Value = yearly
                If yearly >= 0 Then
                    ws.Range("J" & rowcount).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & rowcount).Interior.ColorIndex = 3
                End If
            
            ws.Range("K" & rowcount).Value = percent
            ws.Range("K" & rowcount).NumberFormat = "0.00%"
            ws.Range("L" & rowcount).Value = totalvolume
        
            rowcount = rowcount + 1
            counter = 0
            volume = 0
            
        Else
            counter = counter + 1
            volume = volume + ws.Cells(Row, 7).Value
            
        End If
    
    Next Row
    

'compare rows to find highest or lowest value, use tickertracker to mark row number

    For Row = 2 To LastRow
        If ws.Cells(Row, 11).Value > Previous Then
            Previous = ws.Cells(Row, 11).Value
            tickertracker = ws.Cells(Row, 9).Value
        
        End If
    Next Row
    
    ws.Range("Q2").Value = Previous
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("P2").Value = tickertracker
    
    Previous = 0
    tickertracker = 2
 
    For Row = 2 To LastRow
        If ws.Cells(Row, 11).Value < Previous Then
            Previous = ws.Cells(Row, 11).Value
            tickertracker = ws.Cells(Row, 9).Value
        
        End If
    Next Row
    
    ws.Range("Q3").Value = Previous
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("P3").Value = tickertracker
    
    Previous = 0
    tickertracker = 2
    
    For Row = 2 To LastRow
        If ws.Cells(Row, 12).Value > Previous Then
            Previous = ws.Cells(Row, 12).Value
            tickertracker = ws.Cells(Row, 9).Value
        
        End If
    Next Row
    
    ws.Range("Q4").Value = Previous
    ws.Range("P4").Value = tickertracker
    
    Previous = 0
    tickertracker = 2
    
   Next ws
       
End Sub




