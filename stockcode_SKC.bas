Attribute VB_Name = "Module1"



Sub stocks():

'loop code through all the worksheets
For Each ws In Worksheets


    'setting up the table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'setting up table for bonus
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"


    'TICKER COLUMN
    'set variables for ticker, endnumA ,and summary table
    Dim Ticker As String
    Dim endnumA As Long
    Dim Table As Long

    'Set endnum to find end of column A
    endnumA = ws.Range("A1").End(xlDown).Row

    'Set the beginning row for summary Table
    Table = 2

        'start loop with beginning of data to endnumA
        For i = 2 To endnumA
    
            'set up if statement looking for change in value between cells in column A
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                'set Ticker to reflect the value just before the change
                Ticker = ws.Cells(i, 1).Value
            
                'put the values for ticker in column M
                ws.Range("I" & Table).Value = Ticker
            
                'add to table so each value gets it own cell
                Table = Table + 1
            End If
        Next i


    'YEARLY CHANGE, PERCENT CHANGE, AND TOTAL STOCK VOLUME
    'Set variables for year open, year close, and year change
    Dim YRChange As Double
    Dim YROpen As Double
    Dim YRClose As Double

    'set Variables for Percent Change aand Stock Volume
    Dim PercentChange As Double
    Dim StockVol As Double

    'set the intial start value for YROpen
    YROpen = ws.Cells(2, 3).Value

    'set stock volume to 0
    StockVol = 0

    'set the value for table back to 2 so it doesnt start after the final value for ticker stuff
    Table = 2

        'start loop from 2 to endnumA again
        For i = 2 To endnumA
    
            'If statement to look for change in column A
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
                'set value for year close
                YRClose = ws.Cells(i, 6).Value
            
                'calculate the year change (minusclose from open)
                YRChange = YRClose - YROpen
            
                'calulate the percent change by dividing the year change by year open
                    If (YROpen = 0 And YRClose = 0) Then
                        PercentChange = 0
                        'PercentChange = "n/a"
                        'I checked through both datasets and there was only one stock in each that had this problem
                    ElseIf (YROpen = 0 And YRClose <> 0) Then
                        PercentChange = 1
                    Else: PercentChange = YRChange / YROpen
                    End If
                    
            
                'calculate stock volume by adding up column G
                StockVol = StockVol + ws.Cells(i, 7).Value
            
                'Do Ranges for YRChange, Percent Change, Stock Volume
                ws.Range("J" & Table).Value = YRChange
                ws.Range("K" & Table).Value = PercentChange
                ws.Range("L" & Table).Value = StockVol
                'ws.Range("M" & Table).Value = YRClose (used for troubleshooting)
                'ws.Range("N" & Table).Value = YROpen (used for troubleshooting)
                
                'format percent change
                ws.Range("K" & Table).NumberFormat = "0.00%"
            
                'make sure each value gets its own cell
                Table = Table + 1
            
                'set stock volume back to 0, so they dont all add together
                StockVol = 0
                
                'set next year open value, using the same change from this itteration of if
                YROpen = ws.Cells(i + 1, 3).Value
            
            Else
                'set other, non-change columns to add up
                StockVol = StockVol + ws.Cells(i, 7).Value
            
            End If
          
        Next i
    
    
    'CONDITIONAL FORMATING FOR YEARLY CHANGE (COLORS)
    'set and define end for column M
    Dim endnumI As Long
    endnumI = ws.Cells(Rows.Count, 10).End(xlUp).Row

        'set loop from 2 to end of column M
        For i = 2 To endnumI
    
            'If greater than 0, the cell is green
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            
            'If less than 0, the cell is red
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            'If any equal zero then it is yellow
            ElseIf ws.Cells(i, 10).Value = 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 6
        
            End If
        Next i
        
    'find the max value in the percent change column
    ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(Range(ws.Cells(2, 11), ws.Cells(endnumI, 11)))
    
    'find the minimum value in percent change
    ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(Range(ws.Cells(2, 11), ws.Cells(endnumI, 11)))
    
    'find greatest total volume
    ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(Range(ws.Cells(2, 12), ws.Cells(endnumI, 12)))
    
    'format the cells with percents into percent format
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    'define variables for greatest percent increase,greatest percent decrease, and Greatest Total Volume
    Dim GrPerIn As Double
    Dim GrPerDec As Double
    Dim GrTotVol As Double
    
    'set variables equal to values from above
    GrPerIn = ws.Cells(2, 17).Value
    GrPerDec = ws.Cells(3, 17).Value
    GrTotVol = ws.Cells(4, 17).Value
    
        'loop through numbers 2 to endnumI
        For i = 2 To endnumI
        
            'find tickers using the values, and put them into table for greatest %increase and decrease
            If GrPerIn = ws.Cells(i, 11).Value Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            ElseIf GrPerDec = ws.Cells(i, 11).Value Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            End If
            
            'find ticker for greatest stock volume, put in table
            If GrTotVol = ws.Cells(i, 12).Value Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            End If
        Next i
        
MsgBox (ws.Name)
    
Next ws
     

End Sub

