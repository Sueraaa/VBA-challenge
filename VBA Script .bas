Attribute VB_Name = "Module2"
Sub Final_Code()

'Declare and set worksheet
Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In Worksheets

'Create a variable to hold the ticker value
Dim Ticker As String
Ticker = " "
Dim TickerRow As Integer
TickerRow = 1

               'Set new variables for open,close prices and yearly change
                Dim open_price As Double
                open_price = 0
                Dim close_price As Double
                close_price = 0
                Dim yearly_change As Double
                yearly_change = 0
                Dim percent_change As Double
                percent_change = 0
                Dim totalVolume As Double
                totalVolume = 0

               'Declare variables for Variables for Greatest % Increase, Decrease
                Dim greatestIncrease As Double
                greatestIncrease = 0
                Dim greatestDecrease As Double
                greatestDecrease = 0
                Dim greatestTotalVolume As Double
                greatestTotalVolume = 0
                Dim greatestIncreaseTicker As String
                greatestIncreaseTicker = " "
                Dim greatestDecreaseTicker As String
                greatestDecreaseTicker = " "
                Dim greatestVolumeTicker As String
                greatestVolumeTicker = " "

'Define rowCount of Colunm A
Dim rowCount As Long
rowCount = Cells(ws.Rows.Count, "A").End(xlUp).Row

'Creation of the loop
For i = 2 To rowCount

                'Ticker symbol output
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Add 1 to the ticker row
                TickerRow = TickerRow + 1
                
                'Extract the ticker symbol from the current row
                Ticker = ws.Cells(i, 1).Value
                ws.Cells(TickerRow, "I").Value = Ticker
                
                'Calculate yearly change
                close_price = ws.Cells(i, 6).Value
                yearly_change = close_price - open_price
                ws.Cells(TickerRow, "J").Value = yearly_change
            
                'Calculate percentage change
                If open_price <> 0 Then
                percent_change = (yearly_change / open_price) * 100
                End If
                ws.Cells(TickerRow, "K").Value = percent_change
                
                'Calculate total stock volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                ws.Cells(TickerRow, "L").Value = totalVolume
                
                'Calculate Greatest % Increase
                If percent_change > greatestIncrease Then
                    greatestIncrease = percent_change
                    greatestIncreaseTicker = Ticker
                End If
                
                'Calculate Greatest %Decrease
                If percent_change < greatestDecrease Then
                    greatestDecrease = percent_change
                    greatestDecreaseTicker = Ticker
                End If
                
                'Calulate GreatestTotal Volume
                If totalVolume > greatestTotalVolume Then
                    greatestTotalVolume = totalVolume
                    greatestVolumeTicker = Ticker
                End If
                             
                'Reset open price and total volume for the next tickersymbol
                open_price = 0
                totalVolume = 0
                
Else
                'Update open price for next iteration
                If open_price = 0 Then
                open_price = ws.Cells(i, 3).Value
                End If
                
                'Storing the final Totalvolume in column 7
                totalVolume = totalVolume + ws.Cells(i, 7).Value

'End of the loop and calling the next iteration
End If
Next i
 
        'Write the results for Greatest % Increase in column Q & P
        ws.Range("Q2").Value = greatestIncrease
        ws.Range("P2").Value = greatestIncreaseTicker
        
        'Write the results for Greatest %Decrease in column Q & P
        ws.Range("Q3").Value = greatestDecrease
        ws.Range("P3").Value = greatestDecreaseTicker
        
        'Write the results for GreatestTotal Volume in column Q & P
        ws.Range("Q4").Value = greatestTotalVolume
        ws.Range("P4").Value = greatestVolumeTicker

Next ws

End Sub









