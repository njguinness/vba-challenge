Attribute VB_Name = "Module1"
Sub Try_Number_XYZ()
'i didn't have fun

'define variables

Dim Ticker As String

'Define a variable for Open_Price

Dim Open_Price As Double

'Define avariable for Close_Price

Dim Close_Price As Double

'Define a variable for Quarterly Change

Dim Quarterly_Change As Double

'Define a variable for Total Volume

Dim Total_Volume As Double

'Define a variable for percent change

Dim Percent_Change As Double

'Define a variable to for Start_Row

Dim Start_Row As Integer

'Define variable of the worksheet

Dim ws As Worksheet



For Each ws In Worksheets
'Create Column Names

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

 'loop start
    Start_Row = 2
    previous_i = 1
    Total_Volume = 0

    'Go to the last row of ticker

    EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        'get data for each Ticker

        For i = 2 To EndRow

            'start new ticker symbol

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


            Ticker = ws.Cells(i, 1).Value


            previous_i = previous_i + 1


            Open_Price = ws.Cells(previous_i, 3).Value
            Close_Price = ws.Cells(i, 6).Value

            'total volume

            For j = previous_i To i

                Total_Volume = Total_Volume + ws.Cells(j, 7).Value

            Next j

                  If Open_Price = 0 Then

                Percent_Change = Close_Price

            Else
                Quarterly_Change = Close_Price - Open_Price

                Percent_Change = Quarterly_Change / Open_Price

            End If
      

  'collapsed data

            ws.Cells(Start_Row, 9).Value = Ticker
            ws.Cells(Start_Row, 10).Value = Quarterly_Change
            ws.Cells(Start_Row, 11).Value = Percent_Change

            'Use percentage format

            ws.Cells(Start_Row, 11).NumberFormat = "0.00%"
            ws.Cells(Start_Row, 12).Value = Total_Volume

            Start_Row = Start_Row + 1

           'reset
            Total_Volume = 0
            Quarterly_Change = 0
            Percent_Change = 0

            
            previous_i = i

        End If

   'yay done this part and onto the next
'see the pretty numbers that show gain or loss

    Next i

'Biggest Loser or Winner or whatever helps my taxes at the end of the year
'please don't tell the irs

'we need to find the last cell that has numbers in precent change
'if i have a clue they might be colored but we'll see


    kEndRow = ws.Cells(Rows.Count, "K").End(xlUp).Row

 'make numbers zero to compare

    Increase = 0
    Decrease = 0
    Greatest = 0

        For k = 3 To kEndRow

            'check versus previous value
            last_k = k - 1

            'where the % change is now
            current_k = ws.Cells(k, 11).Value

            'the percent change prior to the % change we were just looking at
            prevous_k = ws.Cells(last_k, 11).Value

            volume = ws.Cells(k, 12).Value

           
            prevous_vol = ws.Cells(last_k, 12).Value

 
'if current number greater than previous numner
            If Increase > current_k And Increase > prevous_k Then

                Increase = Increase

                increase_name = ws.Cells(k, 9).Value

            ElseIf current_k > Increase And current_k > prevous_k Then

                Increase = current_k

                'name for increase percentage
                increase_name = ws.Cells(k, 9).Value

            ElseIf prevous_k > Increase And prevous_k > current_k Then

                Increase = prevous_k

                ' name for increase percentage
                increase_name = ws.Cells(last_k, 9).Value

            End If


            If Decrease < current_k And Decrease < prevous_k Then

                'Define decrease as decrease

                Decrease = Decrease

               

            ElseIf current_k < Increase And current_k < prevous_k Then

                Decrease = current_k


                decrease_name = ws.Cells(k, 9).Value

            ElseIf prevous_k < Increase And prevous_k < current_k Then

                Decrease = prevous_k

                decrease_name = ws.Cells(last_k, 9).Value

            End If

   
   

            If Greatest > volume And Greatest > prevous_vol Then

                Greatest = Greatest

               

            ElseIf volume > Greatest And volume > prevous_vol Then

                Greatest = volume

                'name for greatest volume
                greatest_name = ws.Cells(k, 9).Value

            ElseIf prevous_vol > Greatest And prevous_vol > volume Then

                Greatest = prevous_vol

                greatest_name = ws.Cells(last_k, 9).Value

            End If

        Next k
  '--------------------------------------------------
    ' Assign names for greatest increase,greatest decrease, and  greatest volume

    ws.Range("N1").Value = "Column Name"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker Name"
    ws.Range("P1").Value = "Value"

    'Get for greatest increase, greatest increase, and  greatest volume Ticker name
    ws.Range("O2").Value = increase_name
    ws.Range("O3").Value = decrease_name
    ws.Range("O4").Value = greatest_name
    ws.Range("P2").Value = Increase
    ws.Range("P3").Value = Decrease
    ws.Range("P4").Value = Greatest

    'Greatest increase and decrease in percentage format

    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"



'The end row for column J

    jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row


        For j = 2 To jEndRow

            'if greater than or less than zero
            If ws.Cells(j, 10) > 0 Then

                ws.Cells(j, 10).Interior.ColorIndex = 4

            Else

                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If

        Next j

'onto the next ws
Next ws

End Sub
