Sub XLHOMEWORK():
'Before applying this code, make sure to select Cell "A1" or the equivilant for your table

    'A requirement for this task is to have this code work on multiple worksheets
    'This can be coded with the following
        For Each ws In Worksheets
    'Adding "ws." to the start of code now will ensure that the line of code words
    'across multiple spreadsheets

    'The data being transcribed needs to be organised into a new table with titles
        ws.Range("I1").Value = "Ticker Symbol"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
    'This can also be done with "Cells(#, #).Value = "______"""
    
    'State the variables required for this particular activity
        Dim Ticker As String
        Dim TotalVolume As Double
            TotalVolume = 0
        Dim SummaryRow As Integer
            SummaryRow = 2
        Dim CloseP As Double
            
        Dim OpenP As Double
        Dim YearlyChange As Double
        Dim PercentageChange As Double

    'Create a variable to find the last row for each worksheet in this large set of data
        Last = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'A conditional loop will be required to find the unique ticker symbols
        For i = 2 To Last

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Symbol = ws.Cells(i, 1).Value
            ws.Range("I" & SummaryRow).Value = Symbol
    'The following will calculate the total volume for each Ticker Symbol
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            ws.Range("L" & SummaryRow).Value = TotalVolume
    
    'Set a formula to calculate the yearly change
        CloseP = ws.Cells(i, 6).Value
        YearlyChange = (CloseP - OpenP)

    'Print the yearly change values into their assigned column
        ws.Range("J" & SummaryRow).Value = YearlyChange

    'To remove the div/0 error, the following can be applied
        If OpenP = 0 Then
            PercentageChange = 0
        Else
            PercentageChange = YearlyChange / OpenP
        End If
    
    'Display the yearly change as percentage change in assigned column
        ws.Range("K" & SummaryRow).Value = PercentageChange
    'The following will display the numbers in a percentage format
        ws.Range("K" & SummaryRow).NumberFormat = "0.00%"

    'Reset the row counter
        SummaryRow = SummaryRow + 1
    'Reset total volume to 0
        TotalVolume = 0
    'Reset Open Price
            OpenP = ws.Cells(i + 1, 3)
        Else
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        End If
        
    'Conditional formatting to make analysis of yearly change easier
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Font.Color = RGB(51, 100, 55)
            ws.Cells(i, 10).Interior.Color = RGB(209, 245, 212)
        ElseIf Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Font.Color = RGB(152, 26, 26)
            ws.Cells(i, 10).Interior.Color = RGB(252, 172, 172)
        End If
        Next i
    Next ws
End Sub
