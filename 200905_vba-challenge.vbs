Sub main()
'--------------------------------------------------------------
' Assumptions made
' Dates are sorted chronologically for the purpose of this exercise
' All data have the same number of columns however, there is an If
' statement incase there are more columns in a certain sheet
'--------------------------------------------------------------
For Each ws In Worksheets

    ' Define groups of variables for easier management
    Dim last_row, last_column, s_row As Long
    Dim total_vol As Double
    Dim t_col, yc_col, pc_col, tsv_col As Integer
    Dim o_price, c_price, yc, pc  As Double
    Dim headers() As Variant
    
    last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
    last_column = ws.Cells(Columns.Count).End(xlToLeft).Column
    'MsgBox ("Last column is: " + last_column)
    
    '--------------------------------------------------------------
    ' Define summary table header
    headers = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    ws.Range("I1:L1").Value = headers
    
    ' Define column number of table
    t_col = 9
    yc_col = 10
    pc_col = 11
    tsv_col = 12
    
    ' Initialise variables
    total_vol = 0
    s_row = 2
    
    ' Set initial open price
    o_price = ws.Cells(2, 3).Value
       
'--------------------------------------------------------------
    For i = 2 To last_row
    
        ' If ticker changes
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
            
            ' Find the last close price
            c_price = ws.Cells(i, 6).Value
            ' Find yearly change
            yc = c_price - o_price
            
            ' Check open price was not 0 as to prevent overflow error
            If o_price > 0 Then
                pc = yc / o_price
            Else
                pc = 0
            End If
            
            total_vol = total_vol + ws.Cells(i, 7).Value
            
            ' Set values in summary table
            ws.Cells(s_row, t_col).Value = ws.Cells(i, 1).Value
            ws.Cells(s_row, yc_col).Value = yc
            ws.Cells(s_row, pc_col).Value = pc
            ws.Cells(s_row, tsv_col).Value = total_vol
            
            ' Format cells to read easier
            ws.Columns("K").NumberFormat = "0.00%"
            ws.Columns("L").NumberFormat = "#,##0"
            
            ' Go down one row in the summary table to enter new ticker
            s_row = s_row + 1
            ' Reset total_vol for next count
            total_vol = 0
            ' Find open price for the next ticker
            o_price = ws.Cells(i + 1, 3).Value
        
        '-------------------
        ' If it's the same ticker
        Else
            total_vol = total_vol + ws.Cells(i, 7).Value
        End If

    Next i
    
'--------------------------------------------------------------
    Dim sum_last_row, inc, dec, tv, y_inc, y_dec As Double
    Dim inc_tic, dec_tic, tv_tic As String

    ' Last row of raw ticker set
    last_row = ws.Cells(Rows.Count, "J").End(xlUp).Row
    
    ' Initialise variables
    inc = 0
    y_inc = 0
    dec = 0
    y_dec = 0
    tv = 0

    ' Test if current current percentage change is either higher or lower than
    ' the previous defined highest value or lowest value
    For j = 2 To last_row
        If ws.Cells(j, 11).Value > inc Then
            inc = ws.Cells(j, 11).Value
            inc_tic = ws.Cells(j, 9).Value
        ElseIf ws.Cells(j, 11).Value < dec Then
            dec = ws.Cells(j, 11).Value
            dec_tic = ws.Cells(j, 9).Value
        Else
        End If
        
        ' Finding highest stock volume
        If ws.Cells(j, 12).Value > tv Then
            tv = ws.Cells(j, 12).Value
            tv_tic = ws.Cells(j, 9).Value
        End If
    Next j
    
    '--------------------------------------------------------------
    ' Formatting negative values as red and positive as green
    For j = 2 To last_row
        If ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
            ws.Cells(j, 10).Font.ColorIndex = 2
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 10
            ws.Cells(j, 10).Font.ColorIndex = 2
        End If
    Next j
    
'--------------------------------------------------------------
    ' Print the overall summary table
    Dim heading(), item() As Variant
    
    heading = Array("Ticker", "Value")
    item = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
        
    ws.Range("P1:Q1") = heading
    ws.Range("O2:O4") = Application.Transpose(item)
    ws.Cells(2, 16).Value = inc_tic
    ws.Cells(2, 17).Value = inc
    ws.Cells(3, 16).Value = dec_tic
    ws.Cells(3, 17).Value = dec
    ws.Cells(4, 16).Value = tv_tic
    ws.Cells(4, 17).Value = tv

    'Formating of table
    ws.Range("Q2,Q3").NumberFormat = "0.00%"
    ws.Range("O1:Q1").Interior.ColorIndex = 1
    ws.Range("O1:Q1").Font.ColorIndex = 2
    ws.Range("Q1").HorizontalAlignment = xlRight
    ws.Range("O1:Q4").BorderAround _
        ColorIndex:=1, Weight:=xlMedium
    ws.Range("I:Q").Columns.AutoFit

Next ws

End Sub

' Clear values for testing code again
Sub reset()

For Each Sheet In Worksheets
    Sheet.Range("I:R").EntireColumn.Delete
Next Sheet

End Sub

