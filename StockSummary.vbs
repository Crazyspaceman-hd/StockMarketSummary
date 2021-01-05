Attribute VB_Name = "Module11"
Sub wallStreet()

    'define variables: vol-totalling stock volume, begin-opening stock price
    Dim vol As Variant
    Dim begin As Double
    
    'insert column headers
    [I1:L1] = [{"Ticker","Yearly Change","Percent Change","Total Stock Volume"}]
    
    'define rowcount variable and find the correct number
    Dim rowcount
    rowcount = Cells(Rows.Count, 1).End(xlUp).Row
    'Set up j variable to start at 2 to iterate output through correct rows
    j = 2
    'Set vol to 0 for consistency
    vol = 0
        
    For i = 2 To rowcount
    
    'If ticker symbol does not match the one above
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            'Set opening price
            begin = Cells(i, 3)
            'Set initial volume
            vol = vol + Cells(i, 7)
            'If ticker symbol does not match the one below, and begin is greater than 0.
        ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            'Print Ticker in Ticker Column
            Cells(j, 9).Value = Cells(i, 1)
            'Add last volume
            vol = vol + Cells(i, 7)
            'Print Total Volume
            Cells(j, 12).Value = vol
            'Reset Volume to 0
            vol = 0
            'Print Yearly Change
            Cells(j, 10).Value = Cells(i, 6) - begin
            'Print Percentage Change w/ divide by 0 protection
                If begin > 0 Then
                Cells(j, 11).Value = Cells(j, 10) / begin
                Else
                Cells(j, 11).Value = 0
                End If
            'Iterate j varible
            j = j + 1
        'If ticker symbol matches
        Else
        'add volume
        vol = vol + Cells(i, 7)
        'check if begin is 0 and if so check to see if it has changed
        If begin = 0 Then begin = Cells(i, 3)
        
             
    
    End If
    
    Next i
    
    'AutoFit Columns to correct width
    Columns("I:L").AutoFit
    'Format Percent Change Column to correct number format
    Columns("K:K").NumberFormat = "0.00%"
    'Add Conditional Formatting, found process on automateexcel.com
    Cells(1, 10).Select
            Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
            Formula1:="=""Yearly Change"""
    Columns("J:J").Select
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
        Selection.FormatConditions(2).Interior.Color = RGB(255, 0, 0)
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
        Selection.FormatConditions(3).Interior.Color = RGB(0, 255, 0)
       
    
    
    
End Sub

