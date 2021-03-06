Attribute VB_Name = "VBA_Wallstreet_loop"
Sub Wallstreet():

'Added on error to resolve Error 6 on line 56
On Error Resume Next

    'Define the variables
    'Count # of Worksheets
    Dim ws_count As Integer
    
    'standard for loops
    Dim i As Integer
    Dim j As Long
    Dim k As Integer
    'find ticker
    Dim t As Integer
    
    
    'Values for Variants
    Dim Total_Volume As Variant
    Dim open_price As Double
    Dim close_price As Double
    


    'Set intiial values
    k = 2
    ws_count = ActiveWorkbook.Worksheets.Count
 
    
    
        'Begin Loop for i
        For i = 1 To ws_count
     
        
        'Create labels
        ActiveWorkbook.Worksheets(i).Cells(1, 9).Value = "Ticker"
        ActiveWorkbook.Worksheets(i).Cells(1, 12).Value = "Total_Volume"
        ActiveWorkbook.Worksheets(i).Cells(1, 10).Value = "Yearly_Change"
        ActiveWorkbook.Worksheets(i).Cells(1, 11).Value = "Percent_Change"
        
        'Set column to Currency
        Columns("K2:K").NumberFormat = "$#,##0.00"
        
        ActiveWorkbook.Worksheets(i).Cells(1, 15).Value = "Ticker"
        ActiveWorkbook.Worksheets(i).Cells(1, 16).Value = "Value"
        ActiveWorkbook.Worksheets(i).Cells(2, 14).Value = "Greatest % Increase"
        ActiveWorkbook.Worksheets(i).Cells(3, 14).Value = "Greatest % Decrease"
        ActiveWorkbook.Worksheets(i).Cells(4, 14).Value = "Greatest Total Volume"
        
        Total_Volume = 0
     
        'Auto fit Columns
        ActiveWorkbook.Worksheets(i).Columns("A:P").AutoFit
        
        'Beign Loop for j
            For j = 2 To ActiveWorkbook.Worksheets(i).Cells.SpecialCells(xlCellTypeLastCell).Row
                 If ActiveWorkbook.Worksheets(i).Cells(j, 1).Value <> ActiveWorkbook.Worksheets(i).Cells(j + 1, 1).Value Then
                 
                    'Look for pricing
                    close_price = ActiveWorkbook.Worksheets(i).Cells(j, 6).Value
                    ActiveWorkbook.Worksheets(i).Cells(k, 10).Value = close_price - open_price
                    ActiveWorkbook.Worksheets(i).Cells(k, 11).Value = Cells(k, 10).Value / open_price
                    
                    'Make cell style %
                        ActiveWorkbook.Worksheets(i).Cells(k, 10).Style = "Currency"
                        ActiveWorkbook.Worksheets(i).Cells(k, 11).Style = "Percent"
                        
                        'Color Cells based on postitive ot negetive yearly change
                        If ActiveWorkbook.Worksheets(i).Cells(k, 10).Value > 0 Then
                            ActiveWorkbook.Worksheets(i).Cells(k, 10).Interior.ColorIndex = 4
                        Else
                            ActiveWorkbook.Worksheets(i).Cells(k, 10).Interior.ColorIndex = 3
                        End If
                    
                    
                    open_price = 0
                    close_price = 0
                    
                    'Sums
                    ActiveWorkbook.Worksheets(i).Cells(k, 9).Value = ActiveWorkbook.Worksheets(i).Cells(j, 1).Value
                    Total_Volume = Total_Volume + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
                    ActiveWorkbook.Worksheets(i).Cells(k, 12).Value = Total_Volume
                    k = k + 1
                    Total_Volume = 0
                    
                ElseIf ActiveWorkbook.Worksheets(i).Cells(j - 1, 1).Value <> ActiveWorkbook.Worksheets(i).Cells(j, 1).Value Then
                
                    open_price = ActiveWorkbook.Worksheets(i).Cells(j, 3).Value
                    Total_Volume = Total_Volume + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
                    
                Else
                    Total_Volume = Total_Volume + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
                    
                End If
            Next j
        
            ' compare value to findgreatest values in set (Increase, Decrease, Percentage)
            ActiveWorkbook.Worksheets(i).Cells(2, 16).Value = WorksheetFunction.Max(Worksheets(i).Range("K2:K" & ActiveWorkbook.Worksheets(i).Range("K1").CurrentRegion.Rows.Count))
            ActiveWorkbook.Worksheets(i).Cells(3, 16).Value = WorksheetFunction.Min(Worksheets(i).Range("K2:K" & ActiveWorkbook.Worksheets(i).Range("K1").CurrentRegion.Rows.Count))
            ActiveWorkbook.Worksheets(i).Cells(4, 16).Value = WorksheetFunction.Max(Worksheets(i).Range("L2:L" & ActiveWorkbook.Worksheets(i).Range("K1").CurrentRegion.Rows.Count))
            
            'Find Ticker for Min and Max Values
            For t = 2 To ActiveWorkbook.Worksheets(i).Range("K1").CurrentRegion.Rows.Count
                If ActiveWorkbook.Worksheets(i).Cells(t, 11).Value = ActiveWorkbook.Worksheets(i).Cells(2, 16).Value Then
                    ActiveWorkbook.Worksheets(i).Cells(2, 15).Value = ActiveWorkbook.Worksheets(i).Cells(t, 9).Value
                    ' style cell
                    ActiveWorkbook.Worksheets(i).Cells(2, 16).Style = "Percent"
                ElseIf ActiveWorkbook.Worksheets(i).Cells(t, 11).Value = ActiveWorkbook.Worksheets(i).Cells(3, 16).Value Then
                    ActiveWorkbook.Worksheets(i).Cells(3, 15).Value = ActiveWorkbook.Worksheets(i).Cells(t, 9).Value
                    ' style cell
                    ActiveWorkbook.Worksheets(i).Cells(3, 16).Style = "Percent"
                ElseIf ActiveWorkbook.Worksheets(i).Cells(t, 12).Value = ActiveWorkbook.Worksheets(i).Cells(4, 16).Value Then
                    ActiveWorkbook.Worksheets(i).Cells(4, 15).Value = ActiveWorkbook.Worksheets(i).Cells(t, 9).Value
                    ' style cell
                    ActiveWorkbook.Worksheets(i).Cells(4, 16).NumberFormat = "0"
                End If
            Next t
        'Reset the k value for next sheet
            k = 2
            
  Next i
        

End Sub
