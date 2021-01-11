Sub stockinfo():
    'loop through the stocks for one year in all sheets
    For Each ws In Worksheets
       
        'determine last row
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'determine last column
        Dim lastcolumn As Long
        lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

         'Create value to increase row number when displaying data
                Dim x As Integer
                x = 1
                
                 'add new columns
                ws.Range("I1").Value = "Ticker Symbol"
                ws.Range("J1").Value = "Annual Change"
                ws.Range("K1").Value = "Annual Change %"
                ws.Range("L1").Value = "Total Volume"
           
                'Calculation variables for volume & open/close values
                Dim openval As Double
                Dim closeval As Double
                Dim volume As LongLong
                    volume = 0
                    openval = ws.Cells(2, 3).Value
                    closeval = openval
                Dim annualchange As Double
                Dim changepercent As Double

                'Bonus values for high/low percent & largest volume
                Dim highval As Double
                Dim lowval As Double
                Dim maxvol As LongLong
                                
                'Convert values in date column to dates
                Dim currentdate As String
                Dim converteddate As Date
                        
            
                'Set open & close date variables & initial dates to compare
                Dim opendate As Date
                Dim closedate As Date
                opendate = DateSerial(Left(ws.Cells(2, 2).Value, 4), Mid(ws.Cells(2, 2).Value, 5, 2), Right(ws.Cells(2, 2).Value, 2))
                closedate = DateSerial(Left(ws.Cells(2, 2).Value, 4), Mid(ws.Cells(2, 2).Value, 5, 2), Right(ws.Cells(2, 2).Value, 2))
                
            'Run through data in sheet
            For i = 2 To lastrow

                'Check if ticker symbol matches next
                If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                    
                    'Add volume for each cell that matches
                    volume = volume + ws.Cells(i, 7).Value
                    
                    'Compare current date to open/close dates & set open/close values
                        currentdate = ws.Cells(i, 2).Value
                        converteddate = DateSerial(Left(currentdate, 4), Mid(currentdate, 5, 2), Right(currentdate, 2))
                             
                             If converteddate > closedate Then
                                    closedate = DateSerial(Left(ws.Cells(i, 2).Value, 4), Mid(ws.Cells(i, 2).Value, 5, 2), Right(ws.Cells(i, 2).Value, 2))
                                    closeval = ws.Cells(i, 6).Value
                                    
                            ElseIf ws.Cells(i, 2).Value < opendate Then
                                    opendate = DateSerial(Left(ws.Cells(i, 2).Value, 4), Mid(ws.Cells(i, 2).Value, 5, 2), Right(ws.Cells(i, 2).Value, 2))
                                    openval = ws.Cells(i, 3).Value
                            End If
                    
                Else
               
                    'Move to next row of data to be filled in
                    x = x + 1
                    
                    'Fill in Ticker
                    ws.Cells(x, 9).Value = ws.Cells(i, 1)

                    'Fill in total Volume for Ticker
                    volume = volume + ws.Cells(i, 7).Value
                    ws.Cells(x, 12).Value = volume
                                      

                        'Compare current date to open/close dates & set open/close values
                        currentdate = ws.Cells(i, 2).Value
                        converteddate = DateSerial(Left(currentdate, 4), Mid(currentdate, 5, 2), Right(currentdate, 2))
                             
                             If converteddate > closedate Then
                                    closedate = DateSerial(Left(ws.Cells(i, 2).Value, 4), Mid(ws.Cells(i, 2).Value, 5, 2), Right(ws.Cells(i, 2).Value, 2))
                                    closeval = ws.Cells(i, 6).Value
                                    
                            ElseIf ws.Cells(i, 2).Value < opendate Then
                                    opendate = DateSerial(Left(ws.Cells(i, 2).Value, 4), Mid(ws.Cells(i, 2).Value, 5, 2), Right(ws.Cells(i, 2).Value, 2))
                                    openval = ws.Cells(i, 3).Value
                            End If
                            
                            'Determine annual change & percent
                            annualchange = (closeval - openval)
                            
                                'check if open value is zero
                                If openval = "0" Then
                                    changepercent = annualchange
                                Else
                                    changepercent = (annualchange / openval)
                                End If
                            
                           
                            'Display change values
                            ws.Cells(x, 10).Value = annualchange
                            ws.Cells(x, 11).Value = changepercent
                            
                            
                            'Determine if change is positive or negative & color code
                            If annualchange > 0 Then
                                ws.Cells(x, 10).Interior.ColorIndex = 4
                                
                            
                            ElseIf annualchange < 0 Then
                                ws.Cells(x, 10).Interior.ColorIndex = 3
                                
                            End If
                                      
                    'Set new Open & close dates/values (if on last row, then set open/close dates to last date in sheet)
                        If i = lastrow Then
                        opendate = DateSerial(Left(ws.Cells(i, 2).Value, 4), Mid(ws.Cells(i, 2).Value, 5, 2), Right(ws.Cells(i, 2).Value, 2))
                        closedate = DateSerial(Left(ws.Cells(i, 2).Value, 4), Mid(ws.Cells(i, 2).Value, 5, 2), Right(ws.Cells(i, 2).Value, 2))
                    
                        Else
                        opendate = DateSerial(Left(ws.Cells((i + 1), 2).Value, 4), Mid(ws.Cells((i + 1), 2).Value, 5, 2), Right(ws.Cells((i + 1), 2).Value, 2))
                        closedate = DateSerial(Left(ws.Cells((i + 1), 2).Value, 4), Mid(ws.Cells((i + 1), 2).Value, 5, 2), Right(ws.Cells((i + 1), 2).Value, 2))
                        End If
                    
                        openval = ws.Cells(i + 1, 3).Value
                        closeval = openval
                   
                    
                End If
            'reset volume
            volume = 0

            Next i
        'set percent values as percent format
        For p = 2 To x
        Cells(p, 11).NumberFormat = "0.00%"
        Next p

    Next

End Sub


