Sub stockmarket():

' define all variables

    Dim ws As Worksheet
    Dim ticker As String
    Dim vol As Double
    Dim yearopen As Double
    Dim yearclose As Double
    Dim yearchange As Double
    Dim percentchange As Double
    
    Dim newtablerow As Double
       
       newtablerow = 1
       yearclose = 0
       vol = 0
       
    
'set column headers for each sheet

            For Each ws In ThisWorkbook.Worksheets
            
       
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
    
'set up lastrow

            lastrow = 1
            While ws.Cells(lastrow, 1) <> ""
            lastrow = lastrow + 1
            Wend

' set up loop
            
            yearopen = ws.Cells(2, 3).Value
            
            For x = 2 To lastrow
            
'this bit is trying to see if the ticker has changed
            
            vol = vol + ws.Cells(x, 7).Value
            
            If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
                yearclose = ws.Cells(x, 6).Value
                newtablerow = newtablerow + 1
                ticker = ws.Cells(x, 1).Value
                
'vol = vol + ws.Cells(x, 7).Value
                
                yearchange = yearclose - yearopen
                
                If yearclose <> 0 Then
                    percentchange = (yearchange / yearclose)
                
                End If
                
                ws.Cells(newtablerow, 9).Value = ticker
                ws.Cells(newtablerow, 10).Value = yearchange
                ws.Cells(newtablerow, 11).Value = percentchange
                ws.Cells(newtablerow, 12).Value = vol
                
'get the next yearopen because we know it is changing
                
                yearopen = ws.Cells(x + 1, 3).Value
                
                vol = 0
                
            End If
            
        Next
        
         ws.Columns("K").NumberFormat = "0.00%"
         ws.Columns("J").NumberFormat = "0.00"
         
     Next
     
'conditional formatting
   
    For y = 2 To lastrow
    
        Set r = ActiveSheet.Cells(y, 10)
        
        
            If r < 0 Then
                r.Interior.Color = vbRed
            ElseIf r > 0 Then
               r.Interior.Color = vbGreen
        
            End If
        
        Next
           
End Sub



