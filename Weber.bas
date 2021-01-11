Attribute VB_Name = "Module1"
Sub VBAChallenge():

    ' Loop through all worksheets
    For Each ws In Worksheets
    
        ' Set initial variable to hold Ticker
        Dim Ticker As String
        
        ' Keep track of each Ticker location in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Set initial variable to hold the Volume
        Dim Volume As String
        Volume = 0
        
        ' Set up summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Look through each row
        Dim lastrow As Double
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
        
            ' Check if the Ticker is still the same
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Set the Ticker name
                Ticker = ws.Cells(i, 1).Value
                
                ' Print the Ticker name in the summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                
                ' Add the Volume for each Ticker
                Volume = Volume + ws.Cells(i, 7).Value
                
                ' Print the Total Stock Volume to summary table
                ws.Range("L" & Summary_Table_Row).Value = Volume
                
                ' Reset the Volume
                Volume = 0
                
                ' Add one to Summary_Table_Row
                Summary_Table_Row = Summary_Table_Row + 1
                
            ' If the cell immediately following a row is the same Ticker...
            Else
                
                ' Add to the Ticker Volume
                Volume = Volume + ws.Cells(i, 7).Value
                
            End If
                    
        
        Next i
        
    
    Next ws
    
    
    

End Sub




















