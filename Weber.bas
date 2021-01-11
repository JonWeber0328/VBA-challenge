Attribute VB_Name = "Module1"
Sub VBA_Challenge():

    ' Loop through all worksheets
    For Eash ws In Worksheets
    
        ' Set initial variable to hold Ticker
        Dim Ticker As String
        Dim Ticker As Double
        
        ' Keep track of each Ticker location in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Set initial variable to hold the Volume
        Dim Volume As String
        Volume = 0
        
        ' Look through each row
        LastRow = ws.Cells(Rows.Count, 1).End(x1Up).Row
        
        For i = 2 To LastRow
        
            ' Check if the Ticker is still the same
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                ' Set the Ticker name
                Ticker = Cells(i, 1).Value
                
                ' Print the Ticker name in the summary table
                Range("I" & Summary_Table_Row).Value = Ticker
                
                ' Add one to Summary_Table_Row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Add the Volume for each Ticker
                Volume = Volume + Cells(i, 7).Value
                
                
        
        
        
        Next i
        
    
    Next ws
    
    
    

End Sub

