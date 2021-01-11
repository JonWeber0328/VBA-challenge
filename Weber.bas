Attribute VB_Name = "Module1"
Sub VBAChallenge():
    ' ----------------------------------------------------------
    ' Loop through all worksheets
    For Each ws In Worksheets
    
        ' -----------------------------------------------------------
        ' Set up the variables
        ' -----------------------------------------------------------
        ' Set initial variable to hold Ticker
        Dim Ticker As String
              
        'Set initial variable to hold the Volume
        Dim Volume As String
        Volume = 0
        
        ' Set variable for Open Date
        Dim Open_Date As Double
        Open_Date = 20160101
        
        ' Set variable for Open Price and Close Price
        Dim Open_Price As Double
        Dim Close_Price As Double
        
        ' Set variable for Yearly Change and Percent Change
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        
        ' Keep track of each Ticker location in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        ' -------------------------------------------------------------
        ' Set up summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' -------------------------------------------------------------
        ' Look through each row
        Dim lastrow As Double
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
        
            ' Check date to get Open Price
            If Open_Date = ws.Cells(i, 2).Value Then
                
                ' Capture Open price
                Open_Price = ws.Cells(i, 3).Value
                
            
            ' Check if the Ticker is still the same
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Set the Ticker name
                Ticker = ws.Cells(i, 1).Value
                
                ' Print the Ticker name in the summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                
                ' Add the Volume for each Ticker
                Volume = Volume + ws.Cells(i, 7).Value
                
                ' Print the Total Stock Volume to summary table
                ws.Range("L" & Summary_Table_Row).Value = Volume
                
                ' Capture Close Price
                Close_Price = ws.Cells(i, 6).Value
                
                ' Calculate Yearly Change
                Yearly_Change = Close_Price - Open_Price
                
                ' Print the Yearly_Change to summary table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                    ' Set conditional formatting that will highlight possitive and negative changes
                    
                ' Calculate Percentage Change
                    ' Make sure not dividing by 0
                    If Open_Price = 0 Then
                
                        Percent_Change = NA
                
                    Else
                        Percent_Change = (Yearly_Change / Open_Price)
                    End If
                
                
                ' Print Percentage Change to summary table
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                
                ' Format Percent Change cells
                ws.Range("K" & Summary_Table_Row).Style = "Percent"
                
                ' Reset: Volume, Open Price, Close Price
                Volume = 0
                Open_Price = 0
                Close_Price = 0
                Yearly_Change = 0
                Percentage_Change = 0
                
                
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




















