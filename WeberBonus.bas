Attribute VB_Name = "Module11"
    ' ----------------------------------------------------------
    ' *Bonus* Find the greatest change for: % increase, % decrease, and volume.
    ' ----------------------------------------------------------
    Sub GreatestChange():

    For Each ws In Worksheets
        ' Create the summary table
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        ' Create the variables
        Dim Ticker1 As String
        Dim Ticker2 As String
        Dim Ticker3 As String
        Dim Greatest_P_Increase As Double
        Dim Greatest_P_Decrease As Double
        Dim Greatest_Volume As Double
        Dim lastrow1 As String
        Dim lastrow2 As String
    
        lastrow1 = ws.Cells(Rows.Count, 11).End(xlUp).Row
        lastrow2 = ws.Cells(Rows.Count, 12).End(xlUp).Row
    
        ' -----------------------------------------------------------
        ' Loop through all worksheet rows and find Greatest % Increase
        ' -----------------------------------------------------------
 
        Greatest_P_Increase = 0
        For i = 2 To lastrow1
    
            If ws.Cells(i, 11).Value > Greatest_P_Increase Then
        
                Greatest_P_Increase = ws.Cells(i, 11).Value
                Ticker1 = ws.Cells(i, 9).Value
                
            Else
            
                Greatest_P_Increase = Greatest_P_Increase
                Ticker1 = Ticker1
                
            End If
        
                ' Print Greatest Percent Increase and Ticker to summary table
                ws.Cells(2, 16).Value = Ticker1
                ws.Cells(2, 17).Value = Greatest_P_Increase
                    
                    ' Format cell to percent
                    ws.Range("Q2").Style = "Percent"
                
        Next i
        
       
            ' -----------------------------------------------------------
            ' Loop through all worksheet rows and find Greatest % Decrease
            ' -----------------------------------------------------------

            Greatest_P_Decrease = 0
            For i = 2 To lastrow1
        
                If ws.Cells(i, 11).Value < Greatest_P_Decrease Then
        
                    Greatest_P_Decrease = ws.Cells(i, 11).Value
                    Ticker2 = ws.Cells(i, 9).Value
                    
                Else
            
                    Greatest_P_Increase = Greatest_P_Decrease
                    Ticker2 = Ticker2
                
                End If
            
                    ' Print Greatest Percent Decrease and Ticker to summary table
                    ws.Cells(3, 16).Value = Ticker2
                    ws.Cells(3, 17).Value = Greatest_P_Decrease
                    
                        ' Format cell to percent
                        ws.Range("Q3").Style = "Percent"
            Next i
        
                ' -----------------------------------------------------------
                ' Loop through all worksheet rows and find Greatest Total Volume
                ' -----------------------------------------------------------

                Greatest_Volume = 0
                For i = 2 To lastrow2
    
                    If ws.Cells(i + 1, 12).Value > Greatest_Volume Then
        
                        Greatest_Volume = ws.Cells(i + 1, 12).Value
                        Ticker3 = ws.Cells(i + 1, 9).Value
        
                    Else
            
                        Greatest_Volume = Greatest_Volume
                        Ticker3 = Ticker3
                    
                    End If
                    
                Next i
                
                    ' Print Greatest Percent Increase and Ticker to summary table
                    ws.Cells(4, 16).Value = Ticker3
                    ws.Cells(4, 17).Value = Greatest_Volume
                    
                        ' Format Greatest Volume Cell
                        ws.Cells(4, 17).Style = "Comma"
                
    Next ws
End Sub

