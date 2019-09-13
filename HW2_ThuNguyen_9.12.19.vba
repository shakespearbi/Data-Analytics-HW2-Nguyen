Attribute VB_Name = "Module1"
Sub Stock_calc()
'Hard
    ' 'ticker var
    Dim ticker As String
    
    'following ticker
    Dim n_ticker As String
    
    'stock volume
    Dim vol As Double
    
    Dim lRow As Long
    
    'starting position of data
    Dim pos As Double
    pos = 2
        
    'opening price
    Dim opening As Double
    
    'closing price
    Dim closing As Double
    
    'yearly change
    Dim yearly_chg As Double
    
    'percent change
    Dim percent_chg As Double


    'total volume
    Dim tot_vol As Double
    tot_vol = 0

        
        'Retrieve number of rows
        lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Total Stock Volume"
        Cells(1, 11).Value = "Yearly Change"
        Cells(1, 12).Value = "Percent Change"
        
        opening = Cells(pos, 3).Value
        
        
        For i = 2 To lRow
        
            ticker = Cells(i, 1).Value
            n_ticker = Cells(i + 1, 1).Value
            vol = Cells(i, 7).Value
            
                'check if current ticker value is different from next ticker
                If ticker <> n_ticker Then
                   
                    Range("I" & pos).Value = ticker
                    
                    'calculate total volume
                    tot_vol = tot_vol + vol
                    Range("J" & pos).Value = tot_vol
                    
                    'reset total volume
                    tot_vol = 0
                
                    closing = Cells(i, 6).Value
                   
                    'calculate yearly change
                    yearly_chg = closing - opening
                    
                    Range("K" & pos).Value = yearly_chg
                    Range("K" & pos).NumberFormat = "0.000000000"
                    
                    If opening <> 0# Then
                        'calculate percent change
                        percent_chg = (closing - opening) / opening
                    Else
                        percent_chg = 0#
                    End If
                    
                    Range("L" & pos).Value = percent_chg
                    Range("L" & pos).NumberFormat = "0.00%"
                    
                    'define next opening price
                    opening = Cells(i + 1, 3).Value
            
                    pos = pos + 1
                
                'check if ticker and next ticker are equal
                Else
                    tot_vol = tot_vol + vol
                 End If
            
        Next i
        
        'Conditional formating for yearly change
        For i = 2 To lRow
            If Cells(i, 11).Value < 0 Then
                Cells(i, 11).Interior.ColorIndex = 3
            Else
                Cells(i, 11).Interior.ColorIndex = 4
            End If
        Next i
        
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Total Volume"
        
        'retrieve length of yearly change column
        Dim year_len As Long
        year_len = Cells(Rows.Count, 12).End(xlUp).Row
        
        'min
        Dim percent_min As Double
        percen_min = 0#
        
        'max
        Dim percent_max As Double
        percent_max = 0#
        
        Cells(1, 16).Value = "Ticket"
        Cells(1, 17).Value = "Value"
        
        
        For i = 2 To year_len
            'find max
            If percent_max < Cells(i, 12).Value Then
                Cells(2, 16).Value = Cells(i, 9).Value
                percent_max = Cells(i, 12).Value
                Cells(2, 17).Value = percent_max
                Cells(2, 17).NumberFormat = "0.00%"
            'find min
            ElseIf percent_min > Cells(i, 12).Value Then
                Cells(3, 16).Value = Cells(i, 9).Value
                percent_min = Cells(i, 12).Value
                Cells(3, 17).Value = percent_min
                Cells(3, 17).NumberFormat = "0.00%"
            End If
        Next i
        
       
        Dim vol_len As Long
        'retrieve length of total stock volume column
        vol_len = Cells(Rows.Count, 10).End(xlUp).Row
        
        Dim totvol_max As Double
        totvol_max = 0#
        
        'Finding maximum stock volume
        For i = 2 To vol_len
            If totvol_max < Cells(i, 10).Value Then
                Cells(4, 16).Value = Cells(i, 9).Value
                totvol_max = Cells(i, 10).Value
                Cells(4, 17).Value = totvol_max
            End If
        Next i
    
End Sub

'Challenge
Sub Stock_Challenge()

    For Each ws In Worksheets
      
        'activates worksheet
        ws.Activate
        
        Call Stock_calc
        
    Next ws
        
End Sub

