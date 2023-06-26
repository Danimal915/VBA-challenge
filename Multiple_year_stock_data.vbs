Attribute VB_Name = "Module1"
Sub stocks_for_year():
    Dim PRESENT_TKR As String
    Dim NEXT_TKR As String
    Dim nexti As Double
    Dim previousi As Double
    Dim DAILY_VOLUME As Double
    Dim PREVIOUS_VOLUME As Double
    Dim YTD_VOLUME As Double
    Dim final_total_volume As Double
    Dim current_closing_price As Double
    Dim final_closing_price As Double
    Dim opening_price As Double
    Dim ytd_change As Double
    Dim percent_change As Double
    Dim final_ytd_change As Double
    Dim final_percent_change As Double
    Dim count As Integer
    Dim Ticker As String

   
    opening_price = Cells(2, 3).Value
    count = 0
    Ticker = Cells(2, 1).Value
        
    For i = 2 To 753001
    
   
    PRESENT_TKR = Cells(i, 1).Value
    NEXT_TKR = Cells(i + 1, 1).Value
    
    nexti = i + 1
    previousi = i - 1
  
    
    'Calculate running volume for year to date
    'Check that current row ticker matches next row ticker
    If PRESENT_TKR = NEXT_TKR Then
        DAILY_VOLUME = Cells(i, 7).Value
        YTD_VOLUME = YTD_VOLUME + DAILY_VOLUME
        
        
        current_closing_price = Cells(i, 6).Value
        final_closing_price = Cells((i + 1), 6).Value

        
        ytd_change = current_closing_price - opening_price
        percent_change = ytd_change / opening_price
        
        final_ytd_change = final_closing_price - opening_price
        final_percent_change = final_ytd_change / opening_price
        
        final_total_volume = YTD_VOLUME

        
    Else
        YTD_VOLUME = 0
        'NEXT_TTL = 0
        'PREVIOUS_TTL = 0
        opening_price = Cells(i + 1, 3).Value
        count = count + 1
        Ticker = Cells(i, 12).Value

    End If

    
    'Displays each new Ticker in each successive Row for Column 10
    Cells((count + 2), 10) = Cells(i, 1).Value '''''Unique Tickers
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Cells((count + 2), 11) = opening_price
    Cells((count + 2), 12) = final_closing_price
    Cells((count + 2), 13) = final_ytd_change
    Cells((count + 2), 14) = FormatPercent(final_percent_change)
    Cells((count + 2), 15) = final_total_volume
    
    Next i
    
    Cells(2, 18) = "Greatest % Increase"
    Cells(3, 18) = "Greatest % Decrease"
    Cells(4, 18) = "Greatest Total Volume"
    
    For j = 2 To 3002
    
    'Get ticker for max change
    If Cells(j, 14).Value = Cells(2, 20).Value Then
        Cells(2, 19).Value = Cells(j, 10).Value
    End If
    
    'Get ticker for min change
    If Cells(j, 14).Value = Cells(3, 20).Value Then
        Cells(3, 19).Value = Cells(j, 10)
    End If
    
    'Get ticker for max volume
    If Cells(j, 15).Value = Cells(4, 20).Value Then
        Cells(4, 19).Value = Cells(j, 10).Value
    End If
    
    Next j
        
    
End Sub
