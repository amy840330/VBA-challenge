Attribute VB_Name = "Module1"
Sub Wt()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In Worksheets
            ws.Activate
            alpha
            cc
        Next
        Application.ScreenUpdating = True

End Sub
            
    
Sub alpha()


Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent change"
Range("L1").Value = " Total Stock Volumn"

Dim ticker_name As String
Dim summary_table As Integer
Dim lastrow As Long
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As String
    
summary_table = 2
        
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
            
open_price = Cells(2, 3).Value

                 For i = 2 To lastrow
                    
                    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                        ticker_name = Cells(i, 1).Value
                        total_volume = total_volume + Cells(i, 7).Value
                
                        close_price = Cells(i, 6).Value
                
                        yearly_change = Round(close_price - open_price, 2)
                
                        Range("I" & summary_table).Value = ticker_name
                        Range("J" & summary_table).Value = yearly_change
                
                
                        If yearly_change <> 0 And open_price <> 0 Then
                        
                        percent_change = Format(yearly_change / open_price, "Percent")
                        
                        Else
                        
                        percent_change = 0
                        
                        End If
                        
                        
                        Range("K" & summary_table).Value = percent_change
                    
                        
                        Range("L" & summary_table).Value = total_volume
                        
                        open_price = Cells(i + 1, 3).Value
                        
                        summary_table = summary_table + 1
                
                        total_volume = 0
                        
                        
                
                    Else
                
                        total_volume = total_volume + Cells(i, 7).Value
                
                   End If
                
                Next i
    
End Sub

Sub cc()

Dim lastrow2 As Long
lastrow2 = Cells(Rows.Count, 10).End(xlUp).Row

For j = 2 To lastrow2

    If Cells(j, 10).Value > 0 Then
    
    Cells(j, 10).Interior.ColorIndex = 4
    
    Else
    
    Cells(j, 10).Interior.ColorIndex = 3
    
    End If
    
Next j

End Sub






