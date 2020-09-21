Sub CalculateSummary()
Dim Ticker As String
Dim Total_Vol As Double
Dim Summary_Table_Row As Integer
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim WS As Worksheet

For Each WS In ActiveWorkbook.Worksheets
WS.Activate

Range("I:Q").Value = ""
Range("I:Q").Interior.ColorIndex = 0
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Summary_Table_Row = 2

Total_Vol = 0

For i = 2 To 797711

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Total_Vol = Total_Vol + Cells(i, 7).Value
        Cells(Summary_Table_Row, 9).Value = Ticker
        Cells(Summary_Table_Row, 12).Value = Total_Vol
        Open_Price = Cells(2, 3).Value
        Close_Price = Cells(i, 6).Value
        Yearly_Change = (Close_Price - Open_Price)
        Cells(Summary_Table_Row, 10).Value = Yearly_Change
        
        If (Open_Price = 0) Then
        
            Percent_Change = 0
            
        Else
        
            Percent_Change = Yearly_Change / Open_Price
            
        End If
        
        Cells(Summary_Table_Row, 11).Value = Percent_Change
        Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        Total_Vol = 0
        
        Open_Price = Cells(i + 1, 3)
        
    Else
    
        Total_Vol = Total_Vol + Cells(i, 7).Value
        
    End If
    
Next i

For i = 2 To 3169

    If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 10
    Else
        Cells(i, 10).Interior.ColorIndex = 3
    End If
    
Next i

Next WS

End Sub
Sub SetTitle()

    Range("I:Q").Value = ""
    Range("I:Q").Interior.ColorIndex = 0
' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"


End Sub

