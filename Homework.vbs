Sub Multi_Yr_StockData()

Dim a As Integer
Dim ws_num As Integer
Dim starting_ws As Worksheet

ws_num = ThisWorkbook.Worksheets.Count

For a = 1 To ws_num

ThisWorkbook.Worksheets(a).Activate

    Dim ticker As String
    Dim vol As Double
Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim Summary_Table_Row As Integer


 Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"


Summary_Table_Row = 2
    

For i = 2 To ActiveSheet.UsedRange.Rows.Count

     If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           
         ticker = Cells(i, 1).Value
           vol = Cells(i, 7).Value
        
         year_open = Cells(i, 3).Value
            year_close = Cells(i, 6).Value

         yearly_change = year_close - year_open
          percent_change = 1 - (year_close / year_open)
        
          Cells(Summary_Table_Row, 9).Value = ticker
           Cells(Summary_Table_Row, 10).Value = yearly_change
            Cells(Summary_Table_Row, 11).Value = percent_change
            Cells(Summary_Table_Row, 12).Value = vol
         Summary_Table_Row = Summary_Table_Row + 1

         vol = 0

     End If

    Next i

    Next a


End Sub

