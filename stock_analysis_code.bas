Attribute VB_Name = "Module1"
Option Explicit
Sub WorksheetLoop()

    Dim Sht As Worksheet
    Dim LastRow, lstrw As Long
    Dim rng1, rng2, i, j, ticka, tickb, ticktable As Integer
    Dim opn, cls, total, volume, perc, dmin, dmax As Double
    Dim tick As String

    For Each Sht In Worksheets

        Sht.Activate
        LastRow = Sht.Range("A" & Rows.Count).End(xlUp).Row
        lstrw = Sht.Range("I" & Rows.Count).End(xlUp).Row
        
        'Advanced filter to copy unique tickers to column I sourced and adapted from: https://stackoverflow.com/questions/36044556/quicker-way-to-get-all-unique-values-of-a-column-in-vba
        Sht.Range("A1:A" & LastRow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("I1"), Unique:=True
        

        ticka = 2
        tickb = 2
        ticktable = 2
        volume = 0
        For i = 2 To LastRow
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                tick = Cells(i, 1).Value
                volume = volume + Cells(i, 7).Value
                Range("L" & ticktable).Value = volume
                ticktable = ticktable + 1
                volume = 0
            Else
                volume = volume + Cells(i, 7).Value
            End If
            
            
            If Cells(i, 1).Value = Cells(ticka, 9).Value Then
                If Cells(i, 2).Value < Cells(i + 1, 2).Value Then
                    opn = Cells(i, 3).Value
                    ticka = ticka + 1
                End If
            End If
            
            If Cells(i, 1).Value = Cells(tickb, 9).Value Then
                If Cells(i, 2).Value > Cells(i + 1, 2).Value Then
                    cls = Cells(i, 6).Value
                    total = cls - opn
                    perc = (cls / opn) - 1
                    Cells(tickb, 10).Value = total
                    Cells(tickb, 11).Value = perc
                    Cells(tickb, 11).NumberFormat = "0.00%"
                    If total < 0 Then
                        Cells(tickb, 10).Interior.ColorIndex = 3
                    Else
                        Cells(tickb, 10).Interior.ColorIndex = 4
                    End If
                    tickb = tickb + 1
                End If
            End If

        Next i
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Range("Q2:Q3").NumberFormat = "0.00%"
        
        
        'min and max functions sourced and adapted from https://www.excelanytime.com/excel/index.php?option=com_content&view=article&id=105:find-smallest-and-largest-value-in-range-with-vba-excel&catid=79&Itemid=475
        
        rng1 = Range("K2:K" & Rows.Count).Value
        rng2 = Range("L2:L" & Rows.Count).Value
        Cells(2, 17).Value = Application.WorksheetFunction.Max(rng1)
        Cells(3, 17).Value = Application.WorksheetFunction.Min(rng1)
        Cells(4, 17).Value = Application.WorksheetFunction.Max(rng2)
        
        ' sourced and adapted from https://www.automateexcel.com/vba/vlookup-xlookup/
        Cells(2, 16).Value = Application.WorksheetFunction.XLookup(Range("Q2").Value, Range("K2:K" & Rows.Count).Value, Range("I2:I" & Rows.Count).Value)
        Cells(3, 16).Value = Application.WorksheetFunction.XLookup(Range("Q3").Value, Range("K2:K" & Rows.Count).Value, Range("I2:I" & Rows.Count).Value)
        Cells(4, 16).Value = Application.WorksheetFunction.XLookup(Range("Q4").Value, Range("L2:L" & Rows.Count).Value, Range("I2:I" & Rows.Count).Value)
        Columns("I:Q").AutoFit
        
    Next Sht

End Sub
