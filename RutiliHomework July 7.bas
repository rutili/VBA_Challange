Attribute VB_Name = "Module1"
Option Explicit

Dim WorksheetName As String
Dim CurrentTicker As String
Dim LastTicker As String
Dim TotalStockVolume  As Variant
Dim NewArea As String
Dim YearlyClose As Currency
Dim YearlyOpen As Currency
Dim YearlyChange As Currency
Dim PercentageChange As Double
Dim xOpen As Double
Dim xClose As Double
Dim xCloseX As Long
Dim PrintRow As Long
Dim xxClose As Variant
Dim R1 As Range
Dim ThereisVolume As Boolean
Dim Ws As Worksheet
Dim LastRow1 As Variant
Dim LastColumn As Long
Dim MaxChangeTicker As String
Dim MinChangeTicker As String
Dim MaxVTicker As String
Dim MaxV As Variant
Dim MaxChange As Double
Dim MinChange As Double
Dim rg As Range
Dim CurrentRow As Variant
Dim SumifResult As LongLong
Dim Cond1 As FormatCondition
Dim Cond2 As FormatCondition
Dim Cond3 As FormatCondition
Dim PercentageChange1 As Double


Sub Vba_challange()


'Iterate Thru Each Sheet

For Each Ws In Worksheets()
    
    LastRow1 = Ws.Cells(Rows.Count, 1).End(xlUp).Row
    LastColumn = Ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    NewArea = ("I1:" + "S" + LTrim(Str(LastRow1)))
    Ws.Range(NewArea).Clear

    LastTicker = "MisMatch"
    MinChange = 0
    MaxChange = 0
    MaxV = 0
    PrintRow = 2
    WorksheetName = Ws.Name
    
    MaxChangeTicker = LastTicker
    MinChangeTicker = LastTicker
    MaxVTicker = LastTicker
    
    MsgBox WorksheetName
            
    
    BuildHeaders
                               '  MinChange = PercentChange
                               '  MaxChange = PercentChange
                               '  MaxV = PercentChange
            
    For CurrentRow = 2 To LastRow1

        ThereisVolume = (Ws.Cells(CurrentRow, 7).Value) <> 0
           
           If ThereisVolume Then
                
                'Read
                  
                    CurrentTicker = Ws.Cells(CurrentRow, 1).Value
                         
                            'Compare Tickers
                            
                      If (CurrentTicker <> LastTicker) Then
                    
                                xOpen = Ws.Cells(CurrentRow, 3).Value
                                SumifResult = Application.WorksheetFunction.SumIf(Ws.Range("A2:A" & LastRow1), CurrentTicker, Range("G1:G" & LastRow1))
                                xClose = Ws.Cells(CurrentRow - 1, 6).Value
                                PercentageChange1 = ((xClose - xOpen) / xOpen)
                                PercentageChange = Round(PercentageChange1, 4)
                            Bonus
                            
                            
                            'Write
                   
                                 Ws.Cells(PrintRow, 9).Value = CurrentTicker
                                 Ws.Cells(PrintRow, 10).Value = xOpen
                                 Ws.Cells(PrintRow, 14).Value = SumifResult
                                 Ws.Cells(PrintRow, 11).Value = xClose
                                 Ws.Cells(PrintRow, 12).Value = xClose - xOpen
                                 Ws.Cells(PrintRow, 13).Value = PercentageChange
                                 PrintRow = PrintRow + 1
                                 
                       End If
                  
                    'Store ticker for comparison
                        LastTicker = CurrentTicker
                             
             End If
    
    Next CurrentRow



Ws.Range("J2:" + "l" + LTrim(Str(LastRow1))).NumberFormat = "$#,##0.00"
Ws.Range("M2:" + "M" + LTrim(Str(LastRow1))).NumberFormat = "#.##%"
Ws.Range("N2:" + "N" + LTrim(Str(LastRow1))).NumberFormat = "#,###"
Ws.Range("r2").NumberFormat = "00.0%"
Ws.Range("r3").NumberFormat = "00.0%"
Ws.Range("r4").NumberFormat = "#,##0"
 
 
ColorCoding

Next Ws


End Sub

Sub BuildHeaders()

            Ws.Cells(1, 9).Value = "Ticker"
            Ws.Cells(1, 10).Value = "Yearly Open"
            Ws.Cells(1, 11).Value = "Yearly Close"
            Ws.Cells(1, 12).Value = "Yearly Change"
            Ws.Cells(1, 13).Value = "Percentage Change"
            Ws.Cells(1, 14).Value = "Total Stock volume"
            
            Ws.Cells(1, 17).Value = "Ticker"
            Ws.Cells(1, 18).Value = "Value"
                    Ws.Range("p2") = ("Greatest % Increase")
                    Ws.Range("p3") = ("Greatest % Decrease")
                    Ws.Range("p4") = ("Greatest Total Volume")
                    
 End Sub

Sub Bonus()

    If PercentageChange > MaxChange Then
       MaxChange = PercentageChange
       MaxChangeTicker = CurrentTicker
    End If
    
    If PercentageChange < MinChange Then
       MinChange = PercentageChange
       MinChangeTicker = CurrentTicker
    End If
    
    If SumifResult > MaxV Then
        MaxV = SumifResult
        MaxVTicker = CurrentTicker
    End If

Ws.Range("q2") = MaxChangeTicker
Ws.Range("q3") = MinChangeTicker
Ws.Range("q4") = MaxVTicker

Ws.Range("r2") = MaxChange
Ws.Range("r3") = MinChange
Ws.Range("r4") = MaxV

End Sub

Sub ColorCoding()


Set rg = Ws.Range("L2", Ws.Range("L2").End(xlDown))

'clear any existing conditional formatting
rg.FormatConditions.Delete

'define the rule for each conditional format
Set Cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, 0)
Set Cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, 0)
Set Cond3 = rg.FormatConditions.Add(xlCellValue, xlEqual, 0)

'define the format applied for each conditional format
With Cond1
.Interior.Color = vbGreen
End With

With Cond2
.Interior.Color = vbRed
End With

End Sub
