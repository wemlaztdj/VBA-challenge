Attribute VB_Name = "VBAchallenge"
Option Explicit
Private MaxIncrease As Double
Private MIticker As String
Private MaxDecrease As Double
Private MDticker As String
Private MaxTotal As LongLong
Private MTticker As String

Sub StockAnalysis()
'loop through every sheet
Application.ScreenUpdating = False
Dim wks As Worksheet
Dim LastRow As LongLong
Dim I As Integer

For Each wks In ActiveWorkbook.Worksheets
    LastRow = wks.Cells(wks.Rows.Count, "A").End(xlUp).Row
    Call Readdata(wks, LastRow)
Next wks
Application.ScreenUpdating = True
End Sub

Private Sub Readdata(wks As Worksheet, LastRow As LongLong)
'read the data on the sheet
'Dim variables
Dim ticker As String
Dim openingPrice As Double
Dim endingPrice As Double
Dim openingPrice As Double
Dim endingPrice As Double
Dim SumVolume As LongLong
Dim I As LongLong
Dim tickerRow As LongLong
    tickerRow = 2
    

'set header
    wks.Range("I1:Q" & LastRow).ClearContents
    wks.Cells(1, 9) = "Ticker"
    wks.Cells(1, 10) = "Yearly Change"
    wks.Cells(1, 11) = "Percent Change"
    wks.Cells(1, 12) = "Total Stock Volume"

'Loop each row

    For I = 2 To LastRow
        If wks.Cells(I, 1) <> ticker Then
            If ticker <> "" Then
                Call writeSum(wks, ticker, openingPrice, endingPrice, SumVolume, tickerRow)
                tickerRow = tickerRow + 1
            End If
            
            ticker = wks.Cells(I, 1)
            thedate = CLng(wks.Cells(I, 2))
            openingPrice = wks.Cells(I, 3)
            endingPrice = wks.Cells(I, 6)
            SumVolume = wks.Cells(I, 7)
        Else
            If theEndDate < CLng(wks.Cells(I, 2)) Then
                endingPrice = wks.Cells(I, 6)
                theEndDate = CLng(wks.Cells(I, 2))
            ElseIf theOpenDate > CLng(wks.Cells(I, 2)) Then
                openingPrice = wks.Cells(I, 3)
                theOpenDate = CLng(wks.Cells(I, 2))
            End If
            SumVolume = SumVolume + wks.Cells(I, 7)
            
        End If
        
    Next I
    Call writeSum(wks, ticker, openingPrice, endingPrice, SumVolume, tickerRow)
'set format for Percentage
    wks.Range("K2:K" & tickerRow).NumberFormat = "0.00%"
    Call setCF(wks, tickerRow)
    Call printMaxs(wks)

End Sub
Private Sub writeSum(wks As Worksheet, ticker As String, openP As Double, endP As Double, total As LongLong, SRow As LongLong)
'calulate the Percent and Sum
    wks.Cells(SRow, 9) = ticker
    wks.Cells(SRow, 10) = openP - endP
    wks.Cells(SRow, 11) = (openP - endP) / openP
    wks.Cells(SRow, 12) = total
    Call GetMaxs(ticker, (openP - endP) / openP, total)
    
End Sub
Private Sub setCF(wks As Worksheet, SRow As LongLong)
'set conditional formatting red & green
    Dim rng As Range
    Set rng = wks.Range("J2:K" & SRow)
    rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    rng.FormatConditions(1).Interior.Color = 255

    rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
    rng.FormatConditions(2).Interior.Color = 5287936

End Sub
Private Sub GetMaxs(ticker, change, volume)
'Check max values
    If change > MaxIncrease Then
        MaxIncrease = change
        MIticker = ticker
    ElseIf change < MaxDecrease Then
        MaxDecrease = change
        MDticker = ticker
    End If
    
    If volume > MaxTotal Then
        MaxTotal = volume
        MTticker = ticker
    End If
    
End Sub

Private Sub printMaxs(wks As Worksheet)
'add headers to the sheet
    wks.Cells(1, 16) = "Ticker"
    wks.Cells(1, 17) = "Value"
    wks.Cells(2, 15) = "Greatest % Increase"
    wks.Cells(3, 15) = "Greatest % Decrease"
    wks.Cells(4, 15) = "Greatest Total Volume"
'set ticker & value
    wks.Cells(2, 16) = MIticker
    wks.Cells(3, 16) = MDticker
    wks.Cells(4, 16) = MTticker

    wks.Cells(2, 17) = MaxIncrease
    wks.Cells(2, 17).NumberFormat = "0.00%"
    wks.Cells(3, 17) = MaxDecrease
    wks.Cells(3, 17).NumberFormat = "0.00%"
    wks.Cells(4, 17) = MaxTotal

Call ResetPublicVariables

End Sub

Private Sub ResetPublicVariables()
'reset variables for each sheet
    MaxIncrease = 0
    MIticker = ""
    MaxDecrease = 0
    MDticker = ""
    MaxTotal = 0
    MTticker = ""
End Sub
