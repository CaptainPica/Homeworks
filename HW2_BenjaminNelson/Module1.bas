Attribute VB_Name = "Module1"
Sub Stocks():
    Dim Headers() As Variant
    Headers = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    Dim ticker_start As Long ' Is placed at the start of each stock as volumes are added
    Dim ticker_location As Long ' Keeps track of where we are in the <ticker> column
    Dim Volume As Double ' Total Volume for a stock
    Dim record As Integer ' Increments to a new row when recording information for a new stock
    Dim BigEnergy As Single ' Biggest +% change so far
    Dim BETicker As String 'Location of which ticker currently has the biggest change
    Dim SmallEnergy As Single ' Biggest -% change so far
    Dim SETicker As String 'Location of which ticker currently has the smallest change
    Dim BigVolume As Double ' Biggest total volume so far
    Dim BVTicker As String 'Location of which ticker currently has the biggest volume
    BigEnergy = -3.402823E+38 'Smallest value Single can hold. Everything is bigger or equal.
    SmallEnergy = 3.402823E+38 'Biggest value Single can hold. Everything is lesser or equal.
    BigVolume = 0 'Smallest volume that can exist
    Dim i As Integer 'For loop iterators
    Dim j As Integer
    For i = 1 To Worksheets.Count 'Runs through all the sheets and processes each
        Sheets(i).Activate
        ticker_start = 2 'Resets each variable on each new sheet
        ticker_location = 2
        record = 2
        Volume = 0 'initializes volume with a zero value upon first use
        For j = 0 To 3 'Sets the moderate HW header values for recorded data for each new sheet
            Cells(1, 9 + j).Value = Headers(j)
        Next j
        Range("K:K").NumberFormat = "0.00%" 'Sets the percent change column to the right format
        Do While IsEmpty(Cells(ticker_location, 1).Value) = False 'Checks to see if all ticker values have been read
            Do While Cells(ticker_start, 1).Value = Cells(ticker_location, 1).Value 'Runs while ticker is the same
                Volume = Volume + Cells(ticker_location, 7).Value 'Adds up all the volume for a ticker
                ticker_location = ticker_location + 1
            Loop
            Cells(record, 9).Value = Cells(ticker_start, 1) 'This section stores the requested values for a ticker
            Cells(record, 10).Value = Cells(ticker_location - 1, 6) - Cells(ticker_start, 3)
            If Cells(record, 10).Value > 0 Then 'Changes cell color for yearly change
                Cells(record, 10).Interior.Color = RGB(0, 255, 0)
            Else
                Cells(record, 10).Interior.Color = RGB(255, 0, 0)
            End If
            If Cells(ticker_start, 3) = 0 Then 'Checks if stock opened at zero, so no /0 errors
                Cells(record, 11).Value = 0
            Else
                Cells(record, 11).Value = Cells(record, 10).Value / Cells(ticker_start, 3)
                If BigEnergy < Cells(record, 11).Value Then 'Records bigger or smaller % changes than already found
                    BigEnergy = Cells(record, 11).Value
                    BETicker = Cells(ticker_start, 1)
                ElseIf SmallEnergy > Cells(record, 11).Value Then
                    SmallEnergy = Cells(record, 11).Value
                    SETicker = Cells(ticker_start, 1)
                End If
            End If
            Cells(record, 12).Value = Volume
            If Volume > BigVolume Then 'Finds biggest total volume of stock by storing bigger volumes
                BigVolume = Volume
                BVTicker = Cells(ticker_start, 1)
            End If
            record = record + 1 'Sets everything up to process next stock on sheet
            Volume = 0
            ticker_start = ticker_location
        Loop
    Next i
    Sheets(1).Activate 'Records hard HW values
    Range("P2:P3").NumberFormat = "0.00%"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 15).Value = BETicker
    Cells(2, 16).Value = BigEnergy
    Cells(3, 15).Value = SETicker
    Cells(3, 16).Value = SmallEnergy
    Cells(4, 15).Value = BVTicker
    Cells(4, 16).Value = BigVolume
    For i = 1 To Worksheets.Count 'Autofits columns of all sheets so everything looks nice
        Sheets(i).Activate
        Columns("A:P").AutoFit
    Next i
    Sheets(1).Activate 'Sets the right sheet to be active so you can see results most easily
End Sub
