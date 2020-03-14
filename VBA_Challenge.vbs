Attribute VB_Name = "Module2"
Sub CreateSummaryTable()

Dim LastRow As Long

' Set an initial variable for holding the Ticker Symbol
Dim Ticker As String

' Set an initial variable for holding the total per Ticker symbol
Dim TickerTotal As Double
TickerTotal = 0

' Keep track of the location for each Ticker symbol in the summary table
Dim Summary_Table_Row As Long
Summary_Table_Row = 2

'Keep track of start and end Ticker prices
Dim TickerOpen As Double
Dim TickerClose As Double
TickerOpen = 0
TickerClose = 0

'Set initial variables to keep track of price changes
Dim YearChange As Double
Dim PctChange As Double
YearChange = 0
PctChange = 0

'Add a sheet for Ticker Open and Close prices
Sheets.Add.Name = "TickerData"
Sheets("TickerData").Move Before:=Sheets(1)
Dim Ticker_Data_Row As Long
Ticker_Data_Row = 2
'Add header row to TickerData sheet
Sheets("TickerData").Range("A1").Value = "Ticker"
Sheets("TickerData").Range("B1").Value = "Ticker Open"
Sheets("TickerData").Range("C1").Value = "Ticker Close"


    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Reset Summary Table Row for each WS
        Summary_Table_Row = 2
        
        ' ------------------------------------------------------------------
        ' FIND UNIQUE TICKER SYMBOLS AND WRITE SUMMARY TABLE VALUES FOR EACH
        ' ------------------------------------------------------------------
        For i = 2 To LastRow
            
            'For first Ticker symbol, set opening price in TickerData sheet
            If i = 2 Then
                Ticker = ws.Cells(i, 1)
                Sheets("TickerData").Range("A" & Ticker_Data_Row).Value = Ticker
                TickerOpen = ws.Cells(i, 3).Value
                Sheets("TickerData").Range("B" & Ticker_Data_Row).Value = TickerOpen
            'Else
            End If
            
            ' Check if next row is the same as the current the Ticker symbol, if it is not...
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                ' Set the Ticker Symbol
                Ticker = ws.Cells(i, 1).Value

                ' Add to the Ticker Total
                TickerTotal = TickerTotal + ws.Cells(i, 7).Value

                ' Print the Ticker Symbol in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker

                ' Print the Ticker Total Volume to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = TickerTotal
                
                'For the last row, set the close value in TickerData sheet
                TickerClose = ws.Cells(i, 6).Value
                Sheets("TickerData").Range("C" & Ticker_Data_Row).Value = TickerClose
                
                'For next Ticker symbol, set opening price in TickerData sheet
                Ticker = ws.Cells(i + 1, 1)
                Sheets("TickerData").Range("A" & Ticker_Data_Row + 1).Value = Ticker
                TickerOpen = ws.Cells(i + 1, 3).Value
                Sheets("TickerData").Range("B" & Ticker_Data_Row + 1).Value = TickerOpen
                
                ' Print the Ticker Yearly Change to the Summary Table
                TickerOpen = Sheets("TickerData").Range("B" & Ticker_Data_Row).Value
                TickerClose = Sheets("TickerData").Range("C" & Ticker_Data_Row).Value
                YearChange = TickerClose - TickerOpen
                ws.Range("J" & Summary_Table_Row).Value = YearChange
                
                ' Print the Ticker Percent Change to the Summary Table
                If TickerOpen <> 0 And YearChange <> 0 Then
                    PctChange = (YearChange / TickerOpen)
                Else
                    PctChange = 0
                End If
                'MsgBox ("PctChange " & Ticker & " " & YearChange & "/" & TickerOpen & "* 100 =" & PctChange)
                ws.Range("K" & Summary_Table_Row).Value = PctChange

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
                ' Reset the variables
                TickerTotal = 0
                YearChange = 0
                PctChange = 0


            ' If the cell immediately following current row is the same ticker symbol...
            Else

                ' Add to the Ticker Total
                TickerTotal = TickerTotal + ws.Cells(i, 7).Value

            End If
            
            If Sheets("TickerData").Range("A" & Ticker_Data_Row + 1) <> "" Then
                Ticker_Data_Row = Ticker_Data_Row + 1
            'Else
            End If
            
        Next i

        'IF statement to ignore TickerData sheet when creating summary tables
        If ws.Name <> "TickerData" Then
            ' --------------------------------------------
            ' WRITE SUMMARY TABLE HEADERS AND FORMAT
            ' --------------------------------------------
            ws.Range("I1") = "Ticker Symbol"
            ws.Range("J1") = "Yearly Change"
            ws.Range("K1") = "Percent Change"
            ws.Range("L1") = "Total Volume"
            ws.Range("P1") = "Ticker"
            ws.Range("Q1") = "Value"
            ws.Range("O2") = "Greatest % Increase"
            ws.Range("O3") = "Greatest % Decrease"
            ws.Range("O4") = "Greatest Total Volume"
            ws.Range("I1:Q1").Font.Bold = True
            ws.Range("I1:Q1").Font.Underline = True
            ws.Range("J:J").Style = "Currency"
            ws.Range("K:K").Style = "Percent"
            ws.Range("Q2:Q3").Style = "Percent"
            ws.Range("I:Q").HorizontalAlignment = xlCenter
            Columns("I:Q").Select
            Selection.EntireColumn.AutoFit
        
            ' --------------------------------------------
            ' SUMMARY TABLE COMPLETE
            ' --------------------------------------------
            
            ' --------------------------------------------
            ' OVERALL STOCK SUMMARY DATA POPULATION
            ' --------------------------------------------
            
            ws.Range("Q2").Value = Application.WorksheetFunction.Max(ws.Range("K" & 2 & ":" & "K" & LastRow))
            ws.Range("Q3").Value = Application.WorksheetFunction.Min(ws.Range("K" & 2 & ":" & "K" & LastRow))
            ws.Range("Q4").Value = Application.WorksheetFunction.Max(ws.Range("L" & 2 & ":" & "L" & LastRow))
        
            ' --------------------------------------------
            ' TEMP DATA TO TEST CONDITIONAL FORMATTING
            ' --------------------------------------------
            
            'ws.Range("J2").Value = 1
            'ws.Range("K2").Value = 1
            'ws.Range("J3").Value = -1
            'ws.Range("K3").Value = -1
        
            ' -----------------------------------------------
            ' CONDITIONAL FORMATTING FOR YEARLY AND % CHANGES
            ' -----------------------------------------------
            
            For i = 2 To LastRow
                If ws.Cells(i, 10).Value > 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 4
                    ws.Cells(i, 11).Interior.ColorIndex = 4
                ElseIf ws.Cells(i, 10).Value < 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                    ws.Cells(i, 11).Interior.ColorIndex = 3
                Else
                End If
            Next i
    
        End If
    
    Next ws
    
    Sheets("TickerData").Range("A1").Select
    'MsgBox ("")
    

End Sub
