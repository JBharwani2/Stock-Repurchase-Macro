Sub VolumeLimitUpdate()
    '
    ' VolumeLimitUpdate Macro
    ' Created by Jeremy Bharwani on 5/27/21
    ' (questions- email jcb926@gmail.com)
    '
    '   This macro is used to update an excel worksheet with daily stock market data to solve for the trading volume
    '   limit. The program is designed to disregards weekends and holidays and then divide each week with a border.
    '   Equations and formatting is also done automatically by the program. If the sheet is already up to date, nothing
    '   will change. Only updates up to the previously market closed day (not current day while market is still open).
    '   The bottom-most line will show the next market-open day's volume limit which is useful when approaching a new
    '   week. The range for the condition limit equation is updated anytime there is a 2+ day gap since the previous
    '   entry (this is when the border is created at the end of the week).
    '
    '   Website Scraped: MarketWatch.com
    '   References Used: "Microsoft HTML Object Library" and "Microsoft Internet Controls"
    '
    '   WARNING: Do not alter cell J2, it is used as a temporary holding place during the macro process.
    '

    'VARIABLES -------------------------------------------------------------------------------------------------------
    Dim http As Object
    Dim html As New HTMLDocument
    Dim four_week_start As Integer
    Dim four_week_end As Integer
    Dim row As Integer
    Dim backlog_num As Integer


    'SET EXCEL LOCATIONS ---------------------------------------------------------------------------------------------
    FindEndOfData row:=row
    FourWeekRange four_week_start:=four_week_start, four_week_end:=four_week_end, row:=row
    
 
'CONNECTS TO WEBPAGE ---------------------------------------------------------------------------------------------
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://www.marketwatch.com/investing/stock/{TICKER}/download-data", False
    http.send
    html.body.innerHTML = http.responseText
    
 
'DATA GRAB --------------------------------------------------------------------------------------------------------
    'Scrapes data from top row of historical quotes page
    'Top row is the previous market close values and does not show live updates from current day's trading
    Set Table = html.getElementsByTagName("tbody")(4)
    Set Data = Table.Children(0)
    backlog_num = -1

    'Sets place-holder with date of the latest update
    Sheets("volume limit").Range("J" & 2).Value = Format(Data.Children(0).Children(0).innerText, "M/d/yyyy")
    'Uses place-holder's date to find how many un-updated days there have been since the last update
    GetBacklog Table:=Table, backlog_num:=backlog_num, row:=row
    'Clears place-holder
    Sheets("volume limit").Range("J" & 2).Value = ""

    'Fills in each column with values or formulas (loop depends on backlog of un-updated days since last update)
    While backlog_num >= 0
        Set Data = Table.Children(backlog_num)
    
        Sheets("volume limit").Range("A" & row).Value = Format(Data.Children(0).Children(0).innerText, "M/d/yyyy")
        SeparateWeeks four_week_start:=four_week_start, four_week_end:=four_week_end, row:=row

        Sheets("volume limit").Range("B" & row).Value = Data.Children(1).innerText
        Sheets("volume limit").Range("C" & row).Value = Data.Children(2).innerText
        Sheets("volume limit").Range("D" & row).Value = Data.Children(3).innerText
        Sheets("volume limit").Range("E" & row).Value = Data.Children(4).innerText
        Sheets("volume limit").Range("F" & row).Value = Data.Children(5).innerText
        Sheets("volume limit").Range("G" & row).Formula = "=ROUND(AVERAGE($F$" & four_week_start & ":$F$" & four_week_end & ")*0.25,-2)"
        Sheets("volume limit").Range("H" & row).Formula = "=IFERROR(VLOOKUP($A" & row & ",Activity!$A$7:$I$1462,3,FALSE),0)"
        Sheets("volume limit").Range("I" & row).Formula = "=IF(H" & row & "<G" & row & "," & Chr(34) & Chr(34) & ",+H" & row & "-G" & row & ")"

        'Formats newly input data to match with previous
        Sheets("volume limit").Range("A" & row & ":I" & row).Font.Name = "Arial"
        Sheets("volume limit").Range("B" & row & ":E" & row).NumberFormat = "0.00"
        Sheets("volume limit").Range("F" & row & ":I" & row).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"
        Sheets("volume limit").Range("I" & row).Font.Color = RGB(255, 0, 0)

        row = row + 1
        backlog_num = backlog_num - 1
    Wend
    
    LimitForecast four_week_start:=four_week_start, four_week_end:=four_week_end, row:=row
    MsgBox "Volume Limit Sheet has been updated to current date"
End Sub



'Finds the next available row in the sheet
Public Sub FindEndOfData(ByRef row As Integer)
    row = 3
    While IsEmpty(Sheets("volume limit").Range("A" & row)) = False
        row = row + 1
    Wend
End Sub



'Finds the start and end of the previous four week period
Public Sub FourWeekRange(ByRef four_week_start As Integer, ByRef four_week_end As Integer, row As Integer)
    four_week_start = row
    four_week_end = row
    Dim i As Integer
    i = 0
    While i < 5
        If Sheets("volume limit").Range("A" & four_week_start).Borders(xlEdgeTop).LineStyle <> xlLineStyleNone Then
            i = i + 1
        End If
        four_week_start = four_week_start - 1
    Wend
    four_week_start = four_week_start + 1
    i = 0
    While i < 1
        If Sheets("volume limit").Range("A" & four_week_end).Borders(xlEdgeTop).LineStyle <> xlLineStyleNone Then
            i = i + 1
        End If
        four_week_end = four_week_end - 1
    Wend
End Sub



'Finds the oldest data that has not been added to the sheet yet
Public Sub GetBacklog(Table, ByRef backlog_num As Integer, row As Integer)
    While Sheets("volume limit").Range("A" & row - 1) <> Sheets("volume limit").Range("J" & 2).Value
        backlog_num = backlog_num + 1
        Set Data = Table.Children(backlog_num + 1)
        Sheets("volume limit").Range("J" & 2).Value = Format(Data.Children(0).Children(0).innerText, "M/d/yyyy")
    Wend
End Sub



'Creates a border at the end of the week by checking if at least two full days have passed since the last entry and then updates the four week range
Public Sub SeparateWeeks(ByRef four_week_start As Integer, ByRef four_week_end As Integer, row As Integer)
    If Sheets("volume limit").Range("A" & row).Value - Sheets("volume limit").Range("A" & row - 1).Value > 2 Then
        Sheets("volume limit").Range("A" & row & ":F" & row).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        FourWeekRange four_week_start:=four_week_start, four_week_end:=four_week_end, row:=row
    End If
End Sub



'Prints the volume limit of the next day in which the market will be open, only useful for the start of a new week
Public Sub LimitForecast(ByRef four_week_start As Integer, ByRef four_week_end As Integer, row As Integer)
    If DateTime.Date - Sheets("volume limit").Range("A" & row - 1) > 2 Then
        Sheets("volume limit").Range("A" & row & ":F" & row).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        FourWeekRange four_week_start:=four_week_start, four_week_end:=four_week_end, row:=row
    End If
    Sheets("volume limit").Range("G" & row).Formula = "=ROUND(AVERAGE($F$" & four_week_start & ":$F$" & four_week_end & ")*0.25,-2)"
    Sheets("volume limit").Range("G" & row - 1).Copy
    Sheets("volume limit").Range("G" & row).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub

