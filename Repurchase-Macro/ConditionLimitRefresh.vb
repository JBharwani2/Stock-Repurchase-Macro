Imports mshtml

Module ConditionLimitRefresh
    Sub ConditionLimitRefresh()
        '
        ' ConditionLimitRefresh Macro
        ' Created by Jeremy Bharwani on 5/27/21
        ' (questions- email jcb926@gmail.com)
        '
        '   This macro can be used to update the volume limit sheet with all data since the last time it was updated. It
        '   disregards weekends and holidays and divides each week with a border. Equations and formatting is also
        '   done automatically as needed. If sheet is already up to date, nothing will change.
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
        FindEndOfData(row:=row)
        FourWeekRange(four_week_start:=four_week_start, four_week_end:=four_week_end, row:=row)
    
 
        'CONNECTS TO WEBPAGE ---------------------------------------------------------------------------------------------
            Set http = CreateObject("MSXML2.XMLHTTP")
            http.Open "GET", "https://finance.yahoo.com/quote/CPSS/history", False
            http.send
        html.body.innerHTML = http.responseText
    
 
        'DATA GRAB --------------------------------------------------------------------------------------------------------
            'Scrapes data from top row of yahoo finance's historical data page
            Set Table = html.getElementsByClassName("BdT Bdc($seperatorColor) Ta(end) Fz(s) Whs(nw)")
            Set Data = Table(0)
            backlog_num = -1

        'Sets place-holder
        Sheets("volume limit").Range("J" & 2).Value = Format(Data.Children(0).innerText, "M/d/yyyy")
        'Uses place-holder
        GetBacklog(Table:=Table, backlog_num:=backlog_num, row:=row)
        'Clears place-holder
        Sheets("volume limit").Range("J" & 2).Value = ""

        'Fills in each column with values or formulas (loop depends on number of un-updated days since last update)
        While backlog_num >= 0
            Set Data = Table(backlog_num)
    
            Sheets("volume limit").Range("A" & row).Value = Format(Data.Children(0).innerText, "M/d/yyyy")
            SeparateWeeks(four_week_start:=four_week_start, four_week_end:=four_week_end, row:=row)

            Sheets("volume limit").Range("B" & row).Value = Data.Children(1).innerText
            Sheets("volume limit").Range("C" & row).Value = Data.Children(2).innerText
            Sheets("volume limit").Range("D" & row).Value = Data.Children(3).innerText
            Sheets("volume limit").Range("E" & row).Value = Data.Children(4).innerText
            Sheets("volume limit").Range("F" & row).Value = Data.Children(6).innerText
            Sheets("volume limit").Range("G" & row).Formula = "=ROUND(AVERAGE($F$" & four_week_start & ":$F$" & four_week_end & ")*0.25,-2)"
            Sheets("volume limit").Range("H" & row).Formula = "=IFERROR(VLOOKUP($A" & row & ",Activity!$A$7:$I$801,3,FALSE),0)"
            Sheets("volume limit").Range("I" & row).Formula = "=IF(H" & row & "<G" & row & "," & Chr(34) & Chr(34) & ",+H" & row & "-G" & row & ")"

            'Formats newly input data to match with previous
            Sheets("volume limit").Range("A" & row & ":I" & row).Font.Name = "Arial"
            Sheets("volume limit").Range("B" & row & ":E" & row).NumberFormat = "0.00"
            Sheets("volume limit").Range("F" & row & ":I" & row).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"
            Sheets("volume limit").Range("I" & row).Font.Color = RGB(255, 0, 0)

            'Iteration through each row that needs to be updated
            row = row + 1
            backlog_num = backlog_num - 1
        Wend
    
        MsgBox("Volume Limit Sheet has been updated to current date")
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
            Set Data = Table(backlog_num + 1)
            Sheets("volume limit").Range("J" & 2).Value = Format(Data.Children(0).innerText, "M/d/yyyy")
        Wend
    End Sub



    'Creates a border at the end of the week by checking if at least two full days have past since last entry and then updates the four week range
    Public Sub SeparateWeeks(ByRef four_week_start As Integer, ByRef four_week_end As Integer, row As Integer)
        If Sheets("volume limit").Range("A" & row).Value - Sheets("volume limit").Range("A" & row - 1).Value > 2 Then
            Sheets("volume limit").Range("A" & row & ":F" & row).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            FourWeekRange four_week_start:=four_week_start, four_week_end:=four_week_end, row:=row
        End If
    End Sub

End Module
