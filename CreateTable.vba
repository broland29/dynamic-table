
' inspiration : https://stackoverflow.com/questions/45274764/creating-a-table-based-upon-number-of-rows-columns-entered/45288195#45288195
' data validation done through plain excel : https://www.extendoffice.com/documents/excel/5938-excel-data-validation-allow-only-numbers.html
' worksheet stuff: https://stackoverflow.com/questions/6310240/switching-sheets-in-excel-vba
' grid: https://stackoverflow.com/questions/52417729/excel-vba-help-for-grid-lines
' x : number of students
' y : number of problems

Sub CreateTable()
    
    ' password to protect sheet (see: end of script)
    Dim pass As String
    pass = "1234"
    
    For Each sht In ThisWorkbook.Sheets
        sht.Unprotect password:=pass
    Next sht
    
    ' swap to table
    Sheets("Táblázat").Activate
    
    Dim ws As Worksheet
    Set ws = Sheets("Beállítások")

    ' get new values
    Dim x As Integer
    Dim y As Integer
    x = ws.range("D3").Value
    y = ws.range("D4").Value

    ' get old values
    Dim oldX As Integer
    Dim oldY As Integer
    oldX = ws.range("B5").Value
    oldY = ws.range("C5").Value

    
    
    Set ws = Sheets("Táblázat")
    
    ' delete rows
    For i = 1 To oldX + 1
        ws.range("B6").EntireRow.Delete
    Next i

    ' insert rows
    ws.range("B6").EntireRow.Insert
    ws.range("C6:E6").Merge
    ws.Cells(6, 3).Value = "Átlagok"
    
    For i = 1 To x
        ws.range("B6").EntireRow.Insert
        ws.range("C6:E6").Merge             ' cell for name
        ws.Cells(6, 2) = "D" & x - i + 1    ' cell for index
    Next i
    
    ' only modify columns if needed (since it erases problem-related data)
    'If oldY <> y Then - TODO: find better solution
        ' delete all columns
        For i = 1 To oldY + 4
            ws.range("F4").EntireColumn.Delete
        Next i
        
        ' insert columns for statistics
        For i = 1 To 4
            ws.range("F4").EntireColumn.Insert
        Next i
        
        ' insert columns for grades
        For i = 1 To y
            ws.range("F4").EntireColumn.Insert
            ws.Cells(4, 6) = "F" & y - i + 1
        Next i
    'End If
    
    ' m - the column of sum
    ' Chr(64 + m) is letter of column m - ascii of A is 65
    Dim m As Integer
    m = 6 + y
    
    ws.Cells(4, m) = "Össz"
    ws.Cells(4, m + 1) = "Százalék"
    ws.Cells(4, m + 2) = "Jegy"
    ws.Cells(4, m + 3) = "Osztályzat"
    
    '' insert functions ''
    ' max sum
    ' example: =SUM(F5:K5)
    ws.Cells(5, m).Formula = "=SUM(F5:" & Chr(63 + m) & "5)"
            
    For i = 1 To x
        ' current row
        Dim n As Integer
        n = i + 5
        
        ' sum
        ' example: =SUM(F6:K6)
        ws.Cells(n, m).Formula = "=SUM(F" & n & ":" & Chr(63 + m) & n & ")"
        
        ' percentage
        ' example: =ROUND(L6/L5*100,2)
        ws.Cells(n, m + 1).Formula = "=ROUND(" & Chr(64 + m) & n & "/" & Chr(64 + m) & "5*100,2)"
            
        ' grade
        ' example: =ROUND(M6,-1)/10
        ws.Cells(n, m + 2).Formula = "=ROUND(" & Chr(65 + m) & n & ",-1)/10"
            
        ' mark
        ' example: =IF(M6>=Beállítások!E6,"FB",IF(M6>=Beállítások!E7,"B",IF(M6>=Beállítások!E8,"S","I")))
        ws.Cells(n, m + 3).Formula = "=IF(" _
            & Chr(65 + m) _
            & n _
            & ">=Beállítások!E6,""FB"",IF(" _
            & Chr(65 + m) _
            & n _
            & ">=Beállítások!E7,""B"",IF(" _
            & Chr(65 + m) _
            & n _
            & ">=Beállítások!E8,""S"",""I"")))"
    Next i
    
    ' class average for each separete problem + add constraint
    ' example: =ROUND(AVERAGE(F6:F15),2)
    

    For i = 1 To y
        ws.Cells(n + 1, 5 + i).Formula = "=ROUND(AVERAGE(" & Chr(69 + i) & "6:" & Chr(69 + i) & 5 + x & "),2)"
        With range(Chr(69 + i) & "5").Validation
            .Delete
            .Add Type:=xlValidateWholeNumber, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlGreater, Formula1:="0"
            .IgnoreBlank = True
            .InputTitle = "Maximális pont"
            .ErrorTitle = "Helytelen adat"
            .InputMessage = "Egy pozitív, nem nulla érték."
            .ErrorMessage = "A maximális pont nagyobb mint 0!"
        
        End With
            
        With range(Chr(69 + i) & "6:" & Chr(69 + i) & 5 + x).Validation
            .Delete
            .Add Type:=xlValidateWholeNumber, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, Formula1:="0", Formula2:="=" & Chr(69 + i) & "$5"
            .InputTitle = "Részpont"
            .ErrorTitle = "Helytelen adat"
            .InputMessage = "Egy érték 0 és maximális pont között."
            .ErrorMessage = "A részpont egy érték 0 és maximális pont között!"
        End With
        
    Next i
    
    For i = 1 To x
        With range("C" & 5 + i).Validation
            .Delete
            .Add Type:=xlValidateCustom, _
                AlertStyle:=xlValidAlertStop, _
                Formula1:="=ISTEXT(C" & 5 + i & ")"
            .IgnoreBlank = True
            .InputTitle = "Diák neve"
            .ErrorTitle = "Helytelen adat"
            .InputMessage = "A diák neve csak karakterekbõl állhat."
            .ErrorMessage = "A diák neve csak karakterekbõl állhat!"
        End With
    Next i
    
    ' sum
    ' example: =SUM(F16:K16)
    ws.Cells(n + 1, m).Formula = "=SUM(" & Chr(70) & n + 1 & ":" & Chr(69 + y) & n + 1 & ")"

    ' percent
    ' example: =ROUND(L16/L5*100,2)
    ws.Cells(n + 1, m + 1).Formula = "=ROUND(" & Chr(70 + y) & n + 1 & "/" & Chr(70 + y) & "5*100,2)"

    ' grade
    ' example: =ROUND(M16,-1)/10
    ws.Cells(n + 1, m + 2).Formula = "=ROUND(" & Chr(65 + m) & n + 1 & ",-1)/10"

    ' mark
    ' =IF(M16>=Beállítások!E6,"FB",IF(M16>=Beállítások!E7,"B",IF(M16>=Beállítások!E8,"S","I")))
    ws.Cells(n + 1, m + 3).Formula = "=IF(" _
            & Chr(65 + m) _
            & n + 1 _
            & ">=Beállítások!E6,""FB"",IF(" _
            & Chr(65 + m) _
            & n + 1 _
            & ">=Beállítások!E7,""B"",IF(" _
            & Chr(65 + m) _
            & n + 1 _
            & ">=Beállítások!E8,""S"",""I"")))"

    ' draw borders
    Dim area As range
    Dim aux As String
    aux = "B4:" & Chr(67 + m) & (n + 1)
    
    For Each area In ws.range(aux).Areas
        With area.Borders
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
    Next

    ' set alignment
    aux = "F4:" & Chr(67 + m) & n + 1
    range(aux).Select
    Selection.VerticalAlignment = xlCenter
    Selection.HorizontalAlignment = xlCenter

    ' set print area
    With ws.PageSetup
        .PrintArea = "$A$3:$" & Chr(74 + y) & "$" & 7 + x
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Orientation = xlLandscape
    End With

    
    
    Set ws = Sheets("Beállítások")
    
    ' update old values
    ws.Cells(5, 2).Value = x
    ws.Cells(5, 3).Value = y
    
    ' "deselect" previous selection
    range("A1").Select
    
    
    ' set editability

    For Each sht In ThisWorkbook.Sheets
        sht.Protect password:=pass, Userinterfaceonly:=True
    Next sht
    
    
    ws.range("D3", "D4").Locked = False
    ws.range("D3", "D4").Interior.Color = RGB(255, 249, 227)
    
    ws.Cells(1, 1).Formula = "=Táblázat!" & Chr(70 + y) & "5"
    
    Set ws = Sheets("Táblázat")
    
    ws.range("C6:E" & 5 + x).Locked = False
    ws.range("C6:E" & 5 + x).Interior.Color = RGB(255, 249, 227)
    
    ws.range("F5:" & Chr(69 + y) & 5 + x).Locked = False
    ws.range("F5:" & Chr(69 + y) & 5 + x).Interior.Color = RGB(255, 249, 227)
    
End Sub
