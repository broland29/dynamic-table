
' https://www.mrexcel.com/board/threads/vba-cannot-identify-chart-even-using-exact-sequence-from-recorded-macro.845832/

Sub UpdateCharts()

    ' worksheets
    Dim src As Worksheet
    Set src = Sheets("Táblázat")
    Dim dst As Worksheet
    Set dst = Sheets("Gráfok")
    Dim sts As Worksheet
    Set sts = Sheets("Beállítások")
    
    '' line chart ''
    
    ' number of students and problems
    Dim x As Integer
    Dim y As Integer
    x = sts.range("D3").Value
    y = sts.range("D4").Value
    
    
    dst.range("A1:A" & y).Locked = False
    
    ' row index of "whole"
    Dim r1 As Integer
    r1 = 5
    
    ' row index of "part"
    Dim r2 As Integer
    r2 = 6 + x
    
    ' invisible text hehe
    dst.range("A1:A" & y).Font.Color = RGB(255, 255, 255)
    
    For i = 1 To y
        
        If src.Cells(r1, 5 + i) = 0 _
                Or IsEmpty(src.Cells(r1, 5 + i)) _
                Or IsError(src.Cells(r1, 5 + i)) _
                Or IsEmpty(src.Cells(r2, 5 + i)) _
                Or IsError(src.Cells(r2, 5 + i)) Then
            dst.Cells(i, 1) = 0
        Else
            dst.Cells(i, 1) = src.Cells(r2, 5 + i) / src.Cells(r1, 5 + i) * 100
        End If
    Next i
    
    
    Dim cht As Chart
    Set cht = dst.ChartObjects("Chart 2").Chart
    cht.SetSourceData Source:=dst.range("A1:A" & y)

    '' pie chart ''
    
    Dim countFB, countB, countS, countI As Integer
    countFB = 0
    countB = 0
    countS = 0
    countI = 0
    
    Dim c As String
    For i = 1 To x
        
        If IsError(src.Cells(5 + i, 9 + y)) Or IsEmpty(src.Cells(5 + i, 9 + y)) Then
            'nothing
        Else
            c = src.Cells(5 + i, 9 + y)
            If StrComp(c, "FB") = 0 Then
                countFB = countFB + 1
            ElseIf StrComp(c, "B") = 0 Then
                countB = countB + 1
            ElseIf StrComp(c, "S") = 0 Then
                countS = countS + 1
            ElseIf StrComp(c, "I") = 0 Then
                countI = countI + 1
            End If
        End If
        
        
    Next i

    dst.Cells(35, 7).Value = countFB
    dst.Cells(36, 7).Value = countB
    dst.Cells(37, 7).Value = countS
    dst.Cells(38, 7).Value = countI
    
    dst.range("A1:A" & y).Locked = True
End Sub
