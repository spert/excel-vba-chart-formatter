
Option Explicit On

Public Function ConvertTo2DArray(DoubleInput As Variant) As Variant

    Dim i As Integer, j As Integer, x As Integer, y As Integer
    Dim arrResult() As Variant

    If Not IsArray(DoubleInput) Then

        ReDim Preserve arrResult(1 To 1, 1 To 1)

        arrResult(1, 1) = DoubleInput

        ConvertTo2DArray = arrResult

        Exit Function

    End If

    x = UBound(DoubleInput, 1) - LBound(DoubleInput, 1) + 1

    On Error Resume Next

    y = UBound(DoubleInput, 2) - LBound(DoubleInput, 2) + 1



    If (x > 0 And y = 1) Then

        arrResult = DoubleInput

        ConvertTo2DArray = arrResult

        Exit Function

    End If


    If (x > 0 And y = 0) Or x = y Then

        ReDim Preserve arrResult(1 To x, 1 To 1)

        For i = 1 To x

            arrResult(i, 1) = DoubleInput(i)

        Next

        ConvertTo2DArray = arrResult

        Exit Function

    End If


    If x = 1 And y > 0 Then

        ReDim Preserve arrResult(1 To y, 1 To 1)

        For i = 1 To y

            arrResult(i, 1) = DoubleInput(1, i)

        Next

        ConvertTo2DArray = arrResult

        Exit Function

    End If

    ConvertTo2DArray = arrResult

End Function

Function IsDateAxis(ax As axis) As Boolean

    Dim typ As Long

    typ = -1

    On Error Resume Next

    typ = ax.MajorUnitScale

    On Error GoTo 0

    IsDateAxis = (typ > -1)

End Function

Function IsLineChart(ChartType As XlChartType) As Boolean

    Dim lineChart(10) As Integer
    lineChart(1) = 4
    lineChart(2) = 65
    lineChart(3) = 66
    lineChart(4) = 67
    lineChart(5) = 63
    lineChart(6) = 64 '

    Dim i As Integer

    For i = 1 To UBound(lineChart)

        If ChartType = lineChart(i) Then

            IsLineChart = True
            Exit Function

        End If

    Next

    IsLineChart = False

End Function

Function IsColumnBarAreaChart(ChartType As XlChartType) As Boolean

    Dim columnChart(10) As Integer
    columnChart(1) = 51
    columnChart(2) = 52
    columnChart(3) = 53
    columnChart(4) = 57
    columnChart(5) = 58
    columnChart(7) = 1
    columnChart(8) = 76
    columnChart(9) = 77
    columnChart(10) = 59

    Dim i As Integer

    For i = 1 To UBound(columnChart)

        If ChartType = columnChart(i) Then

            IsColumnBarAreaChart = True
            Exit Function

        End If

    Next

    IsColumnBarAreaChart = False

End Function




Public Sub FormatRange(MyRange As Range)

    Dim rF As Range

    For Each rF In MyRange

        rF.Borders.Color = RGB(255, 255, 255)
        rF.ColumnWidth = 21.45

        With rF.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(217, 225, 242)
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

    Next

End Sub

'Public Function GetSeriesColor(index As Integer, ColOrWei) As Long
'
'     Dim color As Long
'     Dim Weight As Long
'
'     Select Case index
'        Case 1
'          color = RGB(223, 158, 48)
'          Weight = 2
'        Case 2
'          color = RGB(190, 82, 3)
'          Weight = 2
'        Case 3
'          color = RGB(44, 68, 94)
'          Weight = 2
'        Case 4
'          color = RGB(46, 108, 129)
'          Weight = 2
'        Case 5
'          color = RGB(100, 178, 199)
'          Weight = 2
'        Case Else
'          color = RGB(100, 178, 199)
'          Weight = 2
'    End Select
'
'     Select Case UCase(ColOrWei)
'        Case "COLOR"
'            GetSeriesColor = color
'        Case "WEIGHT"
'            GetSeriesColor = Weight
'    End Select
'
'End Function

Public Function GetChartObject(ob As Object) As ChartObject

    Dim oSelection As Object
    Set oSelection = ob

    Do While TypeName(oSelection) <> "Application"

        If TypeName(oSelection) = "ChartObject" Then
           
           Set GetChartObject = oSelection
           Exit Function

        Else
        
            Set oSelection = oSelection.Parent
            
        End If

    Loop
    
    Set GetChartObject = Nothing

End Function


Function SERIESFUNC(Optional n, Optional cat, Optional Vals, Optional order, Optional BubSize) As Variant

    'Code source in this module is from: Excel 2010 Power Programming with VBA By John Walkenbach, page 640
    Dim i As Integer
    Dim result(1 To 5, 1 To 6) As Variant
    result(1, 1) = "Empty"
    result(2, 1) = "Empty"
    result(3, 1) = "Empty"
    result(4, 1) = "Empty"
    result(5, 1) = "Empty"
    '    Result(1, 2) = ""
    '    Result(2, 2) = ""
    '    Result(3, 2) = ""
    '    Result(4, 2) = ""
    '    Result(5, 2) = ""


    If IsMissing(n) Then

        result(1, 1) = "String"
        result(1, 2) = ""

    ElseIf TypeName(n) = "Range" Then

        result(1, 1) = "Range_single"
        result(1, 2) = n.Areas(1).Address(, , , True)

    ElseIf TypeName(n) = "String" Then

        result(1, 1) = "String"
        result(1, 2) = n
        'Debug.Print ("string")

    ElseIf TypeName(n) = "Error" Then 'Error = Closed external workbook

        result(1, 1) = "Closed_external_workbook"
        result(1, 2) = ""

    End If


    If IsMissing(cat) Then

        result(2, 1) = "Empty"

    ElseIf TypeName(cat) = "Range" Then

        If cat.Areas.count = 1 Then

            result(2, 1) = "Range_single"
            result(2, 2) = cat.Areas(1).Address(, , , True)

        Else

            result(2, 1) = "Range_multiple"
            For i = 1 To cat.Areas.count
                result(2, 2) = result(2, 2) & cat.Areas(i).Address(, , , True)
                If i <> cat.Areas.count Then result(2, 2) = result(2, 2) & "&&"
            Next i

        End If

    ElseIf TypeName(cat) = "Variant()" Then

        Dim bString As Boolean
        bString = False

        result(2, 1) = "Variant()"

        For i = LBound(cat) To UBound(cat)

            result(2, 2) = result(2, 2) & cat(i)

            If i <> UBound(cat) Then result(2, 2) = result(2, 2) & "&&"

        Next i

    ElseIf TypeName(cat) = "String" Then

        result(2, 1) = "String"
        result(2, 2) = ""

    ElseIf TypeName(cat) = "Error" Then
        result(2, 1) = "Closed_external_workbook"
        result(2, 2) = ""
    End If

    '    Debug.Print (TypeName(Vals))

    If IsMissing(Vals) Then
        result(3, 1) = "Empty"
    ElseIf TypeName(Vals) = "Range" Then

        If Vals.Areas.count = 1 Then

            result(3, 1) = "Range_single"
            result(3, 2) = Vals.Areas(1).Address(, , , True)

        Else

            result(3, 1) = "Range_multiple"
            For i = 1 To Vals.Areas.count
                result(3, 2) = result(3, 2) & Vals.Areas(i).Address(, , , True)
                If i <> Vals.Areas.count Then result(3, 2) = result(3, 2) & "&&"
            Next i

        End If

    ElseIf TypeName(Vals) = "Variant()" Then

        result(3, 1) = "Variant()"
        'Result(3, 2) = Result(3, 2) & "{"

        For i = LBound(Vals) To UBound(Vals)
            result(3, 2) = result(3, 2) & Vals(i)
            If i <> UBound(Vals) Then result(3, 2) = result(3, 2) & "&&"
        Next i

        'Result(3, 2) = Result(3, 2) & "}"

    ElseIf TypeName(Vals) = "String" Then

        result(3, 1) = "String"
        result(3, 2) = ""

    ElseIf TypeName(Vals) = "Error" Then

        result(3, 1) = "Closed_external_workbook"
        result(3, 2) = ""

    End If

    If IsMissing(order) Then

        result(4, 1) = "Empty"

    ElseIf TypeName(order) = "Double" Then

        result(4, 1) = "Double"
        result(4, 2) = order

    End If

    SERIESFUNC = result

End Function

Sub ApplyRecentColors(ChartProps As MyChart)
    'PURPOSE: Use A List Of RGB Codes To Load Colors Into Recent Colors Section of Color Palette
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    ActiveWorkBook.ActiveSheet.Cells(1, 1).Select

    Dim x As Long
    Dim CurrentFill As Variant

    'Array List of RGB Color Codes to Add To Recent Colors Section (Max 10)
    'ColorList = Array(RGB(255, 215, 57), RGB(133, 138, 140), RGB(0, 121, 188), RGB(138, 204, 97), RGB(255, 141, 67), RGB(244, 145, 212), RGB(233, 158, 48), RGB(190, 82, 33), RGB(100, 178, 199))

    'Store ActiveCell's Fill Color (if applicable)
    If ActiveCell.Interior.ColorIndex <> xlNone Then CurrentFill = ActiveCell.Interior.Color

    'Optimize Code
    Application.ScreenUpdating = False

    'Loop Through List Of RGB Codes And Add To Recent Colors
    For x = LBound(ChartProps.SeriesColors) To UBound(ChartProps.SeriesColors)
        ActiveCell.Interior.Color = ChartProps.SeriesColors(x)
        DoEvents
        'Application.SendKeys ("")
        SendKeys "%hhm~"
    DoEvents
    Next x

    'Return ActiveCell Original Fill Color
    If CurrentFill = Empty Then
        ActiveCell.Interior.ColorIndex = xlNone
    Else
        ActiveCell.Interior.Color = CurrentFill
    End If

End Sub




