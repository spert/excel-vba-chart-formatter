
Option Explicit On

Const cModule = "clsSeries"

Private oSeries As Series
Private arrSources As Variant
Private clsMyAxis As clsAxes
Private oOldChart As Chart
Private tChartProps As MyChart
Private xlsAxisGroup As XlAxisGroup
Private shtNewWorkSheet As Worksheet
Private iColumn As Integer
Private bIsValidSeries As Boolean

Public Sub Initiate(ExcelSeries As Series, OldChart As Chart)
  
  Set oSeries = ExcelSeries
  Set oOldChart = OldChart
  arrSources = SplitFormula(ExcelSeries)
    xlsAxisGroup = ExcelSeries.AxisGroup


End Sub

Public Property Let SetSeriesAxis(MyAxis As clsAxes)

    Set clsMyAxis = MyAxis

End Property

Public Property Get GetSeriesAxis() As clsAxes

    Set GetSeriesAxis = clsMyAxis

End Property

Public Property Get GetSeriesAxisGroup() As XlAxisGroup

        GetSeriesAxisGroup = xlsAxisGroup
        
End Property

Public Property Let SetColumn(Column As Integer)

        iColumn = Column
        
End Property

Public Property Let SetChartProps(ChartProps As MyChart)

   tChartProps = ChartProps

End Property


Public Property Get IsValidSeries() As Boolean

        IsValidSeries = bIsValidSeries
        
End Property

Public Property Let SetNewWorkSheet(NewWorkSheet As Worksheet)

    Set shtNewWorkSheet = NewWorkSheet

End Property

Public Sub SetNewCategoryRange()

    If clsMyAxis Is Nothing Then

        Exit Sub

    End If

    If Not IsEmpty(clsMyAxis.GetCategoryArray(5)) Then

        arrSources(2, 5) = clsMyAxis.GetCategoryArray(5)

    End If

End Sub

Public Sub SetNewValueRange()

    Const cProc = "SetNewValueRange"

    On Error GoTo ErrorHandler

    Dim firstCell As Range
    Set firstCell = shtNewWorkSheet.Cells(tChartProps.SeriesDataOffset(1, 1), tChartProps.SeriesDataOffset(2, 1) + iColumn)
        
    Dim r As Range
    Set r = firstCell.Resize(UBound(arrSources(3, 3)), 1)
    Call FormatRange(r)

    arrSources(3, 5) = r.Address(, , , True)

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub

Public Sub SetNewNameRange()

    Const cProc = "SetNewNameRange"

    On Error GoTo ErrorHandler

    Dim firstCell As Range
    Set firstCell = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(6, 1 + iColumn)
        
    arrSources(1, 5) = firstCell.Address(, , , True)

    Call FormatRange(firstCell)

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub

Public Sub SetNewNameLinkRange()

    Const cProc = "SetNewNameLinkRange"

    On Error GoTo ErrorHandler

    Dim firstCell As Range
    Set firstCell = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(8, 1 + iColumn)
        
    arrSources(1, 6) = firstCell.Address(, , , True)

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub

Public Sub PrintSeriesName(SeriesNumber As Integer)

    Const cProc = "PrintSeriesName"

    On Error GoTo ErrorHandler

    If IsEmpty(arrSources(1, 5)) Then
        Exit Sub
    End If

    Dim r As Range
    Set r = Range(arrSources(1, 5))
    
    Call FormatRange(r)

    If Len(Trim(arrSources(1, 3)(1, 1))) = 0 Then

        r.Value = "#" & SeriesNumber

    Else

        r.Value = arrSources(1, 3)

    End If

    r.NumberFormat = arrSources(1, 4)

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub

Public Sub PrintSeriesNameLink()

    Const cProc = "PrintSeriesNameLink"

    On Error GoTo ErrorHandler

    Dim firstCell As Range
    Set firstCell = Range(arrSources(1, 6))

    Dim nameCell As Range
    Set nameCell = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(6, 1 + iColumn)
    
    Dim scaleCell As Range
    Set scaleCell = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(7, 1 + iColumn)
    
    firstCell.Formula = "=" & nameCell.Address(False, False) & "&" & Chr(34) & " " & Chr(34) & "&" & scaleCell.Address(False, False)

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub

Public Sub PrintSeriesValues()

    Const cProc = "PrintSeriesValues"

    On Error GoTo ErrorHandler

    Dim r As Range
    Set r = Range(arrSources(3, 5))

    Call FormatRange(r)

    r = arrSources(3, 3)
    r.NumberFormat = arrSources(3, 4)

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub


Public Sub PrintSeriesValuesAsLinks()

    Const cProc = "PrintSeriesValuesAsLinks"

    On Error GoTo ErrorHandler

    Dim rSource As Range
    Set rSource = Range(arrSources(3, 2)) 'source
   
    Dim rTarget As Range
    Set rTarget = Range(arrSources(3, 5)) 'target
    Call FormatRange(rTarget)
    rTarget.NumberFormat = arrSources(3, 4)

    Dim c As Integer
    Dim sSource As String

    For c = 1 To rSource.Cells.count

        sSource = "='" & rSource(c).Parent.Name & "'!" & rSource(c).Address(External:=False)
        rTarget.Cells(c).Formula = sSource

    Next

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub


Public Sub PrintSeriesScale(bMultipleAxesGroups As Boolean)

    Const cProc = "PrintSeriesScale"

    On Error GoTo ErrorHandler

    Dim firstCell As Range
    Set firstCell = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(7, 1 + iColumn)
    
    Call FormatRange(firstCell)

    Dim strValid As String
    strValid = tChartProps.Language_ScaleValidation(0) & "," & tChartProps.Language_ScaleValidation(1)

    Call FormatRange(firstCell)

    If bMultipleAxesGroups Then
        If xlsAxisGroup = xlPrimary Then
            firstCell.Value = tChartProps.Language_ScaleValidation(0)
        Else
            firstCell.Value = tChartProps.Language_ScaleValidation(1)
        End If
    End If

    With firstCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=strValid
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = False
    End With

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub

Property Get GetMySources() As Variant

   GetMySources = arrSources

End Property

Property Let SetMySources(Sources As Variant)

   arrSources = Sources

End Property

'Code source in this module is from: Excel 2010 Power Programming with VBA By John Walkenbach, page 640

Function SplitFormula(s As Series) As Variant

    Const cProc = "SplitFormula"

    On Error GoTo ErrorHandler

    Dim ResultArray As Variant
    Dim Func As String
    Dim i As Integer

    Dim strAddress As String
    Dim rngR As Range
    Dim intI As Integer
    Dim arrS(1 To 1) As String
    Dim arrString() As String


    '  If Not IsValidExcelSeries(s) Then
    '
    '    ReDim ResultArray(1 To 5, 1 To 6) As Variant
    '
    '    'ResultArray(1, ...) Series Name
    '    'ResultArray(2, ...) Series Category (XValues)
    '    'ResultArray(3, ...) Series Values (Values)
    '    'ResultArray(4, ...) Series Order : Not implemented
    '    'ResultArray(5, ...) Series bubble : Not implemented
    '
    '    'ResultArray(..., 1) Data Type
    '    'ResultArray(..., 2) Source range or value
    '    'ResultArray(..., 3) Array of values
    '    'ResultArray(..., 4) Format
    '    'ResultArray(..., 5) New range for values
    '    'ResultArray(..., 6) Series Name Link range : Implemented in case of name only
    '
    '     For i = 1 To 5
    '
    '        ResultArray(i, 1) = "Invalid"
    '        ResultArray(i, 2) = "Invalid"
    '        arrS(1) = "Invalid"
    '        ResultArray(i, 3) = ConvertTo2DArray(arrS)
    '        ResultArray(i, 4) = "Invalid"
    '        ResultArray(i, 5) = "Invalid"
    '        ResultArray(i, 6) = "Invalid"
    '
    '     Next
    '
    '    bIsValidSeries = False
    '    SplitFormula = ResultArray
    '
    '    Exit Function
    '
    '  End If

    Func = Application.Substitute(s.Formula, "=SERIES", "'" & ThisWorkbook.Name & "'!SERIESFUNC")
    ResultArray = Evaluate(Func)


    ' =========== Names ======================

    If ResultArray(1, 1) = "Range_single" Then

        strAddress = ResultArray(1, 2)
       Set rngR = Range(strAddress)
       ResultArray(1, 3) = ConvertTo2DArray(rngR.Value)
        ResultArray(1, 4) = rngR.NumberFormat

    End If

    If ResultArray(1, 1) = "String" Then

        arrS(1) = ResultArray(1, 2)
        ResultArray(1, 3) = ConvertTo2DArray(arrS)
        ResultArray(1, 4) = "General"

    End If

    ' =========== Categories =================

    If ResultArray(2, 1) = "Range_single" Then

        strAddress = ResultArray(2, 2)
    Set rngR = Range(strAddress)
    ResultArray(2, 3) = ConvertTo2DArray(rngR.Value)
        ResultArray(2, 4) = rngR.NumberFormat


    End If

    If ResultArray(2, 1) = "Range_multiple" Then

        ResultArray(2, 3) = ConvertTo2DArray(s.XValues)
        strAddress = Left(ResultArray(2, 2), InStr(1, ResultArray(2, 2), "&&") - 1)
    Set rngR = Range(strAddress)
    ResultArray(2, 4) = rngR.NumberFormat

    End If

    If ResultArray(2, 1) = "Variant()" Then

        arrString = Split(ResultArray(2, 2), "&&")

        ResultArray(2, 3) = ConvertTo2DArray(ResultArray(2, 2))
        ResultArray(2, 4) = "General"

    End If

    If ResultArray(2, 1) = "Closed_external_workbook" Then

        arrS(0) = "Reference to a closed external workbook"
        ResultArray(2, 3) = arrS
        ResultArray(2, 4) = "General"

        On Error GoTo ExitCategory

        ResultArray(2, 3) = ConvertTo2DArray(s.XValues)

ExitCategory:

    End If


    ' =========== Values ====================

    If ResultArray(3, 1) = "Range_single" Then

        ResultArray(3, 3) = ConvertTo2DArray(s.Values)
        strAddress = ResultArray(3, 2)
    Set rngR = Range(strAddress)
    ResultArray(3, 4) = rngR.NumberFormat

    End If

    If ResultArray(3, 1) = "Range_multiple" Then

        ResultArray(3, 3) = ConvertTo2DArray(s.Values)

        strAddress = Left(ResultArray(3, 2), InStr(1, ResultArray(3, 2), "&&") - 1)
    Set rngR = Range(strAddress)
    ResultArray(3, 4) = rngR.NumberFormat

    End If

    If ResultArray(3, 1) = "Variant()" Then

        arrString = Split(ResultArray(3, 2), "&&")

        ResultArray(3, 3) = ConvertTo2DArray(arrString)
        ResultArray(3, 4) = "General"

    End If

    If ResultArray(3, 1) = "Closed_external_workbook" Then

        arrS(0) = "Reference to a closed external workbook"
        ResultArray(3, 3) = arrS
        ResultArray(3, 4) = "General"

        On Error GoTo ExitValue

        ResultArray(3, 3) = ConvertTo2DArray(s.Values)

ExitValue:

    End If

    bIsValidSeries = True

    SplitFormula = ResultArray

    Exit Function

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


End Function


Public Sub ToString()

    Const cProc = "ToString"

    'ResultArray(1, ...) Series Name
    'ResultArray(2, ...) Series Category (XValues)
    'ResultArray(3, ...) Series Values (Values)
    'ResultArray(4, ...) Series Order : Not implemented
    'ResultArray(5, ...) Series bubble : Not implemented

    'ResultArray(..., 1) Data Type
    'ResultArray(..., 2) Source range or value
    'ResultArray(..., 3) Array of values
    'ResultArray(..., 4) Format
    'ResultArray(..., 5) New range for values
    'ResultArray(..., 6) Series Name Link range : Implemented in case of name only


    Dim k As Integer, l As Integer, kk As Integer

    Debug.Print(" ========== 1. Series Name ===========")

    Debug.Print("(1,1) -------- Name Data Type: " & arrSources(1, 1))
    Debug.Print("(1,2) -------- Name Source range or value: " & arrSources(1, 2))

    For k = 1 To UBound(arrSources(1, 3))
        Debug.Print("(1,3) -------- Name Array of values [" & k & "]: " & arrSources(1, 3)(k, 1))
    Next

    Debug.Print("(1,4) -------- Name Format: " & arrSources(1, 4))
    Debug.Print("(1,5) -------- Name New range for values: " & arrSources(1, 5))
    Debug.Print("(1,6) -------- Name Series Name Link range: " & arrSources(1, 6))

    Debug.Print(" =========== 2. Series Category =============")
    Debug.Print("(2,1) -------- Category Data Type: " & arrSources(2, 1))
    Debug.Print("(2,2) -------- Category Source range or value: " & arrSources(2, 2))

    For kk = 1 To UBound(arrSources(2, 3))
        Debug.Print("(2,3) -------- Category Array of values [" & kk & "]: " & arrSources(2, 3)(kk, 1))
    Next

    Debug.Print("(2,4) -------- Category Format: " & arrSources(2, 4))
    Debug.Print("(2,5) -------- Category New range for values: " & arrSources(2, 5))
    Debug.Print("(2,6) -------- Category Series Name Link range: " & arrSources(2, 6))

    Debug.Print(" =========== 3. Series Value =============")
    Debug.Print("(3,1) -------- Value Data Type: " & arrSources(3, 1))
    Debug.Print("(3,2) -------- Value Source range or value: " & arrSources(3, 2))

    For k = 1 To UBound(arrSources(3, 3))
        Debug.Print("(3,3) -------- Value Array of values [" & k & "]: " & arrSources(3, 3)(k, 1))
    Next

    Debug.Print("(3,4) -------- Value Format: " & arrSources(3, 4))
    Debug.Print("(3,5) -------- Value New range for values: " & arrSources(3, 5))
    Debug.Print("(3,6) -------- Value Series Name Link range: " & arrSources(3, 6))


End Sub

