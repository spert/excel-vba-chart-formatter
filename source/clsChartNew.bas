
Option Explicit On

Const cModule = "clsChart"

Private oOldChart As Chart

Private tChartProps As MyChart

Private shtNewWorkSheet As Worksheet

Private oNewChart As ChartObject

Private sTitle() As String

Private arrMySeries() As clsSeries

Private arrMyAxes() As clsAxes

Private bMultipleAxesGroups As Boolean

Private bIsValidExcelChart As Boolean


Public Sub Initiate(OldChart As Chart, ChartProps As MyChart)
  
  Set oOldChart = OldChart
  tChartProps = ChartProps
    bMultipleAxesGroups = False

End Sub

Public Property Let SetNewWorkSheet(NewWorkSheet As Worksheet)

    Set shtNewWorkSheet = NewWorkSheet

End Property


Public Property Let SetMyAxes(MyAxes() As clsAxes)

   arrMyAxes = MyAxes

End Property

Public Property Let SetMySeries(MySeries() As clsSeries)

  arrMySeries = MySeries

End Property

Public Property Get GetNewChart() As Chart

    Set GetNewChart = oNewChart.Chart

End Property


Public Sub PrintChartNumber()

    Const cProc = "PrintChartNumber"

    On Error GoTo ErrorHandler

    Dim rngNumber As Range
   Set rngNumber = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(1, 1)
   rngNumber.HorizontalAlignment = xlLeft

    rngNumber.Formula = "=" & "RIGHT(@CELL(" & Chr(34) & "filename" & Chr(34) & ",B1),LEN(@CELL(" & Chr(34) & "filename" & Chr(34) & ",B1))-FIND(" & Chr(34) & "]" & Chr(34) & ",@CELL(" & Chr(34) & "filename" & Chr(34) & ",B1)))"

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub

Public Sub PrintChartTitle()

    Const cProc = "PrintChartTitle"

    On Error GoTo ErrorHandler

    Dim rngHeadingLink As Range
   Set rngHeadingLink = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(2, 5)

   Dim rngHeading As Range
   Set rngHeading = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(2, 1)

   Dim rngNumber As Range
   Set rngNumber = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(1, 1)

   rngHeadingLink.Formula = "=" & Chr(34) & tChartProps.Language_Descriptions(7) & Chr(34) & "& " & rngNumber.Address(, , , False, False) & "&" & Chr(34) & ". " & Chr(34) & "&" & rngHeading.Address(, , , False, False)
    rngHeadingLink.Font.Bold = True

    Dim fTitleBox As Shape
    On Error Resume Next
   Set fTitleBox = oNewChart.Chart.Shapes("ChartFormatterTitleBox")
        
   If Not fTitleBox Is Nothing Then
        fTitleBox.OLEFormat.Object.Formula = rngHeadingLink.Address(, , , True)
    End If

    If oNewChart.HasTitle Then
        oNewChart.Chart.ChartTitle.Formula = "=" & rngHeadingLink.Address(, , , True)
    End If


    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub


Public Sub PrintSourceTextBox()

    Const cProc = "PrintSourceTextBox"

    On Error GoTo ErrorHandler

    On Error GoTo ErrorHandler

    Dim fSources As Shape
    On Error Resume Next
    Set fSources = oNewChart.Chart.Shapes("ChartFormatterSourceBox")
        
    If fSources Is Nothing Then
        Exit Sub
    End If

    Dim rngSourceLink As Range
    Set rngSourceLink = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(3, 5)
    
    Dim rngSource As Range
    Set rngSource = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(3, 1)
    rngSourceLink.Formula = "=IF(ISNUMBER(SEARCH(" & Chr(34) & "," & Chr(34) & "," & rngSource.Address(, , , False, False) & "))," & Chr(34) & tChartProps.Language_Descriptions(6) & Chr(34) & "," & Chr(34) & tChartProps.Language_Descriptions(5) & Chr(34) & ")&" & rngSource.Address(, , , False, False)
    fSources.OLEFormat.Object.Formula = rngSourceLink.Address(, , , True)

    Exit Sub
ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


End Sub


Public Sub PrintChartAxisTitle()

    Const cProc = "PrintChartAxisTitle"

    Dim rngLeftAxisTitle As Range
    Set rngLeftAxisTitle = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(4, 1)
    
    Dim rngRightAxisTitle As Range
    Set rngRightAxisTitle = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(5, 1)

    Dim ax As axis

    For Each ax In oNewChart.Chart.Axes

        If ax.HasTitle = True And ax.Type = xlValue And ax.AxisGroup = xlPrimary Then

            rngLeftAxisTitle = ax.AxisTitle.Text

            Call FormatRange(rngLeftAxisTitle)

            ax.AxisTitle.Formula = "=" & rngLeftAxisTitle.Address(, , , True)

        End If

        If ax.HasTitle = True And ax.Type = xlValue And ax.AxisGroup = xlSecondary Then

            rngRightAxisTitle = ax.AxisTitle.Text

            Call FormatRange(rngRightAxisTitle)

            ax.AxisTitle.Formula = "=" & rngRightAxisTitle.Address(, , , True)

        End If

    Next

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

 End Sub


Public Sub AssignNewRanges()

    Const cProc = "AssignNewRangesToAxes"

    Dim iColumn As Integer, i As Integer

    On Error GoTo ErrorHandler

    If IsEmpty(arrMyAxes) = True Then

        GoTo NoAxesAvailable

    End If


    iColumn = 0

    For i = 1 To UBound(arrMyAxes)

        arrMyAxes(i).SetColumn = iColumn
        arrMyAxes(i).SetNewWorkSheet = shtNewWorkSheet
        Call arrMyAxes(i).SetNewCategoryRange
        Call arrMyAxes(i).PrintCategoryValues

        iColumn = iColumn + 1

    Next

NoAxesAvailable:

    For i = 1 To UBound(arrMySeries)

        arrMySeries(i).SetColumn = iColumn
        arrMySeries(i).SetNewWorkSheet = shtNewWorkSheet
        Call arrMySeries(i).SetNewCategoryRange
        Call arrMySeries(i).SetNewValueRange
        Call arrMySeries(i).SetNewNameRange
        Call arrMySeries(i).SetNewNameLinkRange

        Call arrMySeries(i).PrintSeriesName(i)
        Call arrMySeries(i).PrintSeriesNameLink
        Call arrMySeries(i).PrintSeriesValues
        Call arrMySeries(i).PrintSeriesScale(bMultipleAxesGroups)

        iColumn = iColumn + 1

    Next

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


 End Sub

Public Sub MapNewSeries()

    Const cProc = "MapNewSeries"

    On Error GoTo ErrorHandler

    Dim i As Integer

    For i = 1 To oNewChart.Chart.SeriesCollection.count


        Dim serS As Series
       Set serS = oNewChart.Chart.SeriesCollection(i)
        
       If Not StrComp(arrMySeries(i).GetMySources(1, 1), "Empty", vbBinaryCompare) = 0 Then

            serS.Name = "=" & arrMySeries(i).GetMySources(1, 6)

        End If

        If Not StrComp(arrMySeries(i).GetMySources(2, 1), "Empty", vbBinaryCompare) = 0 Then

            serS.XValues = "=" & arrMySeries(i).GetMySources(2, 5)

        End If

        If Not StrComp(arrMySeries(i).GetMySources(3, 1), "Empty", vbBinaryCompare) = 0 Then

            serS.Values = "=" & arrMySeries(i).GetMySources(3, 5)

        End If

    Next

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


End Sub


Public Sub CopyOldChartToNewWorksheet()

    Const cProc = "CopyOldChartToNewWorksheet"

    On Error GoTo ErrorHandler : 

Set oNewChart = shtNewWorkSheet.ChartObjects.Add(tChartProps.Left, tChartProps.Top, tChartProps.With, tChartProps.Height) '(Left, Top, Width, Height)

oOldChart.ChartArea.Copy
    oNewChart.Activate
    ActiveChart.Paste

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


End Sub


Public Sub ToString()

    Dim i As Integer

    For i = 1 To UBound(arrMySeries)

        Call arrMySeries(i).ToString

    Next

End Sub


