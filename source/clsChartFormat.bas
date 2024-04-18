Option Explicit On

Const cModule = "clsChartFormat"

Private oMyChart As Chart

Private tChartProps As MyChart


Public Sub Initiate(MyChart As Chart, ChartProps As MyChart)
  
  Set oMyChart = MyChart
  tChartProps = ChartProps

End Sub


Public Sub FormatChartSize()


    Const cProc = "FormatChartTitle"

    ' On Error GoTo ErrorHandler:

    Dim oMyChartOb As ChartObject
   Set oMyChartOb = oMyChart.Parent
   oMyChartOb.ShapeRange.Line.Visible = msoFalse
    oMyChartOb.Height = tChartProps.Height
    oMyChartOb.Width = tChartProps.With

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


End Sub


Public Sub FormatChartTitleBox()

    Const cProc = "FormatChartTitle"

    On Error GoTo ErrorHandler

    Dim fTitleBox As Shape

    On Error Resume Next
        Set fTitleBox = oMyChart.Shapes("ChartFormatterTitleBox")
        
        If fTitleBox Is Nothing Then
        Exit Sub
    End If

    fTitleBox.Fill.ForeColor.RGB = tChartProps.Title.BackgroundRGB 'tChartProps.Title.BackgroundRGB
    fTitleBox.Fill.Transparency = 0
    fTitleBox.Fill.Solid
    fTitleBox.Line.Visible = msoFalse
    fTitleBox.TextFrame2.VerticalAnchor = msoAnchorMiddle
    fTitleBox.TextFrame2.HorizontalAnchor = msoAnchorNone
    fTitleBox.TextFrame2.TextRange.Font.Bold = tChartProps.Title.Font.Bold
    fTitleBox.TextFrame2.TextRange.Font.Size = tChartProps.Title.Font.Size
    fTitleBox.TextFrame2.TextRange.Font.Name = tChartProps.Title.Font.Name
    fTitleBox.TextFrame2.MarginLeft = 5.6692913386
    fTitleBox.TextFrame2.MarginRight = 5.6692913386
    fTitleBox.TextFrame2.MarginTop = 2.8346456693
    fTitleBox.TextFrame2.MarginBottom = 2.8346456693
    fTitleBox.TextFrame2.WordWrap = True
    fTitleBox.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = tChartProps.Title.Font.RGB
    fTitleBox.Width = tChartProps.Title.Size.With
    fTitleBox.Height = tChartProps.Title.Size.Height
    fTitleBox.Top = tChartProps.Title.Size.Top
    fTitleBox.IncrementTop -3.75 'order matters

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


End Sub

Public Sub FormatSourceTextBox()

    Const cProc = "FormatSourceTextBox"

    On Error GoTo ErrorHandler

    Dim fSources As Shape

    On Error Resume Next
    Set fSources = oMyChart.Shapes("ChartFormatterSourceBox")
        
    If fSources Is Nothing Then
        Exit Sub
    End If

    Dim sFormula As String
    sFormula = fSources.OLEFormat.Object.Formula

    fSources.Line.Visible = msoFalse
    fSources.TextFrame2.TextRange.Font.Bold = tChartProps.SourceTextBox.Font.Bold
    fSources.TextFrame2.TextRange.Font.Size = tChartProps.SourceTextBox.Font.Size
    fSources.TextFrame2.TextRange.Font.Name = tChartProps.SourceTextBox.Font.Name
    fSources.TextFrame2.TextRange.ParagraphFormat.Alignment = tChartProps.SourceTextBox.TextAlignment

    'fSources.OLEFormat.Object.Formula = sFormula

    Exit Sub
ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


End Sub


Public Sub FormatChartTitle()

    Const cProc = "FormatChartTitle"

    On Error GoTo ErrorHandler

    If Not oMyChart.HasTitle Then
        Exit Sub
    End If

    Dim sFormula As String
    sFormula = oMyChart.ChartTitle.Formula

    oMyChart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = tChartProps.Title.Font.Size
    oMyChart.ChartTitle.Format.TextFrame2.TextRange.Font.Name = tChartProps.Title.Font.Name
    oMyChart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = tChartProps.Title.Font.Bold
    oMyChart.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = tChartProps.Title.Font.RGB

    oMyChart.ChartTitle.Formula = sFormula

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


End Sub


Public Sub FormatChartAxisTitle()

    Const cProc = "FormatChartAxisTitle"

    On Error GoTo ErrorHandler

    Dim ax As axis

    For Each ax In oMyChart.Axes

        ax.TickLabels.Font.Size = tChartProps.AxisTitle.Font.Size
        ax.TickLabels.Font.Name = tChartProps.AxisTitle.Font.Name
        ax.TickLabels.Font.Bold = tChartProps.AxisTitle.Font.Bold
        ax.TickLabels.Font.Color = tChartProps.AxisTitle.Font.RGB

        If ax.HasTitle = True Then

            Dim sFormula As String
            sFormula = ax.AxisTitle.Formula

            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Size = tChartProps.AxisTitle.Font.Size
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Name = tChartProps.AxisTitle.Font.Name
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Bold = tChartProps.AxisTitle.Font.Bold
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = tChartProps.AxisTitle.Font.RGB

            ax.AxisTitle.Formula = sFormula

        End If

    Next

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


End Sub

Public Sub FormatChartSeries()

    Const cProc = "FormatChartSeries"

    On Error GoTo ErrorHandler

    Dim i As Integer

    For i = 1 To oMyChart.SeriesCollection.count

        Dim serS As Series
       Set serS = oMyChart.SeriesCollection(i)
       
       serS.MarkerStyle = -4142

        If IsLineChart(serS.ChartType) Then

            If i <= UBound(tChartProps.SeriesColors) Then

                serS.Format.Line.Weight = tChartProps.SeriesWeight
                serS.Format.Line.ForeColor.RGB = tChartProps.SeriesColors(i - 1)
                serS.Format.Fill.ForeColor.RGB = tChartProps.SeriesColors(i - 1)

            End If

            serS.Format.Line.Weight = tChartProps.SeriesWeight

        ElseIf IsColumnBarAreaChart(serS.ChartType) Then

            serS.Format.Line.Weight = 1

            If i <= UBound(tChartProps.SeriesColors) Then

                serS.Format.Line.ForeColor.RGB = tChartProps.SeriesColors(i - 1)
                serS.Format.Fill.ForeColor.RGB = tChartProps.SeriesColors(i - 1)

            End If

            serS.Format.Line.ForeColor.RGB = serS.Fill.ForeColor.RGB

        End If


    Next

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub

Public Sub FormatChartAxes()

    Const cProc = "FormatChartAxes"

    On Error GoTo ErrorHandler

    Dim ax As axis

    For Each ax In oMyChart.Axes

        ax.TickLabels.Font.Color = RGB(0, 0, 0)

        If ax.AxisGroup = xlPrimary And ax.Type = xlValue Then
            ax.Format.Line.Visible = msoFalse
            ax.HasMajorGridlines = True
            ax.HasMinorGridlines = False
            ax.MajorGridlines.Format.Line.ForeColor.RGB = RGB(191, 191, 191)
            ax.MajorGridlines.Format.Line.Weight = 0.25
            ax.TickLabelPosition = xlTickLabelPositionNextToAxis
            ax.CrossesAt = -99999
            'ax.TickLabels.NumberFormat = "# ##0.0"

        End If

        If ax.AxisGroup = xlSecondary And ax.Type = xlValue Then
            ax.Format.Line.Visible = msoFalse
            ax.HasMajorGridlines = False
            ax.HasMinorGridlines = False
            ax.TickLabelPosition = xlTickLabelPositionNextToAxis
            'ax.TickLabels.NumberFormat = "# ##0.0"
        End If

        If ax.Type = xlSeriesAxis Then

            ax.HasMajorGridlines = False
            ax.HasMinorGridlines = False
            ax.Format.Line.Visible = msoTrue
            ax.Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            ax.Format.Line.Transparency = 0
            ax.Format.Line.Weight = 1.5
            ax.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
            ax.Format.Fill.Transparency = 0
            ax.Format.Fill.Solid
            ax.TickLabelPosition = xlLow
            'ax.TickLabelSpacing =

        End If

        If ax.Type = xlCategory Then

            ax.HasMajorGridlines = False
            ax.HasMinorGridlines = False
            ax.MajorTickMark = xlTickMarkInside
            ax.MinorTickMark = xlTickMarkNone
            ax.Format.Line.Visible = msoTrue
            ax.Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            ax.Format.Line.Transparency = 0
            ax.Format.Line.Weight = 1.5
            ax.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
            ax.Format.Fill.Transparency = 0
            ax.Format.Fill.Solid
            ax.TickLabelPosition = xlLow
            'ax.TickLabelSpacing = 1

        End If


    Next

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


End Sub


Public Sub FormatLegend()

    Const cProc = "FormatLegend"

    On Error GoTo ErrorHandler

    If oMyChart.HasLegend = False Then
        Exit Sub
    End If

    oMyChart.HasLegend = True
    oMyChart.Legend.Left = tChartProps.Legend.Size.Left
    oMyChart.Legend.Top = tChartProps.Legend.Size.Top
    oMyChart.Legend.Width = tChartProps.Legend.Size.With
    oMyChart.Legend.Height = tChartProps.Legend.Size.Height
    oMyChart.Legend.Format.TextFrame2.TextRange.Font.Name = tChartProps.Legend.Font.Name
    oMyChart.Legend.Format.TextFrame2.TextRange.Font.Size = tChartProps.Legend.Font.Size
    oMyChart.Legend.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = tChartProps.Legend.Font.RGB
    oMyChart.Legend.Format.Line.Visible = msoFalse

    Exit Sub
ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub

Public Sub FormatPlotArea()

    Const cProc = "FormatPlotArea"

    On Error GoTo ErrorHandler

    'VBA bug in With setting: https://www.oipapio.com/question-4167359

    With oMyChart.PlotArea
        .Select
        With Selection
            .Height = tChartProps.PlotArea.Size.Height
            .Top = tChartProps.PlotArea.Size.Top
            .Left = tChartProps.PlotArea.Size.Left
            .Width = tChartProps.PlotArea.Size.With
        End With
    End With


    Exit Sub
ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


End Sub


Public Sub Rescale()

    Const cProc = "Rescale"

    Dim i As Integer
    Dim tNewDateScale As MyNewDateScale
    Dim tMyAxis As clsAxes

    On Error GoTo ErrorHandler

    If oMyChart.HasAxis(xlValue, xlPrimary) = True Then

        Dim ax As axis
        Set ax = oMyChart.Axes(xlCategory, xlPrimary)
             
        If IsDateAxis(ax) Then

            For i = 1 To UBound(arrMySeries)

                If arrMySeries(i).GetSeriesAxisGroup = XlAxisGroup.xlPrimary And Not arrMySeries(i).GetSeriesAxis Is Nothing Then

                    Call arrMySeries(i).GetSeriesAxis.Rescale
                    tNewDateScale = arrMySeries(i).GetSeriesAxis.GetMyNewDateScale

                    If tNewDateScale.Rescaled = True Then

                        ax.MinimumScale = tNewDateScale.MinDate
                        ax.MaximumScale = tNewDateScale.MaxDate
                        ax.MajorUnit = tNewDateScale.MajorUnit
                        ax.MajorUnitScale = tNewDateScale.MajorUnitScale
                        'ax.BaseUnit = tNewDateScale.BaseUnit 'BUG: base unit shall remain unchanged
                        ax.TickLabels.NumberFormat = tNewDateScale.FormatCode

                    End If


                End If

            Next

        End If

    End If


    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


End Sub


