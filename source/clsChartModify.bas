

Option Explicit On

Const cModule = "clsChartModify"

Private oMyChart As Chart

Private arrMySeries() As clsSeries

Private arrMyAxes() As clsAxes

Private tChartProps As MyChart


Public Sub InitiateFromOld(OldChart As clsChartOld, ChartProps As MyChart)

  Set oMyChart = OldChart.GetOldChart
  arrMyAxes() = OldChart.GetMyAxes
    arrMySeries() = OldChart.GetMySeries
    tChartProps = ChartProps

End Sub

Public Sub InitiateFromNew(NewChart As clsChartNew, ChartProps As MyChart)

  Set oMyChart = NewChart.GetNewChart
  arrMySeries() = NewChart.GetMySeries
    arrMyAxes() = NewChart.GetMyAxes
    tChartProps = ChartProps

End Sub


Public Sub ModifyChartTitle()

    Const cProc = "ModifyChartTitle"

    On Error GoTo ErrorHandler

    Dim sh As Shape
    For Each sh In oMyChart.Shapes
        If sh.Name = "ChartFormatterTitleBox" Then
            sh.Delete
        End If
    Next sh

    If oMyChart.HasTitle = True Then
        oMyChart.HasTitle = False
    End If

    If tChartProps.Title.Position = MyPosition.Top And tChartProps.Title.BoxEnabled = False Then

        oMyChart.HasTitle = True
        oMyChart.ChartTitle.Text = tChartProps.Title.Text
        'oMyChart.ChartTitle.Top = tChartProps.Title.Size.Top

    End If

    If tChartProps.Title.Position = MyPosition.Top And tChartProps.Title.BoxEnabled = True Then

        Dim fHeading As Shape
            Set fHeading = oMyChart.Shapes.AddTextbox(msoTextOrientationHorizontal, tChartProps.Title.Size.Left, tChartProps.Title.Size.Top, tChartProps.Title.Size.With, tChartProps.Title.Size.Height)  '(Left, Top, Width, Height)
            fHeading.Name = "ChartFormatterTitleBox"
        fHeading.TextFrame2.TextRange.Text = tChartProps.Title.Text

    End If


    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


End Sub

Public Sub ModifySourceBox()

    Const cProc = "ModifySourceBox"

    On Error GoTo ErrorHandler

    Dim sh As Shape
    For Each sh In oMyChart.Shapes
        If sh.Name = "ChartFormatterSourceBox" Then
            sh.Delete
        End If
    Next sh

    If tChartProps.SourceTextBox.Position = MyPosition.Bottom Then

        Dim fSource As Shape
            Set fSource = oMyChart.Shapes.AddTextbox(msoTextOrientationHorizontal, tChartProps.SourceTextBox.Size.Left, tChartProps.SourceTextBox.Size.Top, tChartProps.SourceTextBox.Size.With, tChartProps.SourceTextBox.Size.Height)  '(Left, Top, Width, Height)
            fSource.Name = "ChartFormatterSourceBox"
        fSource.TextFrame2.TextRange.Text = tChartProps.SourceTextBox.Text

    End If

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub


Public Sub RescaleXAxis()

    Const cProc = "RescaleXAxis"

    Dim i As Integer
    Dim tNewDateScale As MyNewDateScale
    Dim tMyAxis As clsAxes

    On Error GoTo ErrorHandler

    If tChartProps.ChartNeedsRescaling And oMyChart.HasAxis(xlValue, xlPrimary) = True Then

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
