

Option Explicit On

Const cModule = "clsChartModify"

Private oMyChart As Chart

Private tChartProps As MyChart


Public Sub Initiate(MyChart As Chart, ChartProps As MyChart)
  
  Set oMyChart = MyChart
  tChartProps = ChartProps

End Sub


Public Sub ModifyChartTitle()

    Const cProc = "ModifyChartTitle"

    ' On Error GoTo ErrorHandler:

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
