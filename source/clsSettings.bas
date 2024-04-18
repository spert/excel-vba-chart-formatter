Option Explicit On

Dim aOptions As Variant

Dim wbActiveWorkBook As Workbook

Dim oOldChart As Chart

Dim aChartProps() As Variant

Dim aLanguage_Descriptions(1) As Variant

Dim aLanguage_ScaleValidation(1) As Variant

Dim aSeriesColors As Variant


Public Sub InitiateSettings(AWorkBook As Workbook, OldChart As Chart, Options As Variant)

    aOptions = Options
   
   Set wbActiveWorkBook = AWorkBook
   
   Set oOldChart = OldChart
 
   aLanguage_Descriptions(0) = Array("English", "Title:", "Chart number:", "Series name:", "Series scale:", "Source: ", "Sources: ", "Chart ", "<title>", "<source>", "Name and scale:", "Left Axis title :", "Right axis title:")
    aLanguage_Descriptions(1) = Array("Eesti keeles", "Pealkiri:", "Joonise number:", "Aegrea nimi:", "Aegrea telg:", "Allikas: ", "Allikad: ", "Joonis ", "<pealkiri>", "<allikas>", "Nimi ja telg:", "Vasaku telje pealkiri:", "Parema telje pealkiri:")
    aLanguage_ScaleValidation(0) = Array("(left scale)", "(right scale)")
    aLanguage_ScaleValidation(1) = Array("(vasak telg)", "(parem telg)")
    'aSeriesColors(0) = Array(RGB(223, 158, 48), RGB(190, 82, 3), RGB(44, 68, 94), RGB(46, 108, 129), RGB(100, 178, 199))
    aSeriesColors = Array(RGB(255, 215, 57), RGB(133, 138, 140), RGB(0, 121, 188), RGB(138, 204, 97), RGB(255, 141, 67), RGB(244, 145, 212), RGB(233, 158, 48), RGB(190, 82, 33), RGB(100, 178, 199))

End Sub

Public Function GetFormatChartProps() As MyChart

    '====== Worksheet
    Dim tChartProps As MyChart
    '    tChartProps = WorksheetPoperties(tChartProps)

    '======= Number of Axis Titles (NB! order matters)
    tChartProps = NumberOfAxesTitles(tChartProps)

    '======= Title text
    tChartProps = TitleText(tChartProps)

    ' ====== Series Colors
    tChartProps = SeriesColorsAndWeight(tChartProps)

    '====== Small chart
    If aOptions(0) Then

        tChartProps = SmallAxisTitle(SmallChartSourceBox(SmallChartTitle(SmallChartLegend(SmallChartSize(tChartProps)))))

    End If

    '====== Slide chart
    If aOptions(1) Then

        tChartProps = LargeAxisTitle(LargeChartSourceBox(LargeChartTitle(LargeChartLegend(LargeChartSize(tChartProps)))))

    End If

    tChartProps = Language1(tChartProps)

    GetFormatChartProps = tChartProps

End Function


Public Function GetCopyChartProps() As MyChart()

    Dim aChartProps() As MyChart

    '====== Worksheet
    Dim tChartProps As MyChart
    '    tChartProps = WorksheetPoperties(tChartProps)

    '======= Copy as links
    tChartProps.CopyValuesAsLinks = aOptions(7)

    '======= Number of Axis Titles (NB! order matters)
    tChartProps = NumberOfAxesTitles(tChartProps)

    '======= Title text
    tChartProps = TitleText(tChartProps)

    ' ====== Series Colors
    tChartProps = SeriesColorsAndWeight(tChartProps)

    '====== Small chart
    If aOptions(0) Then

        tChartProps = SmallAxisTitle(SmallChartSourceBox(SmallChartTitle(SmallChartLegend(SmallChartSize(tChartProps)))))

    End If

    '====== Slide chart
    If aOptions(1) Then

        tChartProps = LargeAxisTitle(LargeChartSourceBox(LargeChartTitle(LargeChartLegend(LargeChartSize(tChartProps)))))

    End If

    '====== 1 Language
    If aOptions(2) Then

        ReDim aChartProps(0)
        aChartProps(0) = Language1(tChartProps)

    End If

    '====== 2 Languages
    If aOptions(3) Then

        ReDim Preserve aChartProps(1)
        aChartProps(0) = Language1(tChartProps)
        aChartProps(1) = Language2(tChartProps, aChartProps(0))

    End If

    GetCopyChartProps = aChartProps

End Function

Private Function SeriesColorsAndWeight(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    tProps.SeriesColors = aSeriesColors
    tProps.SeriesWeight = 2

    SeriesColorsAndWeight = tProps

End Function


Private Function Language1(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    tProps.Language_Descriptions = aLanguage_Descriptions(0)
    tProps.Language_ScaleValidation = aLanguage_ScaleValidation(0)
    tProps.FirstCellOfOutput = [{1;7}]
    tProps.SeriesDataOffset = [{10;8}]
    tProps.Top = 50
    tProps.Left = 25
    tProps.PrintAxesAndValues = True

    Language1 = tProps

End Function

Private Function Language2(props As MyChart, props1 As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    tProps.Language_Descriptions = aLanguage_Descriptions(1)
    tProps.Language_ScaleValidation = aLanguage_ScaleValidation(1)
    tProps.FirstCellOfOutput = [{9;7}]
    tProps.Top = props1.Top + props1.Height + 20
    tProps.Left = props1.Left
    tProps.SeriesDataOffset = [{18;8}]
    tProps.PrintAxesAndValues = False
    props1.SeriesDataOffset = [{18;8}]
              
    Language2 = tProps

End Function

Private Function SmallAxisTitle(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    tProps.AxisTitle.Font.Name = "Calibri"
    tProps.AxisTitle.Font.Size = 10
    tProps.AxisTitle.Font.Bold = False
    tProps.AxisTitle.Font.RGB = RGB(0, 0, 0)

    SmallAxisTitle = tProps

End Function

Private Function LargeAxisTitle(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    tProps.AxisTitle.Font.Name = "Calibri"
    tProps.AxisTitle.Font.Size = 14
    tProps.AxisTitle.Font.Bold = False
    tProps.AxisTitle.Font.RGB = RGB(0, 0, 0)

    LargeAxisTitle = tProps

End Function


Private Function SmallChartLegend(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    tProps.Legend.Font.Name = "Calibri"
    tProps.Legend.Font.Size = 10
    tProps.Legend.Font.Bold = False
    tProps.Legend.Font.RGB = RGB(0, 0, 0)


    SmallChartLegend = tProps

End Function

Private Function LargeChartLegend(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    tProps.Legend.Font.Name = "Calibri"
    tProps.Legend.Font.Size = 14
    tProps.Legend.Font.Bold = False
    tProps.Legend.Font.RGB = RGB(0, 0, 0)


    LargeChartLegend = tProps

End Function


Private Function SmallChartTitle(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    tProps.Title.Font.Name = "Calibri"
    tProps.Title.Font.Size = 12
    tProps.Title.Font.Bold = True

    'TitleTextBox
    If aOptions(5) = True Then
        tProps.Title.BoxEnabled = True
        tProps.Title.BackgroundRGB = RGB(44, 68, 94)
        tProps.Title.Font.RGB = RGB(255, 255, 255)
    Else
        tProps.Title.Font.RGB = RGB(0, 0, 0)
    End If

    SmallChartTitle = tProps

End Function

Private Function LargeChartTitle(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    tProps.Title.Font.Name = "Calibri"
    tProps.Title.Font.Size = 20
    tProps.Title.Font.Bold = True

    'TitleTextBox
    If aOptions(5) = True Then
        tProps.Title.Font.Size = 16
        tProps.Title.BoxEnabled = True
        tProps.Title.BackgroundRGB = RGB(44, 68, 94)
        tProps.Title.Font.RGB = RGB(255, 255, 255)
    Else
        tProps.Title.Font.RGB = RGB(0, 0, 0)
    End If

    LargeChartTitle = tProps

End Function

Private Function NumberOfAxesTitles(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    Dim aNbrAxisTitles(3) As Variant
    aNbrAxisTitles(0) = 0
    aNbrAxisTitles(1) = 0
    aNbrAxisTitles(2) = 0
    aNbrAxisTitles(3) = 0

    Dim ax As Excel.axis

    For Each ax In oOldChart.Axes

        If ax.AxisGroup = xlPrimary And ax.Type = xlCategory And ax.HasTitle Then
            aNbrAxisTitles(2) = aNbrAxisTitles(2) + 1
        End If

        If ax.AxisGroup = xlSecondary And ax.Type = xlCategory And ax.HasTitle Then
            aNbrAxisTitles(3) = aNbrAxisTitles(3) + 1
        End If

        If ax.AxisGroup = xlPrimary And ax.Type = xlValue And ax.HasTitle Then
            aNbrAxisTitles(0) = aNbrAxisTitles(0) + 1
        End If

        If ax.AxisGroup = xlSecondary And ax.Type = xlValue And ax.HasTitle Then
            aNbrAxisTitles(1) = aNbrAxisTitles(1) + 1
        End If

    Next

    tProps.NumberOfAxesTitles = aNbrAxisTitles

    NumberOfAxesTitles = tProps

End Function


Private Function TitleText(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    If oOldChart.HasTitle Then

        If StrComp(oOldChart.ChartTitle.Text, "Chart Title") = 1 Then

            tProps.Title.Text = oOldChart.ChartTitle.Text

        End If

    End If

    Dim sh As Shape
    On Error Resume Next
    Set sh = oOldChart.Shapes("ChartFormatterTitleBox")
            
    If Not sh Is Nothing Then
        tProps.Title.Text = sh.TextFrame2.TextRange.Text
        sh.Delete
    End If

    If tProps.Title.Text = "" Then

        tProps.Title.Text = "<title>"

    End If

    TitleText = tProps

End Function


Private Function SmallChartSourceBox(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    tProps.SourceTextBox.Text = "<source>"
    tProps.SourceTextBox.TextAlignment = msoAlignRight

    Dim sh As Shape
    On Error Resume Next
    Set sh = oOldChart.Shapes("ChartFormatterSourceBox")
            
    If Not sh Is Nothing Then
        tProps.SourceTextBox.Text = sh.TextFrame2.TextRange.Text
    End If

    If aOptions(6) = True Then
        tProps.SourceTextBox.Position = MyPosition.Bottom
    Else
        tProps.SourceTextBox.Position = MyPosition.None
    End If

    tProps.SourceTextBox.Font.Name = "Calibri"
    tProps.SourceTextBox.Font.Size = 10
    tProps.SourceTextBox.Font.Bold = False
    tProps.SourceTextBox.Font.RGB = RGB(0, 0, 0)

    SmallChartSourceBox = tProps

End Function

Private Function LargeChartSourceBox(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    tProps.SourceTextBox.Text = "<source>"
    tProps.SourceTextBox.TextAlignment = msoAlignLeft

    Dim sh As Shape
    On Error Resume Next
    Set sh = oOldChart.Shapes("ChartFormatterSourceBox")
            
    If Not sh Is Nothing Then
        tProps.SourceTextBox.Text = sh.TextFrame2.TextRange.Text
    End If

    If aOptions(6) = True Then
        tProps.SourceTextBox.Position = MyPosition.Bottom
    Else
        tProps.SourceTextBox.Position = MyPosition.None
    End If

    tProps.SourceTextBox.Font.Name = "Calibri"
    tProps.SourceTextBox.Font.Size = 14
    tProps.SourceTextBox.Font.Bold = False
    tProps.SourceTextBox.Font.RGB = RGB(0, 0, 0)

    LargeChartSourceBox = tProps

End Function


Private Function SmallChartSize(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    tProps.Height = 249.249842519685
    tProps.With = 215.499921259843
    tProps.ColumnWith = 8.43

    tProps.Title.Position = MyPosition.Top
    tProps.SourceTextBox.TextAlignment = msoAlignLeft

    tProps.PlotArea.Size.Top = 8
    tProps.PlotArea.Size.Height = 233

    'title missing
    If aOptions(4) = False Then
        tProps.Title.Position = MyPosition.None
        tProps.Title.Size.Height = 0

        'legend
        If oOldChart.HasLegend = True Then
            tProps.Legend.Size.Top = 2
            tProps.Legend.Size.Height = 50
            tProps.PlotArea.Size.Top = 50
            tProps.PlotArea.Size.Height = tProps.PlotArea.Size.Height - 45
        Else
            tProps.PlotArea.Size.Top = 10
            tProps.PlotArea.Size.Height = tProps.PlotArea.Size.Height - 10
        End If

    End If

    'title present but not in text box
    If aOptions(4) = True And aOptions(5) = False Then

        tProps.Title.Position = MyPosition.Top
        tProps.Title.Size.Top = 2
        tProps.Title.Size.Height = 20

        'legend
        If oOldChart.HasLegend = True Then
            tProps.Legend.Size.Top = 25
            tProps.Legend.Size.Height = 40
            tProps.PlotArea.Size.Top = 60
            tProps.PlotArea.Size.Height = tProps.PlotArea.Size.Height - 50
        Else
            tProps.PlotArea.Size.Top = 37
            tProps.PlotArea.Size.Height = tProps.PlotArea.Size.Height - 30
        End If

    End If

    'title present and in text box
    If aOptions(4) = True And aOptions(5) = True Then

        tProps.Title.Position = MyPosition.Top
        tProps.Title.Size.Top = 0
        tProps.Title.Size.Height = 35
        tProps.Title.Size.With = tProps.With

        'legend
        If oOldChart.HasLegend = True Then

            tProps.Legend.Size.Top = 35
            tProps.Legend.Size.Height = 40
            tProps.PlotArea.Size.Top = 70
            tProps.PlotArea.Size.Height = tProps.PlotArea.Size.Height - 60

        Else
            tProps.PlotArea.Size.Top = 37
            tProps.PlotArea.Size.Height = tProps.PlotArea.Size.Height - 30
        End If
    End If

    'source box
    If aOptions(6) = True Then
        tProps.SourceTextBox.Size.Top = 350
        tProps.SourceTextBox.Size.Height = 20
        tProps.Legend.Size.Height = tProps.Legend.Size.Height - 10
        tProps.PlotArea.Size.Height = tProps.PlotArea.Size.Height - 15
    End If

    ' ==================== Vertical alignment =======================
    tProps.PlotArea.Size.Left = 3
    tProps.PlotArea.Size.With = 203
    tProps.Legend.Size.Left = 3
    tProps.Legend.Size.With = 203
    tProps.SourceTextBox.Size.Left = 3
    tProps.SourceTextBox.Size.With = 203

    If tProps.NumberOfAxesTitles(0) > 0 And tProps.NumberOfAxesTitles(1) = 0 Then
        tProps.PlotArea.Size.Left = 10
        tProps.PlotArea.Size.With = tProps.PlotArea.Size.With - 5
    End If

    If tProps.NumberOfAxesTitles(0) = 0 And tProps.NumberOfAxesTitles(1) > 0 Then
        tProps.PlotArea.Size.Left = 3
        tProps.PlotArea.Size.With = tProps.PlotArea.Size.With - 10
    End If

    If tProps.NumberOfAxesTitles(0) > 0 And tProps.NumberOfAxesTitles(1) > 0 Then
        tProps.PlotArea.Size.Left = 10
        tProps.PlotArea.Size.With = tProps.PlotArea.Size.With - 20
    End If


    SmallChartSize = tProps

End Function

Private Function LargeChartSize(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    tProps.Height = 370
    tProps.With = 830
    tProps.ColumnWith = 120

    tProps.Title.Position = MyPosition.Top
    tProps.SourceTextBox.TextAlignment = msoAlignLeft

    tProps.PlotArea.Size.Top = 10
    tProps.PlotArea.Size.Height = 350

    ' ==================== Horizontal alignment =======================
    'title missing
    If aOptions(4) = False Then
        tProps.Title.Position = MyPosition.None
        tProps.Title.Size.Height = 0
    End If

    'title present but not in text box
    If aOptions(4) = True And aOptions(5) = False Then
        tProps.Title.Position = MyPosition.Top
        tProps.Title.Size.Top = 10
        tProps.Title.Size.Height = 70
        tProps.PlotArea.Size.Top = 40
        tProps.PlotArea.Size.Height = tProps.PlotArea.Size.Height - 40
    End If

    'title present and in text box
    If aOptions(4) = True And aOptions(5) = True Then
        tProps.Title.Position = MyPosition.Top
        tProps.Title.Size.Top = 0
        tProps.Title.Size.Height = 50
        tProps.Title.Size.Left = 0
        tProps.Title.Size.With = tProps.With
        tProps.PlotArea.Size.Top = 65
        tProps.PlotArea.Size.Height = tProps.PlotArea.Size.Height - 55
    End If

    'legend
    If oOldChart.HasLegend = True Then
        tProps.Legend.Size.Top = 290
        tProps.Legend.Size.Height = 70
        tProps.PlotArea.Size.Height = tProps.PlotArea.Size.Height - 60
    End If

    'source box
    If aOptions(6) = True Then
        tProps.SourceTextBox.Size.Top = 350
        tProps.SourceTextBox.Size.Height = 20
        tProps.Legend.Size.Height = tProps.Legend.Size.Height - 10
        tProps.PlotArea.Size.Height = tProps.PlotArea.Size.Height - 10
    End If

    ' ==================== Vertical alignment =======================
    tProps.PlotArea.Size.Left = 5
    tProps.PlotArea.Size.With = 810
    tProps.Legend.Size.Left = 5
    tProps.Legend.Size.With = 810
    tProps.SourceTextBox.Size.Left = 5
    tProps.SourceTextBox.Size.With = 810

    If tProps.NumberOfAxesTitles(0) > 0 And tProps.NumberOfAxesTitles(1) = 0 Then
        tProps.PlotArea.Size.Left = 20
        tProps.PlotArea.Size.With = tProps.PlotArea.Size.With - 20
    End If

    If tProps.NumberOfAxesTitles(0) = 0 And tProps.NumberOfAxesTitles(1) > 0 Then
        tProps.PlotArea.Size.Left = 5
        tProps.PlotArea.Size.With = tProps.PlotArea.Size.With - 20
    End If

    If tProps.NumberOfAxesTitles(0) > 0 And tProps.NumberOfAxesTitles(1) > 0 Then
        tProps.PlotArea.Size.Left = 25
        tProps.PlotArea.Size.With = tProps.PlotArea.Size.With - 40
    End If

    LargeChartSize = tProps

End Function

Public Function GetSeriesColors() As Variant

    GetSeriesColors = aSeriesColors

End Function




