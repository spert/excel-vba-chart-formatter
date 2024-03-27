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

Public Function GetFormatProperties() As MyChart()

    Dim aChartProps() As MyChart

    '====== Worksheet
    Dim tChartProps As MyChart
    tChartProps = WorksheetPoperties(tChartProps)

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


    '====== Rescale X Axis (NOT IMPLEMENTED)
    '    If aOptions(5) Then
    '
    '        tChartProps = RescaleXAxis(tChartProps)
    '
    '    End If


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

    GetFormatProperties = aChartProps

End Function

Private Function WorksheetPoperties(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    Dim wb As Workbook

    Dim sht As Worksheet

    Dim i As Integer

    i = 0

    If wb Is Nothing Then Set wb = ActiveWorkBook
    
    Do
        i = i + 1
        
        Set sht = Nothing
    
        On Error Resume Next

        Dim s As String
        s = CStr(i)
        
        Set sht = wb.Sheets(s)
                
        On Error GoTo 0

    Loop Until sht Is Nothing

    tProps.SheetName = CStr(i)

    WorksheetPoperties = tProps

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
        tProps.SourceTextBox.Font.Name = "Calibri"
        tProps.SourceTextBox.Font.Size = 10
        tProps.SourceTextBox.Font.Bold = False
        tProps.SourceTextBox.Font.RGB = RGB(0, 0, 0)
    Else
        tProps.SourceTextBox.Position = MyPosition.None
    End If

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
        tProps.SourceTextBox.Font.Name = "Calibri"
        tProps.SourceTextBox.Font.Size = 14
        tProps.SourceTextBox.Font.Bold = False
        tProps.SourceTextBox.Font.RGB = RGB(0, 0, 0)
    Else
        tProps.SourceTextBox.Position = MyPosition.None
    End If

    LargeChartSourceBox = tProps

End Function


Private Function SmallChartSize(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    tProps.Height = 249.249842519685
    tProps.With = 215.499921259843
    tProps.ColumnWith = 8.43

    Dim aVert(11, 2) As Variant
    aVert(0, 0) = 0  'marging to separate chart border and title (% of chart total height)
    aVert(1, 0) = 10 'title relative height (% of chart total height)
    aVert(2, 0) = 2  'marging to separate title and legend (% of chart total height)
    aVert(3, 0) = 10 'legend height (% of chart total height)
    aVert(4, 0) = 0  'marging to separate legend and axis title on bottom (% of chart total height)
    aVert(5, 0) = 5  'axis title on top plot area(% of chart total height)
    aVert(6, 0) = 59 'plot area relative height (% of chart total height)
    aVert(7, 0) = 5  'axis title on bottom of plot area (% of chart total height)
    aVert(8, 0) = 0  'marging to separate plot area and source box (% of chart total height)
    aVert(9, 0) = 7  'source box height (% of chart total height)
    aVert(10, 0) = 3 'marging to separate source box chart border (% of chart total height)

    tProps.Title.Position = MyPosition.Top
    tProps.SourceTextBox.TextAlignment = msoAlignLeft

    'title missing
    If aOptions(4) = False Then
        aVert(0, 0) = 0
        aVert(1, 0) = 0
        tProps.Title.Position = MyPosition.None
    End If

    'title not in text box
    If aOptions(4) = True And aOptions(5) = False Then
        aVert(0, 0) = 2
        aVert(1, 0) = 5
    End If

    'legend
    If oOldChart.HasLegend = False Then
        aVert(2, 0) = 0
        aVert(3, 0) = 0
    End If

    'source
    If aOptions(6) = False Then
        aVert(8, 0) = 0
        aVert(9, 0) = 0
    End If

    aVert(5, 0) = aVert(5, 0) * tProps.NumberOfAxesTitles(2)
    aVert(7, 0) = aVert(7, 0) * tProps.NumberOfAxesTitles(3)

    aVert(0, 1) = aVert(0, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(1, 1) = aVert(1, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(2, 1) = aVert(2, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(3, 1) = aVert(3, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(4, 1) = aVert(4, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(5, 1) = aVert(5, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(6, 1) = aVert(6, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(7, 1) = aVert(7, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(8, 1) = aVert(8, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(9, 1) = aVert(9, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(10, 1) = aVert(10, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100

    aVert(1, 2) = aVert(0, 1) 'title
    aVert(3, 2) = aVert(0, 1) + aVert(1, 1) + aVert(2, 1)  'legend
    aVert(6, 2) = aVert(0, 1) + aVert(1, 1) + aVert(2, 1) + aVert(3, 1) + aVert(4, 1) + aVert(5, 1)  'plot area
    aVert(9, 2) = aVert(0, 1) + aVert(1, 1) + aVert(2, 1) + aVert(3, 1) + aVert(4, 1) + aVert(5, 1) + aVert(6, 1) + aVert(7, 1) + aVert(8, 1) 'source

    ' ----------------- HEIGHT --------------------
    tProps.Title.Size.Height = tProps.Height * aVert(1, 1) / 100
    tProps.Legend.Size.Height = tProps.Height * aVert(3, 1) / 100
    tProps.PlotArea.Size.Height = tProps.Height * aVert(6, 1) / 100
    tProps.SourceTextBox.Size.Height = tProps.Height * aVert(9, 1) / 100

    ' ----------------- TOP --------------------
    tProps.Title.Size.Top = tProps.Height * aVert(1, 2) / 100
    tProps.Legend.Size.Top = tProps.Height * aVert(3, 2) / 100
    tProps.PlotArea.Size.Top = tProps.Height * aVert(6, 2) / 100
    tProps.SourceTextBox.Size.Top = tProps.Height * aVert(9, 2) / 100

    Dim aHoriz(14, 2) As Variant
    aHoriz(0, 0) = 0  'marging to separate chart title form left border (% of chart total with)
    aHoriz(1, 0) = 100  'title relative with (% of chart total with)
    aHoriz(2, 0) = 0  'marging to separate title from right border (% of chart total with)

    aHoriz(3, 0) = 1   'marging to separate legend from left border (% of chart total with)
    aHoriz(4, 0) = 98  'legend with (% of chart total with)
    aHoriz(5, 0) = 8  'marging to separate legend from right border (% of chart total with)

    aHoriz(6, 0) = 1   'marging to separate axis title on left border(% of chart total with)
    aHoriz(7, 0) = 6   'marging to separate axis title from plot area (% of chart total with)
    aHoriz(8, 0) = 91  'plot area with (% of chart total with)
    aHoriz(9, 0) = 6   'marging to separate plot area from axis title (% of chart total with)
    aHoriz(10, 0) = 5   'marging to separate axis title on left border(% of chart total with)

    aHoriz(11, 0) = 1  'marging to separate legend from left border (% of chart total with)
    aHoriz(12, 0) = 98 'legend with (% of chart total with)
    aHoriz(13, 0) = 1  'marging to separate legend from right border (% of chart total with)

    aHoriz(7, 0) = aHoriz(7, 0) * tProps.NumberOfAxesTitles(0)
    aHoriz(9, 0) = aHoriz(9, 0) * tProps.NumberOfAxesTitles(1)

    aHoriz(0, 1) = aHoriz(0, 0) / (aHoriz(0, 0) + aHoriz(1, 0) + aHoriz(2, 0)) * 100
    aHoriz(1, 1) = aHoriz(1, 0) / (aHoriz(0, 0) + aHoriz(1, 0) + aHoriz(2, 0)) * 100
    aHoriz(2, 1) = aHoriz(2, 0) / (aHoriz(0, 0) + aHoriz(1, 0) + aHoriz(2, 0)) * 100

    aHoriz(3, 1) = aHoriz(3, 0) / (aHoriz(3, 0) + aHoriz(4, 0) + aHoriz(5, 0)) * 100
    aHoriz(4, 1) = aHoriz(4, 0) / (aHoriz(3, 0) + aHoriz(4, 0) + aHoriz(5, 0)) * 100
    aHoriz(5, 1) = aHoriz(5, 0) / (aHoriz(3, 0) + aHoriz(4, 0) + aHoriz(5, 0)) * 100

    aHoriz(6, 1) = aHoriz(6, 0) / (aHoriz(6, 0) + aHoriz(7, 0) + aHoriz(8, 0) + aHoriz(9, 0) + aHoriz(10, 0)) * 100
    aHoriz(7, 1) = aHoriz(7, 0) / (aHoriz(6, 0) + aHoriz(7, 0) + aHoriz(8, 0) + aHoriz(9, 0) + aHoriz(10, 0)) * 100
    aHoriz(8, 1) = aHoriz(8, 0) / (aHoriz(6, 0) + aHoriz(7, 0) + aHoriz(8, 0) + aHoriz(9, 0) + aHoriz(10, 0)) * 100
    aHoriz(9, 1) = aHoriz(9, 0) / (aHoriz(6, 0) + aHoriz(7, 0) + aHoriz(8, 0) + aHoriz(9, 0) + aHoriz(10, 0)) * 100
    aHoriz(10, 1) = aHoriz(10, 0) / (aHoriz(6, 0) + aHoriz(7, 0) + aHoriz(8, 0) + aHoriz(9, 0) + aHoriz(10, 0)) * 100

    aHoriz(11, 1) = aHoriz(11, 0) / (aHoriz(11, 0) + aHoriz(12, 0) + aHoriz(13, 0)) * 100
    aHoriz(12, 1) = aHoriz(12, 0) / (aHoriz(12, 0) + aHoriz(12, 0) + aHoriz(13, 0)) * 100
    aHoriz(13, 1) = aHoriz(13, 0) / (aHoriz(13, 0) + aHoriz(12, 0) + aHoriz(13, 0)) * 100

    ' ----------------- LEFT --------------------
    'title without box
    If aOptions(4) = True And aOptions(5) = False And oOldChart.HasTitle = True Then
        tProps.Title.Size.Left = oOldChart.ChartTitle.Left

    End If

    'title into box
    If aOptions(4) = True And aOptions(5) = True Then
        tProps.Title.Size.Left = tProps.With * aHoriz(0, 1) / 100
        tProps.Title.Size.With = tProps.With * aHoriz(1, 1) / 100
    End If

    tProps.Legend.Size.Left = tProps.With * aHoriz(3, 1) / 100
    tProps.Legend.Size.With = tProps.With * aHoriz(4, 1) / 100

    tProps.PlotArea.Size.Left = tProps.With * (aHoriz(6, 1) + aHoriz(7, 1)) / 100
    tProps.PlotArea.Size.With = tProps.With * aHoriz(8, 1) / 100

    tProps.SourceTextBox.Size.Left = tProps.With * aHoriz(11, 0) / 100
    tProps.SourceTextBox.Size.With = tProps.With * aHoriz(12, 0) / 100


    '    ' ----------------- HEIGHT --------------------
    '    'initial height
    '    tProps.Title.Size.Height = tProps.Height * 0.144
    '    tProps.Legend.Size.Height = tProps.Height * 0.18
    '    tProps.SourceTextBox.Size.Height = tProps.Height * 0.144
    '    tProps.Title.Position = MyPosition.Top
    '
    '    'title
    '    If aOptions(4) = False Then
    '        tProps.Title.Size.Height = 0
    '        tProps.Title.Position = MyPosition.None
    '    End If
    '
    '    'source
    '    If aOptions(6) = False Then
    '        tProps.SourceTextBox.Size.Height = 0
    '        tProps.SourceTextBox.TextAlignment = msoAlignRight
    '    End If
    '
    '    'legend
    '    If oOldChart.HasLegend = False Then
    '       tProps.Legend.Size.Height = 0
    '    End If
    '
    '    'plot area
    '    tProps.PlotArea.Size.Height = tProps.Height - tProps.Title.Size.Height - tProps.Legend.Size.Height - tProps.Height * tProps.NumberOfAxesTitles(2) * 0.07 - tProps.Height * tProps.NumberOfAxesTitles(3) * 0.07 - tProps.Height * 0.08
    '
    '     ' ----------------- TOP --------------------
    '     tProps.Title.Size.Top = 0
    '     tProps.PlotArea.Size.Top = tProps.Height * 0.32 + tProps.Height * tProps.NumberOfAxesTitles(3) * 0.07
    '     tProps.Legend.Size.Top = tProps.Height * 0.144
    '     tProps.SourceTextBox.Size.Top = tProps.Height * 0.95
    '
    '    ' ----------------- LEFT --------------------
    '    tProps.Title.Size.Left = 0
    '    tProps.PlotArea.Size.Left = tProps.With * 0.01 + tProps.With * tProps.NumberOfAxesTitles(0) * 0.03
    '    tProps.Legend.Size.Left = tProps.With * 0.01
    '    tProps.SourceTextBox.Size.Left = tProps.With * 0.01
    '
    '    ' ----------------- WITH --------------------
    '     tProps.Title.Size.With = tProps.With
    '     tProps.PlotArea.Size.With = tProps.With - tProps.With * tProps.NumberOfAxesTitles(0) * 0.03 - tProps.With * tProps.NumberOfAxesTitles(1) * 0.03 - tProps.With * 0.03
    '     tProps.Legend.Size.With = tProps.With - tProps.With * 0.03
    '     tProps.SourceTextBox.Size.With = tProps.With - tProps.With * 0.03


    '
    '    'TitleTop = True
    '    If aOptions(4) = True Then
    '        tProps.Title.Position = MyPosition.Top
    '        tProps.Title.Size.Top = 0
    '        tProps.Title.Size.Left = 0
    '        tProps.Title.Size.Height = tProps.Height * 0.144
    '        tProps.Title.Size.With = tProps.With
    '
    '        tProps.Legend.Size.Top = tProps.Height * 0.14
    '        tProps.Legend.Size.Height = tProps.Height * 0.18
    '        tProps.Legend.Size.With = tProps.With
    '    Else
    '        tProps.Title.Position = MyPosition.None
    '        tProps.Legend.Size.Top = 0
    '        tProps.Legend.Size.Height = tProps.Height * 0.18
    '        tProps.Legend.Size.With = tProps.With
    '    End If
    '
    '    If aOptions(6) = True Then
    '        tProps.SourceTextBox.Size.Top = tProps.Height - tProps.Height * 0.1
    '        tProps.SourceTextBox.Size.Height = tProps.Height * 0.1
    '        tProps.SourceTextBox.Size.With = tProps.With
    '        tProps.SourceTextBox.TextAlignment = msoAlignLeft
    '    Else
    '        tProps.SourceTextBox.Size.Top = 0
    '        tProps.SourceTextBox.Size.Height = 0
    '        tProps.SourceTextBox.Size.With = 0
    '        tProps.SourceTextBox.TextAlignment = msoAlignLeft
    '    End If
    '
    '    tProps.PlotArea.Size.Top = tProps.Title.Size.Height + tProps.Legend.Size.Height
    '    tProps.PlotArea.Size.Height = tProps.Height - tProps.Title.Size.Height - tProps.Legend.Size.Height - tProps.SourceTextBox.Size.Height
    '
    '    ' Vertial axes titles
    '    If tProps.NumberOfAxesTitles(0) = 0 And tProps.NumberOfAxesTitles(1) = 0 Then
    '
    '        tProps.PlotArea.Size.Left = 0
    '        tProps.PlotArea.Size.With = tProps.With - 3
    '
    '    ElseIf tProps.NumberOfAxesTitles(0) = 1 And tProps.NumberOfAxesTitles(1) = 0 Then
    '
    '        tProps.PlotArea.Size.Left = 12
    '        tProps.PlotArea.Size.With = tProps.With - 12
    '
    '    ElseIf tProps.NumberOfAxesTitles(0) = 1 And tProps.NumberOfAxesTitles(1) = 1 Then
    '
    '        tProps.PlotArea.Size.Left = 12
    '        tProps.PlotArea.Size.With = tProps.With - 30
    '
    '    End If
    '
    '    ' Horizontal axes titles
    '    If tProps.NumberOfAxesTitles(2) = 0 And tProps.NumberOfAxesTitles(3) = 0 Then
    '
    '        tProps.PlotArea.Size.Top = tProps.PlotArea.Size.Top
    '        tProps.PlotArea.Size.Height = tProps.PlotArea.Size.Height - 12
    '
    '    ElseIf tProps.NumberOfAxesTitles(2) = 1 And tProps.NumberOfAxesTitles(3) = 0 Then
    '
    '        tProps.PlotArea.Size.Top = tProps.PlotArea.Size.Top - 12
    '        tProps.PlotArea.Size.Height = tProps.PlotArea.Size.Height - 12
    '
    '    End If

    SmallChartSize = tProps

End Function

Private Function LargeChartSize(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props

    tProps.Height = 370
    tProps.With = 830
    tProps.ColumnWith = 120

    Dim aVert(11, 2) As Variant
    aVert(0, 0) = 0  'marging to separate chart border and title (% of chart total height)
    aVert(1, 0) = 7 'title relative height (% of chart total height)
    aVert(2, 0) = 0  'marging to separate title and axis title on top (% of chart total height)
    aVert(3, 0) = 5  'axis title on top plot area(% of chart total height)
    aVert(4, 0) = 59 'plot area relative height (% of chart total height)
    aVert(5, 0) = 5  'axis title on bottom of plot area (% of chart total height)
    aVert(6, 0) = 0  'marging to separate axis title on bottom and legend (% of chart total height)
    aVert(7, 0) = 10 'legend height (% of chart total height)
    aVert(8, 0) = 0  'marging to separate legend and source box (% of chart total height)
    aVert(9, 0) = 7  'source box height (% of chart total height)
    aVert(10, 0) = 3 'marging to separate source box chart border (% of chart total height)

    tProps.Title.Position = MyPosition.Top
    tProps.SourceTextBox.TextAlignment = msoAlignLeft

    'title missing
    If aOptions(4) = False Then
        aVert(0, 0) = 0
        aVert(1, 0) = 0
        tProps.Title.Position = MyPosition.None
    End If

    'title not in text box
    If aOptions(4) = True And aOptions(5) = False Then
        aVert(0, 0) = 2
        aVert(1, 0) = 5
    End If

    'legend
    If oOldChart.HasLegend = False Then
        aVert(7, 0) = 0
        aVert(6, 0) = 0
    End If

    'source
    If aOptions(6) = False Then
        aVert(8, 0) = 0
        aVert(9, 0) = 0
    End If

    aVert(3, 0) = aVert(3, 0) * tProps.NumberOfAxesTitles(2)
    aVert(5, 0) = aVert(5, 0) * tProps.NumberOfAxesTitles(3)

    aVert(0, 1) = aVert(0, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(1, 1) = aVert(1, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(2, 1) = aVert(2, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(3, 1) = aVert(3, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(4, 1) = aVert(4, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(5, 1) = aVert(5, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(6, 1) = aVert(6, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(7, 1) = aVert(7, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(8, 1) = aVert(8, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(9, 1) = aVert(9, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100
    aVert(10, 1) = aVert(10, 0) / (aVert(0, 0) + aVert(1, 0) + aVert(2, 0) + aVert(3, 0) + aVert(4, 0) + aVert(5, 0) + aVert(6, 0) + aVert(7, 0) + aVert(8, 0) + aVert(9, 0) + aVert(10, 0)) * 100

    aVert(1, 2) = aVert(0, 1) 'title
    aVert(4, 2) = aVert(0, 1) + aVert(1, 1) + aVert(2, 1) + aVert(3, 0) 'plot area
    aVert(7, 2) = aVert(0, 1) + aVert(1, 1) + aVert(2, 1) + aVert(3, 1) + aVert(4, 1) + aVert(5, 1) + aVert(6, 1)  'legend
    aVert(9, 2) = aVert(0, 1) + aVert(1, 1) + aVert(2, 1) + aVert(3, 1) + aVert(4, 1) + aVert(5, 1) + aVert(6, 1) + aVert(7, 1) + aVert(8, 1) 'source

    ' ----------------- HEIGHT --------------------
    tProps.Title.Size.Height = tProps.Height * aVert(1, 1) / 100
    tProps.PlotArea.Size.Height = tProps.Height * aVert(4, 1) / 100
    tProps.Legend.Size.Height = tProps.Height * aVert(7, 1) / 100
    tProps.SourceTextBox.Size.Height = tProps.Height * aVert(9, 1) / 100

    ' ----------------- TOP --------------------
    tProps.Title.Size.Top = tProps.Height * aVert(1, 2) / 100
    tProps.PlotArea.Size.Top = tProps.Height * aVert(4, 2) / 100
    tProps.Legend.Size.Top = tProps.Height * aVert(7, 2) / 100
    tProps.SourceTextBox.Size.Top = tProps.Height * aVert(9, 2) / 100

    Dim aHoriz(14, 2) As Variant
    aHoriz(0, 0) = 0  'marging to separate chart title form left border (% of chart total with)
    aHoriz(1, 0) = 100  'title relative with (% of chart total with)
    aHoriz(2, 0) = 0  'marging to separate title from right border (% of chart total with)

    aHoriz(3, 0) = 1   'marging to separate axis title on left border(% of chart total with)
    aHoriz(4, 0) = 3   'marging to separate axis title from plot area (% of chart total with)
    aHoriz(5, 0) = 87  'plot area with (% of chart total with)
    aHoriz(6, 0) = 0  'marging to separate plot area from axis title (% of chart total with)
    aHoriz(7, 0) = 3   'marging to separate axis title on left border(% of chart total with)

    aHoriz(8, 0) = 1   'marging to separate legend from left border (% of chart total with)
    aHoriz(9, 0) = 98  'legend with (% of chart total with)
    aHoriz(10, 0) = 1  'marging to separate legend from right border (% of chart total with)

    aHoriz(11, 0) = 1  'marging to separate legend from left border (% of chart total with)
    aHoriz(12, 0) = 98 'legend with (% of chart total with)
    aHoriz(13, 0) = 1  'marging to separate legend from right border (% of chart total with)

    aHoriz(4, 0) = aHoriz(4, 0) * tProps.NumberOfAxesTitles(0)
    aHoriz(6, 0) = aHoriz(6, 0) * tProps.NumberOfAxesTitles(1)

    aHoriz(0, 1) = aHoriz(0, 0) / (aHoriz(0, 0) + aHoriz(1, 0) + aHoriz(2, 0)) * 100
    aHoriz(1, 1) = aHoriz(1, 0) / (aHoriz(0, 0) + aHoriz(1, 0) + aHoriz(2, 0)) * 100
    aHoriz(2, 1) = aHoriz(2, 0) / (aHoriz(0, 0) + aHoriz(1, 0) + aHoriz(2, 0)) * 100

    aHoriz(3, 1) = aHoriz(3, 0) / (aHoriz(3, 0) + aHoriz(4, 0) + aHoriz(5, 0) + aHoriz(6, 0) + aHoriz(7, 0)) * 100
    aHoriz(4, 1) = aHoriz(4, 0) / (aHoriz(3, 0) + aHoriz(4, 0) + aHoriz(5, 0) + aHoriz(6, 0) + aHoriz(7, 0)) * 100
    aHoriz(5, 1) = aHoriz(5, 0) / (aHoriz(3, 0) + aHoriz(4, 0) + aHoriz(5, 0) + aHoriz(6, 0) + aHoriz(7, 0)) * 100
    aHoriz(6, 1) = aHoriz(6, 0) / (aHoriz(3, 0) + aHoriz(4, 0) + aHoriz(5, 0) + aHoriz(6, 0) + aHoriz(7, 0)) * 100
    aHoriz(7, 1) = aHoriz(7, 0) / (aHoriz(3, 0) + aHoriz(4, 0) + aHoriz(5, 0) + aHoriz(6, 0) + aHoriz(7, 0)) * 100

    aHoriz(8, 1) = aHoriz(8, 0) / (aHoriz(8, 0) + aHoriz(9, 0) + aHoriz(10, 0)) * 100
    aHoriz(9, 1) = aHoriz(9, 0) / (aHoriz(8, 0) + aHoriz(9, 0) + aHoriz(10, 0)) * 100
    aHoriz(10, 1) = aHoriz(10, 0) / (aHoriz(8, 0) + aHoriz(9, 0) + aHoriz(10, 0)) * 100

    aHoriz(11, 1) = aHoriz(11, 0) / (aHoriz(11, 0) + aHoriz(12, 0) + aHoriz(13, 0)) * 100
    aHoriz(12, 1) = aHoriz(12, 0) / (aHoriz(12, 0) + aHoriz(12, 0) + aHoriz(13, 0)) * 100
    aHoriz(13, 1) = aHoriz(13, 0) / (aHoriz(13, 0) + aHoriz(12, 0) + aHoriz(13, 0)) * 100

    ' ----------------- LEFT --------------------
    'title without box
    If aOptions(4) = True And aOptions(5) = False And oOldChart.HasTitle = True Then
        tProps.Title.Size.Left = oOldChart.ChartTitle.Left

    End If

    'title into box
    If aOptions(4) = True And aOptions(5) = True Then
        tProps.Title.Size.Left = tProps.With * aHoriz(0, 1) / 100
        tProps.Title.Size.With = tProps.With * aHoriz(1, 1) / 100
    End If

    tProps.PlotArea.Size.Left = tProps.With * (aHoriz(3, 1) + aHoriz(4, 1)) / 100
    tProps.PlotArea.Size.With = tProps.With * aHoriz(5, 1) / 100

    tProps.Legend.Size.Left = tProps.With * aHoriz(8, 1) / 100
    tProps.Legend.Size.With = tProps.With * aHoriz(9, 1) / 100

    tProps.SourceTextBox.Size.Left = tProps.With * aHoriz(11, 0) / 100
    tProps.SourceTextBox.Size.With = tProps.With * aHoriz(12, 0) / 100

    LargeChartSize = tProps

End Function

Public Function GetSeriesColors() As Variant

    GetSeriesColors = aSeriesColors

End Function


'Private Function NumberOfYAxisTitles() As Integer
'
'    Dim Counter As Integer
'    Counter = 0
'
'    Dim ax As axis
'
'    For Each ax In oOldChart.Axes
'
'        If ax.Type = xlValue And ax.HasTitle = True Then
'
'            Counter = Counter + 1
'
'        End If
'
'    Next ax
'
'    NumberOfYAxisTitles = Counter
'
'End Function





