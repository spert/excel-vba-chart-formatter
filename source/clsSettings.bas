Option Explicit

Dim aOptions As Variant

Dim wbActiveWorkBook As Workbook

Dim oOldChart As Chart

Dim aChartProps() As Variant

Dim aLanguage_Descriptions(1) As Variant

Dim aLanguage_ScaleValidation(1) As Variant
   
Dim aSeriesColors(1) As Variant

   
Public Sub InitiateSettings(AWorkBook As Workbook, OldChart As Chart, Options As Variant)

   aOptions = Options
   
   Set wbActiveWorkBook = AWorkBook
   
   Set oOldChart = OldChart
  
 
   aLanguage_Descriptions(0) = Array("English", "Title:", "", "Series name:", "Series scale:", "Source: ", "Sources: ", "Chart ", "<title>", "<source>", "Name and scale:", "Left Axis title :", "Right axis title:")
   aLanguage_Descriptions(1) = Array("Eesti keeles", "Pealkiri:", "", "Aegrea nimi:", "Aegrea telg:", "Allikas: ", "Allikad: ", "Joonis ", "<pealkiri>", "<allikas>", "Nimi ja telg:", "Vasaku telje pealkiri:", "Parema telje pealkiri:")
   aLanguage_ScaleValidation(0) = Array("(left scale)", "(right scale)")
   aLanguage_ScaleValidation(1) = Array("(vasak telg)", "(parem telg)")
   aSeriesColors(0) = Array(RGB(223, 158, 48), RGB(190, 82, 3), RGB(44, 68, 94), RGB(46, 108, 129), RGB(100, 178, 199))
                 
End Sub
   
Public Function GetFormatProperties() As MyChart()

    Dim aChartProps() As MyChart
        
    '====== Worksheet
    Dim tChartProps As MyChart
    tChartProps = WorksheetPoperties(tChartProps)

    '====== Small chart
    If aOptions(0) Then
    
        tChartProps = SmallAxisTitle(SmallChartSourceBox(SmallChartTitle(SmallChartLegend(SmallChartSize(tChartProps)))))
        
    End If
    
    '====== Slide chart
    If aOptions(1) Then
    
        tChartProps = LargeAxisTitle(LargeChartSourceBox(LargeChartTitle(LargeChartLegend(LargeChartSize(tChartProps)))))
    
    End If
    
    ' ====== Series Colors
    tChartProps = SeriesColorsAndWeight(tChartProps)
    
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
    
    Dim sht As WorkSheet
    
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

    tProps.SeriesColor = aSeriesColors(0)
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
        tProps.Title.Font.RGB = msoThemeColorBackground1
    Else
        tProps.Title.Font.RGB = RGB(0, 0, 0)
    End If

    If oOldChart.HasTitle Then
        
       If StrComp(oOldChart.ChartTitle.Text, "Chart Title") = 1 Then
        
            tProps.Title.Text = oOldChart.ChartTitle.Text
        
        End If
        
    Else
        
        tProps.Title.Text = ""
        
    End If

    SmallChartTitle = tProps

End Function

Private Function LargeChartTitle(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props
    
    tProps.Title.Font.Name = "Calibri"
    tProps.Title.Font.Size = 14
    tProps.Title.Font.Bold = True
          
    'TitleTextBox
    If aOptions(5) = True Then
        tProps.Title.BoxEnabled = True
        tProps.Title.BackgroundRGB = RGB(44, 68, 94)
        tProps.Title.Font.RGB = msoThemeColorBackground1
    Else
        tProps.Title.Font.RGB = RGB(0, 0, 0)
    End If

    If oOldChart.HasTitle Then
        
       If StrComp(oOldChart.ChartTitle.Text, "Chart Title") = 1 Then
        
            tProps.Title.Text = oOldChart.ChartTitle.Text
        
        End If
        
    Else
        
        tProps.Title.Text = ""
        
    End If

    LargeChartTitle = tProps

End Function

Private Function SmallChartSourceBox(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props
          
            
    tProps.SourceTextBox.Font.Name = "Calibri"
    tProps.SourceTextBox.Font.Size = 10
    tProps.SourceTextBox.Font.Bold = False
    tProps.SourceTextBox.Font.RGB = RGB(0, 0, 0)
      
    SmallChartSourceBox = tProps


End Function

Private Function LargeChartSourceBox(props As MyChart) As MyChart

    Dim tProps As MyChart
    tProps = props
          
            
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
    
    'TitleTop = True
    If aOptions(4) = True Then
        tProps.Title.Position = MyPosition.Top
        tProps.Title.Size.Top = 0
        tProps.Title.Size.Left = 0
        tProps.Title.Size.Height = tProps.Height * 0.144
        tProps.Title.Size.With = tProps.With
          
        tProps.Legend.Size.Top = tProps.Height * 0.14
        tProps.Legend.Size.Height = tProps.Height * 0.18
        tProps.Legend.Size.With = tProps.With
    Else
        tProps.Title.Position = MyPosition.None
        tProps.Legend.Size.Top = 0
        tProps.Legend.Size.Height = tProps.Height * 0.18
        tProps.Legend.Size.With = tProps.With
    End If
    
    tProps.SourceTextBox.Size.Top = tProps.Height - tProps.Height * 0.1
    tProps.SourceTextBox.Size.Height = tProps.Height * 0.1
    tProps.SourceTextBox.Size.With = tProps.With
    tProps.SourceTextBox.TextAlignment = msoAlignLeft
              
    tProps.PlotArea.Size.Top = tProps.Title.Size.Height + tProps.Legend.Size.Height
    tProps.PlotArea.Size.Height = tProps.Height - tProps.Title.Size.Height - tProps.Legend.Size.Height - tProps.SourceTextBox.Size.Height
       
    If NumberOfYAxisTitles = 0 Then
    
        tProps.PlotArea.Size.Left = 0
        tProps.PlotArea.Size.With = tProps.With
        
    ElseIf NumberOfYAxisTitles = 1 Then
    
        tProps.PlotArea.Size.Left = 12
        tProps.PlotArea.Size.With = tProps.With - 12
        
    Else
    
        tProps.PlotArea.Size.Left = 12
        tProps.PlotArea.Size.With = tProps.With - 30
        
    End If
    
      
    SmallChartSize = tProps

End Function

Private Function LargeChartSize(props As MyChart) As MyChart
     
    Dim tProps As MyChart
    tProps = props
     
    tProps.Height = 370
    tProps.With = 830
    tProps.ColumnWith = 120

    'TitleTop = True
    If aOptions(4) = True Then
        
        tProps.Title.Position = MyPosition.Top
        tProps.Title.Size.Top = 0
        tProps.Title.Size.Left = 0
        tProps.Title.Size.Height = tProps.Height * 0.1
        tProps.Title.Size.With = tProps.With
        
        tProps.Legend.Size.Top = tProps.Height - tProps.Height * 0.1 - tProps.Height * 0.07
        tProps.Legend.Size.Height = tProps.Height * 0.1
        tProps.Legend.Size.With = tProps.With
                  
    Else
    
        tProps.Title.Position = MyPosition.None
        tProps.Legend.Size.Top = tProps.Height - tProps.Height * 0.07 - tProps.Height * 0.07
        tProps.Legend.Size.Height = tProps.Height * 0.07
        tProps.Legend.Size.With = tProps.With
    
    End If
    
    tProps.SourceTextBox.Size.Top = tProps.Height - tProps.Height * 0.07
    tProps.SourceTextBox.Size.Height = tProps.Height * 0.07
    tProps.SourceTextBox.Size.With = tProps.With
    tProps.SourceTextBox.TextAlignment = msoAlignRight
              
    tProps.PlotArea.Size.Top = tProps.Title.Size.Height
    tProps.PlotArea.Size.Height = tProps.Height - tProps.Title.Size.Height - tProps.Legend.Size.Height - tProps.SourceTextBox.Size.Height - 10
       
    If NumberOfYAxisTitles = 0 Then
    
        tProps.PlotArea.Size.Left = 0
        tProps.PlotArea.Size.With = tProps.With - 10
        
    ElseIf NumberOfYAxisTitles = 1 Then
    
        tProps.PlotArea.Size.Left = 17
        tProps.PlotArea.Size.With = tProps.With - 17
        
    Else
    
        tProps.PlotArea.Size.Left = 17
        tProps.PlotArea.Size.With = tProps.With - 38
        
    End If
  
    LargeChartSize = tProps

End Function

Private Function NumberOfYAxisTitles() As Integer

    Dim Counter As Integer
    Counter = 0
    
    Dim ax As axis
    
    For Each ax In oOldChart.Axes
    
        If ax.Type = xlValue And ax.HasTitle = True Then
                   
            Counter = Counter + 1
        
        End If
        
    Next ax
    
    NumberOfYAxisTitles = Counter

End Function

