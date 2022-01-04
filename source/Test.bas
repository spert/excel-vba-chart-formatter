Option Explicit

Const cModule = "clsChart"

Private oOldChart As Chart

Private tChartProps As MyChart

Private shtNewWorkSheet As WorkSheet

Private oNewChart As ChartObject

Private sTitle() As String

Private arrMySeries() As clsSeries

Private arrMyAxes() As clsAxes

Private bMultipleAxesGroups As Boolean

Private bIsValidExcelChart As Boolean


Public Sub InitiateChartFormat(OldChart As Chart, ChartProps As MyChart)
  
  Set oOldChart = OldChart
  tChartProps = ChartProps
  bMultipleAxesGroups = False

End Sub

Public Property Let SetNewWorksheet(NewWorkSheet As clsWorkSheet)

        Set shtNewWorkSheet = NewWorkSheet.GetNewWorkSheet
        
End Property

Public Property Get GetIsVaidExcelChart() As Boolean

   Set GetIsVaidExcelChart = bIsValidExcelChart

End Property

Public Property Get GetChartTitle() As String

   Set GetChartTitle = sTitle

End Property


Public Sub CollectSeries()

   Const cProc = "CollectSeries"
   
   On Error GoTo ErrorHandler:

   Dim i As Integer, j As Integer, clsMySeries As clsSeries
        
   For i = 1 To oOldChart.SeriesCollection.count
         
          Dim serS As Series
          Set serS = oOldChart.SeriesCollection(i)
          
         Set clsMySeries = New clsSeries
         Call clsMySeries.Initiate(serS, oOldChart, tChartProps)
         
         If (Not arrMySeries) = -1 Then
            
            ReDim Preserve arrMySeries(1)
            Set arrMySeries(1) = clsMySeries
            
         Else
         
            ReDim Preserve arrMySeries(UBound(arrMySeries) + 1)
            Set arrMySeries(UBound(arrMySeries)) = clsMySeries
        
         End If
                  
         If clsMySeries.GetSeriesAxisGroup = xlSecondary Then
            
            bMultipleAxesGroups = True
         
         End If
                 
   Next
   
        Exit Sub
    
ErrorHandler:
    
   ErrorMod.ErrorMessage cProc, cModule
   

End Sub
 
Public Sub CollectAxes()

   Const cProc = "CollectAxes"
   
   Dim i As Integer, j As Integer
   Dim clsMyAxis As clsAxes
   
   On Error GoTo ErrorHandler:


  For i = 1 To UBound(arrMySeries)

   Dim arrCat(5) As Variant
   arrCat(1) = arrMySeries(i).GetMySources(2, 1)
   arrCat(2) = arrMySeries(i).GetMySources(2, 2)
   arrCat(3) = arrMySeries(i).GetMySources(2, 3)
   arrCat(4) = arrMySeries(i).GetMySources(2, 4)
   arrCat(5) = arrMySeries(i).GetMySources(2, 5)

    Set clsMyAxis = New clsAxes
    Call clsMyAxis.Initialize(arrCat, tChartProps)
    'Call clsMyAxis.ToString(i)

    Dim axisFound As Boolean
    axisFound = False

   If (Not arrMyAxes) = -1 Then

    'If arrMyAxes.co Is Nothing Then

      ReDim Preserve arrMyAxes(1)
      Set arrMyAxes(1) = clsMyAxis
      arrMySeries(i).SetSeriesAxis = clsMyAxis

    Else

      For j = 1 To UBound(arrMyAxes) ' - 1

            If clsMyAxis.IsSameAs(arrMyAxes(j)) Then

                  axisFound = True
                  arrMySeries(i).SetSeriesAxis = arrMyAxes(j)

            End If
      Next

      If axisFound = False Then

          ReDim Preserve arrMyAxes(UBound(arrMyAxes) + 1)
          Set arrMyAxes(UBound(arrMyAxes)) = clsMyAxis
          arrMySeries(i).SetSeriesAxis = clsMyAxis

      End If

    End If

  Next
  
          Exit Sub
    
ErrorHandler:
    
   ErrorMod.ErrorMessage cProc, cModule
  
  
End Sub

Public Sub PrintChartTitle()

   Const cProc = "PrintChartTitle"
     
   On Error GoTo ErrorHandler:

   Dim rngHeadingLink As Range
   Set rngHeadingLink = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(2, 5)

   Dim rngHeading As Range
   Set rngHeading = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(2, 1)
   rngHeadingLink.Formula = "=" & Chr(34) & tChartProps.Language_Descriptions(7) & Chr(34) & " & IF(ISERROR(CELL(" & Chr(34) & "Filename" & Chr(34) & ",B1)),RIGHT(CELL(" & Chr(34) & "Failinimi" & Chr(34) & ",B1),LEN(CELL(" & Chr(34) & "Failinimi" & Chr(34) & ",B1))-FIND(" & Chr(34) & "]" & Chr(34) & ",CELL(" & Chr(34) & "Failinimi" & Chr(34) & ",B1))),RIGHT(CELL(" & Chr(34) & "Filename" & Chr(34) & ",B1),LEN(CELL(" & Chr(34) & "Filename" & Chr(34) & ",B1))-FIND(" & Chr(34) & "]" & Chr(34) & ",CELL(" & Chr(34) & "Filename" & Chr(34) & ",B1))))&" & Chr(34) & ". " & Chr(34) & "&" & rngHeading.Address(, , , False, False)
   rngHeadingLink.Font.Bold = True

    If tChartProps.Title.Position = MyPosition.Top And tChartProps.Title.BoxEnabled = True Then
    
        oNewChart.Chart.HasTitle = False
    
        Dim fHeading As Shape
        Set fHeading = oNewChart.Chart.Shapes.AddTextbox(msoTextOrientationHorizontal, tChartProps.Title.Size.Left, tChartProps.Title.Size.Top, tChartProps.Title.Size.With, tChartProps.Title.Size.Height)  '(Left, Top, Width, Height)
        fHeading.OLEFormat.Object.Formula = rngHeadingLink.Address(, , , True)
        fHeading.Fill.ForeColor.RGB = tChartProps.Title.BackgroundRGB
        fHeading.Fill.Transparency = 0
        fHeading.Fill.Solid
        fHeading.Line.Visible = msoFalse
        fHeading.TextFrame2.VerticalAnchor = msoAnchorMiddle
        fHeading.TextFrame2.HorizontalAnchor = msoAnchorNone
        fHeading.TextFrame2.TextRange.Font.Bold = tChartProps.Title.Font.Bold
        fHeading.TextFrame2.TextRange.Font.Size = tChartProps.Title.Font.Size
        fHeading.TextFrame2.TextRange.Font.Name = tChartProps.Title.Font.Name
        fHeading.TextFrame2.MarginLeft = 5.6692913386
        fHeading.TextFrame2.MarginRight = 5.6692913386
        fHeading.TextFrame2.MarginTop = 2.8346456693
        fHeading.TextFrame2.MarginBottom = 2.8346456693
        fHeading.TextFrame2.WordWrap = True
        fHeading.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        fHeading.IncrementTop -3.75

   End If
    
    If tChartProps.Title.Position = MyPosition.Top And tChartProps.Title.BoxEnabled = False Then
    
        oNewChart.Chart.HasTitle = True
        
        oNewChart.Chart.SetElement (msoElementChartTitleAboveChart)
        oNewChart.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = tChartProps.Title.Font.Size
        oNewChart.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Name = tChartProps.Title.Font.Name
        oNewChart.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = tChartProps.Title.Font.Bold
        oNewChart.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
      
        oNewChart.Chart.ChartTitle.Formula = "=" & rngHeadingLink.Address(, , , True)

    End If
              
    If tChartProps.Title.Position = MyPosition.None Then
              
        oNewChart.Chart.HasTitle = False
              
    End If
              
  Exit Sub
    
ErrorHandler:
    
   ErrorMod.ErrorMessage cProc, cModule
  
  
End Sub

Public Sub PrintChartAxisTitle()

    Dim rngLeftAxisTitle As Range
    Set rngLeftAxisTitle = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(9, 1)
    
    Dim rngRightAxisTitle As Range
    Set rngRightAxisTitle = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(10, 1)

    Dim ax As axis
            
    For Each ax In oNewChart.Chart.Axes
                            
            ax.TickLabels.Font.Size = tChartProps.AxisTitle.Font.Size
            ax.TickLabels.Font.Name = tChartProps.AxisTitle.Font.Name
            ax.TickLabels.Font.Bold = tChartProps.AxisTitle.Font.Bold
            ax.TickLabels.Font.color = tChartProps.AxisTitle.Font.RGB
              
        If ax.HasTitle = True And ax.Type = xlValue And ax.AxisGroup = xlPrimary Then
        
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Size = tChartProps.AxisTitle.Font.Size
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Name = tChartProps.AxisTitle.Font.Name
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Bold = tChartProps.AxisTitle.Font.Bold
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = tChartProps.AxisTitle.Font.RGB
                    
            rngLeftAxisTitle = ax.AxisTitle.Text
            
            Call FormatRange(rngLeftAxisTitle)
            
            ax.AxisTitle.Formula = "=" & rngLeftAxisTitle.Address(, , , True)

        End If
  
        If ax.HasTitle = True And ax.Type = xlValue And ax.AxisGroup = xlSecondary Then
            
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Size = tChartProps.AxisTitle.Font.Size
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Name = tChartProps.AxisTitle.Font.Name
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Bold = tChartProps.AxisTitle.Font.Bold
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = tChartProps.AxisTitle.Font.RGB
            
            rngRightAxisTitle = ax.AxisTitle.Text
            
            Call FormatRange(rngRightAxisTitle)
            
            ax.AxisTitle.Formula = "=" & rngRightAxisTitle.Address(, , , True)
            
        End If
  
    Next
     
    
End Sub


Public Sub AssignNewRanges()

   Const cProc = "AssignNewRanges"
   
   Dim iColumn As Integer, i As Integer
   
   On Error GoTo ErrorHandler:

  '  ================ Assign New Range To Axis ==============
   iColumn = 0

   For i = 1 To UBound(arrMyAxes)

        arrMyAxes(i).SetColumn = iColumn
        arrMyAxes(i).SetNewWorksheet = shtNewWorkSheet
        Call arrMyAxes(i).SetNewCategoryRange
        Call arrMyAxes(i).PrintCategoryValues
                
        'Call arrMyAxes(i).ToString(i)
        
        iColumn = iColumn + 1

   Next
      

  ' ================ Assign New Range To Series ==============
   For i = 1 To UBound(arrMySeries)

    arrMySeries(i).SetColumn = iColumn
    arrMySeries(i).SetNewWorksheet = shtNewWorkSheet
    Call arrMySeries(i).SetNewCategoryRange
    Call arrMySeries(i).SetNewValueRange
    Call arrMySeries(i).SetNewNameRange
    Call arrMySeries(i).SetNewNameLinkRange
    
    Call arrMySeries(i).PrintSeriesName
    Call arrMySeries(i).PrintSeriesNameLink
    Call arrMySeries(i).PrintSeriesValues
    Call arrMySeries(i).PrintSeriesNumber(i)
    Call arrMySeries(i).PrintSeriesScale(bMultipleAxesGroups)

    iColumn = iColumn + 1
    
   Next
      
      Exit Sub
    
ErrorHandler:
    
   ErrorMod.ErrorMessage cProc, cModule
   
   
 End Sub

Public Sub MapNewSeries()
    
   Const cProc = "MapNewSeries"
   
   On Error GoTo ErrorHandler:
    
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

Public Sub ApplySeriesFormat()

   Const cProc = "ApplySeriesFormat"
   
   On Error GoTo ErrorHandler:

   Dim i As Integer
    
   For i = 1 To oNewChart.Chart.SeriesCollection.count

       Dim serS As Series
       Set serS = oNewChart.Chart.SeriesCollection(i)
       
       If serS.ChartType = xlLine Then
            serS.Format.Line.Weight = tChartProps.SeriesWeight
       Else
           serS.Format.Line.Weight = 1
      End If
       
        If i <= UBound(tChartProps.SeriesColor) Then
            serS.Format.Line.ForeColor.RGB = tChartProps.SeriesColor(i - 1)
            serS.Format.Fill.ForeColor.RGB = tChartProps.SeriesColor(i - 1)
        Else
            serS.Format.Line.ForeColor.RGB = serS.Fill.ForeColor.RGB
        End If
             
   Next

    Exit Sub
    
ErrorHandler:
    
   ErrorMod.ErrorMessage cProc, cModule

End Sub


Public Sub CopyOldChartToNewWorksheet()

Const cProc = "CopyOldChartToNewWorksheet"

On Error GoTo ErrorHandler:

Set oNewChart = shtNewWorkSheet.ChartObjects.Add(tChartProps.Left, tChartProps.Top, tChartProps.With, tChartProps.Height) '(Left, Top, Width, Height)

oOldChart.ChartArea.Copy
oNewChart.Activate
ActiveChart.Paste

 Exit Sub
 
ErrorHandler:

 ErrorMod.ErrorMessage cProc, cModule


End Sub

Public Sub ChartAxes()

Const cProc = "ChartAxes"

On Error GoTo ErrorHandler:

Dim ax As axis

For Each ax In oNewChart.Chart.Axes

    ax.TickLabels.Font.color = RGB(0, 0, 0)

    If ax.AxisGroup = xlPrimary And ax.Type = xlValue Then
        ax.Format.Line.Visible = msoFalse
        ax.HasMajorGridlines = True
        ax.HasMinorGridlines = False
        ax.MajorGridlines.Format.Line.ForeColor.RGB = RGB(191, 191, 191)
        ax.MajorGridlines.Format.Line.Weight = 0.25
        ax.TickLabelPosition = xlTickLabelPositionNextToAxis
        'ax.TickLabels.NumberFormat = "# ##0.0"
    
    End If
    
    If ax.AxisGroup = xlSecondary And ax.Type = xlValue Then
        ax.Format.Line.Visible = msoFalse
        ax.HasMajorGridlines = False
        ax.HasMinorGridlines = False
        ax.TickLabelPosition = xlTickLabelPositionNextToAxis
        ax.TickLabels.NumberFormat = "# ##0.0"
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
        ax.TickLabelSpacing = 1
    
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
        ax.TickLabelSpacing = 1
        
    End If
    

Next

 Exit Sub
 
ErrorHandler:

 ErrorMod.ErrorMessage cProc, cModule


End Sub


Public Sub PrintLegend()

    Const cProc = "PrintLegend"

    On Error GoTo ErrorHandler:

    oNewChart.Chart.HasLegend = True
    oNewChart.Chart.Legend.Left = tChartProps.Legend.Size.Left
    oNewChart.Chart.Legend.Top = tChartProps.Legend.Size.Top
    oNewChart.Chart.Legend.Width = tChartProps.Legend.Size.With
    oNewChart.Chart.Legend.Height = tChartProps.Legend.Size.Height
    oNewChart.Chart.Legend.Format.TextFrame2.TextRange.Font.Name = tChartProps.Legend.Font.Name
    oNewChart.Chart.Legend.Format.TextFrame2.TextRange.Font.Size = tChartProps.Legend.Font.Size
    oNewChart.Chart.Legend.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = tChartProps.Legend.Font.RGB
    oNewChart.Chart.Legend.Format.Line.Visible = msoFalse

Exit Sub
ErrorHandler:

 ErrorMod.ErrorMessage cProc, cModule

End Sub

Public Sub PrintPlotArea()

    shtNewWorkSheet.Shapes(oNewChart.Name).Line.Visible = msoFalse
    
    'VBA bug in With setting: https://www.oipapio.com/question-4167359
    
    With oNewChart.Chart.PlotArea
       .Select
        With Selection
            .Height = tChartProps.PlotArea.Size.Height
            .Top = tChartProps.PlotArea.Size.Top
            .Left = tChartProps.PlotArea.Size.Left
            .Width = tChartProps.PlotArea.Size.With
            
        End With
    End With
    
End Sub


Public Sub PrintSourceTextBox()
    
    Const cProc = "PrintSourceTextBox"
    
    On Error GoTo ErrorHandler:
    
    Dim fSources As Shape
    Set fSources = oNewChart.Chart.Shapes.AddTextbox(msoTextOrientationHorizontal, tChartProps.SourceTextBox.Size.Left, tChartProps.SourceTextBox.Size.Top, tChartProps.SourceTextBox.Size.With, tChartProps.SourceTextBox.Size.Height) '(Left, Top, Width, Height)
    
    Dim rngSourceLink As Range
    Set rngSourceLink = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(3, 5)
    
    Dim rngSource As Range
    Set rngSource = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1)).Offset(3, 1)
    rngSourceLink.Formula = "=IF(ISNUMBER(SEARCH(" & Chr(34) & "," & Chr(34) & "," & rngSource.Address(, , , False, False) & "))," & Chr(34) & tChartProps.Language_Descriptions(6) & Chr(34) & "," & Chr(34) & tChartProps.Language_Descriptions(5) & Chr(34) & ")&" & rngSource.Address(, , , False, False)
    
    fSources.OLEFormat.Object.Formula = rngSourceLink.Address(, , , True)
    fSources.Line.Visible = msoFalse
    fSources.TextFrame2.TextRange.Font.Bold = tChartProps.SourceTextBox.Font.Bold
    fSources.TextFrame2.TextRange.Font.Size = tChartProps.SourceTextBox.Font.Size
    fSources.TextFrame2.TextRange.Font.Name = tChartProps.SourceTextBox.Font.Name
    fSources.TextFrame2.TextRange.ParagraphFormat.Alignment = tChartProps.SourceTextBox.TextAlignment
    
    Exit Sub
ErrorHandler:
    
     ErrorMod.ErrorMessage cProc, cModule
    

End Sub

Public Sub ValidateExcelChart()
 
   Const cProc = "ValidateExcelChart"
   
   Dim i As Integer
   Dim sFormula As String
   
   bIsValidExcelChart = False
   
   On Error GoTo ErrorHandler:
 
  If oOldChart.SeriesCollection.count = 0 Then
        
      Err.Raise Number:=990, Description:="The chart can't be empty"
    
  End If

  For i = 1 To oOldChart.SeriesCollection.count
         
          Dim serS As Series
          Set serS = oOldChart.SeriesCollection(i)
          
           On Error GoTo InvalidSeries
           
            sFormula = serS.Formula

  Next
  
  bIsValidExcelChart = True
  
  Exit Sub

InvalidSeries:
        
     Err.Raise Number:=991, Description:="One of series has invalid formula"
        
ErrorHandler:

 ErrorMod.ErrorMessage cProc, cModule

 End Sub


Public Sub ToString()

Dim i As Integer

For i = 1 To UBound(arrMySeries)

    Call arrMySeries(i).ToString

Next

End Sub


Public Sub Rescale()

    Const cProc = "Rescale"
    
    Dim i As Integer
    Dim tNewDateScale As MyNewDateScale
    Dim tMyAxis As clsAxes
    
    On Error GoTo ErrorHandler:
        
        
    If oNewChart.Chart.HasAxis(xlValue, xlPrimary) = True Then
    
        Dim ax As axis
        Set ax = oNewChart.Chart.Axes(xlCategory, xlPrimary)
             
        For i = 1 To UBound(arrMySeries)
        
           If arrMySeries(i).GetSeriesAxisGroup = XlAxisGroup.xlPrimary Then
        
               Call arrMySeries(i).GetSeriesAxis.Rescale
               tNewDateScale = arrMySeries(i).GetSeriesAxis.GetMyNewDateScale

               If tNewDateScale.Rescaled = True Then

                   ax.MinimumScale = tNewDateScale.MinDate
                   ax.MaximumScale = tNewDateScale.MaxDate
                   ax.MajorUnit = tNewDateScale.MajorUnit
                   ax.MajorUnitScale = tNewDateScale.MajorUnitScale
                   ax.BaseUnit = tNewDateScale.BaseUnit
                   ax.TickLabels.NumberFormat = tNewDateScale.FormatCode

               End If

            
           End If
                                        
        Next
    
    End If
       
    
     Exit Sub
     
ErrorHandler:
    
     ErrorMod.ErrorMessage cProc, cModule
    

End Sub


'If oNewChart.HasAxis(xlValue, xlPrimary) = True Then
'
'    Dim ax As axis
'    Set ax = oNewChart.HasAxis(xlValue, xlPrimary)
'
'    If ax.CategoryType = xlTimeScale Then
'
'      Dim tMyNewAxis As MyNewDateScale
'      tMyNewAxis = CalculateNewDateSpan(tMyOldAxis)
'
'    End If
'
'End If



