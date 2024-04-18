
Option Explicit On

Const cModule = "clsOldChart"

Private oOldChart As Chart

Private tChartProps As MyChart

Private strTitle As String

Private arrMySeries() As clsSeries

Private arrMyAxes() As clsAxes

Private bMultipleAxesGroups As Boolean

Private bIsValidExcelChart As Boolean

Public Sub Initiate(OldChart As Chart)
  
  Set oOldChart = OldChart
  bMultipleAxesGroups = False

End Sub

Public Property Get GetMyAxes() As clsAxes()

 GetMyAxes = arrMyAxes

End Property

Public Property Get GetMySeries() As clsSeries()

   GetMySeries = arrMySeries

End Property

Public Property Get GetOldChart() As Chart

    Set GetOldChart = oOldChart

End Property

Public Property Get IsMultipleAxisGroups() As Boolean

   IsMultipleAxisGroups = bMultipleAxesGroups

End Property


Public Sub ValidateExcelChart()

    Const cProc = "ValidateExcelChart"

    Dim i As Integer
    Dim sFormula As String

    bIsValidExcelChart = False

    On Error GoTo ErrorHandler

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


Public Sub CollectSeries()

    Const cProc = "CollectSeries"

    On Error GoTo ErrorHandler

    Dim i As Integer, j As Integer, clsMySeries As clsSeries

    For i = 1 To oOldChart.SeriesCollection.count

        Dim serS As Series
          Set serS = oOldChart.SeriesCollection(i)
          
         Set clsMySeries = New clsSeries
         Call clsMySeries.Initiate(serS, oOldChart)

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

    On Error GoTo ErrorHandler


    For i = 1 To UBound(arrMySeries)

        Dim arrCat(5) As Variant
        arrCat(1) = arrMySeries(i).GetMySources(2, 1)
        arrCat(2) = arrMySeries(i).GetMySources(2, 2)
        arrCat(3) = arrMySeries(i).GetMySources(2, 3)
        arrCat(4) = arrMySeries(i).GetMySources(2, 4)
        arrCat(5) = arrMySeries(i).GetMySources(2, 5)

    Set clsMyAxis = New clsAxes
    Call clsMyAxis.Initialize(arrCat)

        If clsMyAxis.IsEmptyAxis() Then

            GoTo NextIteration

        End If

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

NextIteration:


    Next

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub














