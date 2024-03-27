Option Explicit On

Const cModule = "clsWorkSheet"

Dim wbOld As Workbook
Private tChartProps As MyChart

Private rOut As Range

Private shtNewWorkSheet As Worksheet

Public Sub Initiate(Workbook As Workbook, ChartProps As MyChart)

  Set wbOld = Workbook
  tChartProps = ChartProps

End Sub

Property Get GetNewWorkSheet() As Worksheet

    Set GetNewWorkSheet = shtNewWorkSheet

End Property


Public Sub AddNewChartToWorkbook()

    Const cProc = "AddNewChartToWorkbook"

    On Error Resume Next

    Dim shtResult As Worksheet
  Set shtResult = wbOld.Sheets(tChartProps.SheetName)
  
  On Error GoTo ErrorHandler

    If shtResult Is Nothing Then
  
     Set shtResult = wbOld.Sheets.Add(After:=ActiveWorkBook.Worksheets(ActiveWorkBook.Worksheets.count))
     
     shtResult.Name = tChartProps.SheetName

    Else
  
     Set shtResult = wbOld.Sheets(tChartProps.SheetName)
     
  End If

  Set shtNewWorkSheet = shtResult

 Exit Sub
ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


End Sub


Public Sub TurnOffGridLines()

    'Source code: https://stackoverflow.com/questions/40368373/how-can-i-turn-off-gridlines-in-excel-using-vba-without-using-activewindow

    Dim view As WorksheetView

    For Each view In shtNewWorkSheet.Parent.Windows(1).SheetViews

        If view.Sheet.Name = shtNewWorkSheet.Name Then

            view.DisplayGridlines = False
            Exit Sub

        End If

    Next

End Sub


Public Sub PrintHeaders()

    Const cProc = "PrintHeaders"

    On Error GoTo ErrorHandler

    Dim rOut As Range
    Set rOut = shtNewWorkSheet.Cells(tChartProps.FirstCellOfOutput(1, 1), tChartProps.FirstCellOfOutput(2, 1))
    
    rOut.ColumnWidth = 21.45

    rOut.Offset(1, 0).Value = tChartProps.Language_Descriptions(2)
    rOut.Offset(2, 0).Value = tChartProps.Language_Descriptions(1)
    rOut.Offset(3, 0).Value = tChartProps.Language_Descriptions(5)
    rOut.Offset(4, 0).Value = tChartProps.Language_Descriptions(11)
    rOut.Offset(5, 0).Value = tChartProps.Language_Descriptions(12)

    'rOut.Offset(6, 0).Value = tChartProps.Language_Descriptions(2)
    rOut.Offset(6, 0).Value = tChartProps.Language_Descriptions(3)
    rOut.Offset(7, 0).Value = tChartProps.Language_Descriptions(4)
    rOut.Offset(8, 0).Value = tChartProps.Language_Descriptions(10)

    rOut.Offset(1, 1).Resize(1, 4).Merge
    rOut.Offset(2, 1).Resize(1, 4).Merge
    rOut.Offset(3, 1).Resize(1, 4).Merge
    rOut.Offset(4, 1).Resize(1, 4).Merge
    rOut.Offset(5, 1).Resize(1, 4).Merge

    Call FormatRange(rOut.Offset(1, 1).Resize(1, 4))
    Call FormatRange(rOut.Offset(2, 1).Resize(1, 4))
    Call FormatRange(rOut.Offset(3, 1).Resize(1, 4))
    Call FormatRange(rOut.Offset(4, 1).Resize(1, 4))
    Call FormatRange(rOut.Offset(5, 1).Resize(1, 4))

    rOut.Offset(1, 0).Resize(8, 7).Borders(xlEdgeTop).LineStyle = xlContinuous
    rOut.Offset(1, 0).Resize(8, 7).Borders(xlEdgeBottom).LineStyle = xlContinuous

    If tChartProps.Title.Text <> "" Then

        rOut.Offset(2, 1).Value = tChartProps.Title.Text

    Else

        rOut.Offset(2, 1).Value = tChartProps.Language_Descriptions(8)

    End If

    rOut.Offset(3, 1).Value = tChartProps.Language_Descriptions(9)

    Dim rAd As Range
    Set rAd = Range("F1")
    rAd.Select
    rAd.ColumnWidth = tChartProps.ColumnWith

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub


