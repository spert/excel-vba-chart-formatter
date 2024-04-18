Option Explicit On

Const cModule = "clsWorkSheet"

Dim wbOld As Workbook
Private tChartProps As MyChart

Private rOut As Range

Private shtNewWorkSheet As Worksheet

Public Sub Initiate()

  Set wbOld = ActiveWorkbook

End Sub

Property Get GetNewWorkSheet() As Worksheet

    Set GetNewWorkSheet = shtNewWorkSheet

End Property


Public Sub AddNewChartToWorkbook()

    Const cProc = "AddNewChartToWorkbook"

    On Error Resume Next

    Dim shtResult As Worksheet
  Set shtResult = wbOld.Sheets()
  
  On Error GoTo ErrorHandler

    Dim strSheetName As String
    strSheetName = GetNextSheetName()

    If shtResult Is Nothing Then
  
     Set shtResult = wbOld.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
     
     shtResult.Name = strSheetName

    Else
  
     Set shtResult = wbOld.Sheets(strSheetName)
     
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


