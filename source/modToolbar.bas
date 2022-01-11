
Option Explicit

Dim frmCht As frmChartFormat
Const cModule = "modToolbar"

Sub CreateMyMenu()
    
    Dim myCB As CommandBar
    Dim myCBtn As CommandBarButton
    
    On Error Resume Next
    Application.CommandBars("ChtFormat").Delete

    Set myCB = CommandBars.Add(Name:="ChtFormat", Position:=msoBarFloating)
    
    Set myCBtn = myCB.Controls.Add(Type:=msoControlButton)
    
    With myCBtn
     .FaceId = 17
     .Style = msoButtonIconAndCaption
     .Caption = "Format chart"
     .OnAction = "OpenUSerForm"
    End With
 
    ' Show the command bar
    myCB.Visible = True
    
End Sub
 
Private Sub OpenUSerForm()

   Const cProc = "OpenUSerForm"

   On Error GoTo ErrorHandler

    Dim wbW As Workbook
    Set wbW = ActiveWorkBook

    If wbW Is Nothing Then

       Err.Raise Number:=9999, Description:="@Open an Excel workbook first!"
        
    End If

    If wbW.ReadOnly Then
    
      Err.Raise Number:=9999, Description:="@Read only Excel files are not permitted!"
      
    End If
    
    If wbW.Saved = False Then
    
      wbW.Save
    
    End If
    
    
    Dim oChtOb As ChartObject
    Set oChtOb = GetChartObject(Selection)

    If oChtOb Is Nothing Then
    
        Err.Raise Number:=9999, Description:="@Select a chart first to be formatted!"
    
    Else
    
        Set frmCht = New frmChartFormat
        frmCht.SetChartObject = oChtOb
        frmCht.Show
        
    End If
    
    
Exit Sub
    
ErrorHandler:

  ErrorMod.ErrorMessage cProc, cModule
  
   Dim frmError As frmMessage
   Set frmError = New frmMessage
   frmError.SetErrorMessage = ErrorMod.GetErrorMessages
   frmError.Show
    
End Sub




