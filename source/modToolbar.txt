
Option Explicit

Dim frmCht As frmChartFormat


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

   On Error GoTo ErrorHandler
   
   Dim frmError As frmMessage
    
    Dim wbW As Workbook
    Set wbW = ActiveWorkBook

    If wbW Is Nothing Then

        Set frmError = New frmMessage
        frmError.SetErrorMessage = "@Open an Excel workbook first!"
        frmError.Show
        Exit Sub
        
    End If
    
'    If wbW.Saved = False Then
'
'        Set frmError = New frmMessage
'        frmError.SetErrorMessage = "@Save the active Excel workbook first!"
'        frmError.Show
'        Exit Sub
'
'    End If

    wbW.Save
    
    Dim oChtOb As ChartObject
    Set oChtOb = GetChartObject(Selection)

    If oChtOb Is Nothing Then
    
        Set frmError = New frmMessage
        frmError.SetErrorMessage = "@Select a chart first to be formatted!"
        frmError.Show
        Exit Sub
    
    Else
    
        Set frmCht = New frmChartFormat
        frmCht.SetChartObject = oChtOb
        frmCht.Show
        
    End If
    
    
    
Exit Sub
    
ErrorHandler:

MsgBox Err.Description
    
End Sub



