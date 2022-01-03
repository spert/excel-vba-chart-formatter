
Option Explicit

Private Const cModule = "clsError"

Private arrErrMess() As Variant

Public Sub ErrorMessage(strProc As String, strModule As String)
        
    Dim sError As String
    sError = " | Module: " & strModule & " | Proc: " & strProc & " |  ErrorNumber: " & Err.Number & " |  Description: " & Err.Description
   
    If (Not arrErrMess) = True Then
            
        ReDim Preserve arrErrMess(1)
        arrErrMess(1) = "1." & sError
            
    Else
         
            ReDim Preserve arrErrMess(UBound(arrErrMess) + 1)
            arrErrMess(UBound(arrErrMess)) = Str(UBound(arrErrMess)) & "." & sError
        
    End If

    On Error GoTo 0
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
        
End Sub

Public Property Get GetErrorCount() As Integer

        If (Not arrErrMess) = True Then
        
            GetErrorCount = 0
            Exit Property
        Else
            
            GetErrorCount = UBound(arrErrMess)
        
        End If
        
End Property

Public Property Get GetErrorMessages() As String

    Dim sTotal As String
    Dim i As Integer
    
    If (Not arrErrMess) = False Then
     
    Else
            
        For i = 1 To UBound(arrErrMess)
            
            sTotal = sTotal & "@" & arrErrMess(i)
        
        Next
        
     End If
     
     GetErrorMessages = sTotal
        
End Property

Public Sub ResetErrorCounter()

   Erase arrErrMess

End Sub