
Option Explicit


Private Sub btnClose_Click()

    Unload Me

End Sub

Property Let SetErrorMessage(sErrMessage As String)

   lstErrors.Clear

   Dim i As Integer
    
   Dim aErrors() As String
   aErrors = Split(sErrMessage, "@")
        
           
    For i = 1 To UBound(aErrors)
            
            lstErrors.AddItem aErrors(i)
            
    Next
  

End Property
