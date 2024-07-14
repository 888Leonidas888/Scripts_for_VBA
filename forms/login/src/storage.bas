Attribute VB_Name = "storage"
Public Function validateLogin(user As String, password As String) As Boolean
    
    If user = "john" And password = "doe" Then
        validateLogin = True
    Else
        validateLogin = False
    End If
    
End Function
