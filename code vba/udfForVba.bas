Attribute VB_Name = "udfForVba"
'Antes de usar esta función debes asegurarte de tener activadas las siguientes REFERENCIAS:

'1.-Microsft Html Object Library
'2.-Microsoft VBscript Regular Expressions 5.5

Public Function getIpv4() As String

    Dim html As New MSHTML.HTMLDocument
    Dim text As String
    
    Dim arrIp As Variant
    Dim ipLine As String
    
    Dim rg As New RegExp
    Dim ipCollection As MatchCollection
    Dim ip As String
    
    On Error GoTo Cath

    Shell "cmd.exe /k ipconfig | clip", vbHide
    
    Application.Wait Now() + TimeValue("00:00:01")

    text = html.parentWindow.clipboardData.GetData("text")
  

    arrIp = Split(text, vbCrLf)

    For i = LBound(arrIp) To UBound(arrIp)
        If arrIp(i) Like "*Dirección IPv4*" Then
            ipLine = arrIp(i)
        End If
    Next i


    With rg
        .Pattern = "\d+\.\d+\.\d+\.\d+"
         Set ipCollection = .Execute(ipLine)
        
        If Not ipCollection Is Nothing Then
            If ipCollection.Count = 1 Then
    '            Debug.Print "Este es tu ip v4 --> "; ipCollection.Item(0)
                ip = ipCollection.Item(0)
            End If
        End If
    End With

    getIpv4 = ip

Exit Function

Cath:
    Debug.Print Err.Number
    Debug.Print Err.Description

End Function
