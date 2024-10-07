Attribute VB_Name = "converter"
Function toString(value As Variant) As String
        
    toString = "'" & value & "'"

End Function

Function toDate(value As Variant) As String

    toDate = "#" & Format(value, "mm/dd/yyyy") & "#"
    
End Function
