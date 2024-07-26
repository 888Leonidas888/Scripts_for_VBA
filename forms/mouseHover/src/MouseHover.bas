Attribute VB_Name = "MouseHover"
Sub formatFrame(frame As MSForms.frame)
    
    For Each lbl In frame.Controls
        lbl.BorderStyle = 0
    Next lbl
    
End Sub
Sub formatLabel(lbl As MSForms.Label)

     With lbl
        .BorderStyle = 1
        .BorderColor = RGB(0, 0, 255)
    End With
    
End Sub
