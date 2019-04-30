Attribute VB_Name = "Module1"
Function fff()
    For Each w In Application.Workbooks
        w.Save
    Next w
    Application.Quit
End Function

