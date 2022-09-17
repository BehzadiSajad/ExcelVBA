Sub UnhideAll()
'
' Unhide_all sheets Macro
'
' Keyboard Shortcut: Ctrl+Shift+H
'

Dim WS As Worksheet

    For Each WS In Worksheets
        WS.Visible = True
    Next
    
    ActiveWindow.TabRatio = 1
    Sheets(1).Select

End Sub
