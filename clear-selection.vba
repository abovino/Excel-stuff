Sub clearTasks()
' Clears tasks created by the autoFillTasks macro in the boiler spreadsheet

    Sheets("BOILERPLATE").Select
    Rows("3:3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("FORM").Select
    
    MsgBox "Tasks deleted succesfully"
    
End Sub
