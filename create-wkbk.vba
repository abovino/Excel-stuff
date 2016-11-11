Sub createNewWorkbook()
' Creates a new work book from the BOILERPLATE sheet

    Sheets("BOILERPLATE").Select
    Sheets("BOILERPLATE").Copy
End Sub

Sub createCSV()
' Creates a CSV File of the sheet BOILERPLATE

    Dim maxCharCount As Integer, epiTasks As Integer, badAddress As Integer
    maxCharCount = Sheets("FORM").Range("K12")
    epiTasks = Sheets("FORM").Range("K10")
    badAddress = Sheets("FORM").Range("K11")
    
    If maxCharCount + epiTasks + badAddress < 1 Then
    
        Dim todaysDate As String
        todaysDate = Date
        todaysDate = Replace(todaysDate, "/", "-")
        
        Dim username As String
        username = Workbooks("MKGT Task Upload Generator.xlsm").Sheets("FORM").Range("I2")
        
        Dim campaign As String
        campaign = Workbooks("MKGT Task Upload Generator.xlsm").Sheets("FORM").Range("I3")
        
        Dim fileName As String
        fileName = username & " " & campaign & " " & todaysDate
        
        Dim dirPath As String
            dirPath = "[file path here]"
        
        Sheets("BOILERPLATE").Select
        Sheets("BOILERPLATE").Copy
        ActiveWorkbook.SaveAs fileName:= _
            dirPath & fileName & ".csv", FileFormat:=xlCSV, _
            CreateBackup:=False
            
        ActiveSheet.Unprotect
        
        MsgBox "CSV file created succesfully!" & vbNewLine & vbNewLine & "File Location: " & dirPath & fileName
        
    Else
        
        MsgBox "Correct the following errors: " & vbNewLine & vbNewLine & "-EPI Tasks cannot be assigned in Sugar" & vbNewLine & "  EPI Tasks: " & epiTasks & vbNewLine & vbNewLine & "-Check for correct addresses" & vbNewLine & "    Invalid addresses: " & badAddress & vbNewLine & vbNewLine & "-Subject length is > 50 characters" & vbNewLine & "    Invalid Subject lengths: " & maxCharCount
        
    End If
        
End Sub
