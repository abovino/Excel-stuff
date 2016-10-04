'Use removeSpecial(A1) to remove special characters in cell A1
'To add or remove special characters from the list modify line 7 sSpecialChars

Function removeSpecial(sInput As String) As String
    Dim sSpecialChars As String
    Dim i As Long
    sSpecialChars = "\/:*?""<>|&.,- "
    For i = 1 To Len(sSpecialChars)
        sInput = Replace$(sInput, Mid$(sSpecialChars, i, 1), "")
    Next
    removeSpecial = sInput
End Function
