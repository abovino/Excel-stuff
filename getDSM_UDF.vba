'User Defined Function that takes in a 5 digit zip code as the first argument and a 2 letter state abbreviation
'as the second argument and returns the name of a person who covers that territory.
'NOTES:
'-State abbreviation cannot contain periods (Add function to remove special chars)
'-State abbreviation must be all caps (use =UPPER(A1) to convert a string to all caps. Don't add this to getDSM() because the function can take very long to process)
'-Once you have correctly parsed the data it's recommended that you copy the column and paste it over it's self as text otherwise the spreadsheet will be very slow

Function getDSM(zipCode As String, state As String) As String
    
    'If you have many rows of data use the 2 below functions to parse the zip codes and states BEFORE running getDSM().  They will make the function run slower
    'zipCode = Left(zipCode, 5)
    'state = Upper(state)
    
    'Northeast
    Dim neUser1 As String
    Dim neUser2 As String
    Dim neUser3 As String
    Dim neUser4 As String
    
    user1 = "NY"
    user2 = "CT,MA,ME,NH,RI,VT"
    user3 = "MD,VA,WV"
    user4 = "DE,NJ,PA"
    
    'West (Only non-shared states)
    Dim wUser1 As String
    Dim wUser2 As String
    Dim wUser3 As String
    
    wUser1 = "AK,ID,MT,OR,WA"
    wUser2 = "AZ,NV,UT"
    wUser3 = "HI"
    
    'Zips for shared states
    Dim zipUser1 As String
    Dim zipUser2 As String
    Dim zipUser3 As String
    
    '~150 MAX zip codes per line.  Use '& _ ' at end of line to continue string on the next line
    zipUser1 = "00001, 00002, 00003" & _
               "00004, 00005, 00006" & _
               "00007, 00008, 00009"
    zipUser2 = "11111, 11112, 11113" & _ 
               "11114, 11115, 11116" & _
               "11117, 11118, 11119"
    zipUser2 = "22221, 22222, 22223" & _ 
               "22224, 22225, 22226" & _ 
               "22227, 22228, 22229"
    
    Select Case True
        'Northeast
        Case InStr(neUser1, state)
            result = "USER's NAME"
        Case InStr(neUser2, state)
            result = "USER'S NAME"
        Case InStr(neUser3, state)
            result = "USER'S NAME"
        Case InStr(neUser4, state)
            result = "USER'S NAME"
        'West
        Case InStr(wUser1, state)
            result = "USER'S NAME"
        Case InStr(wUser2, state)
            result = "USER'S NAME"
        Case InStr(wUser3, state)
            result = "USER'S NAME"
        'Zip Codes
        Case InStr(zipUser1, zipCode)
            result = "USER'S NAME"
        Case InStr(zipUser2, zipCode)
            result = "USER'S NAME"
        Case InStr(zipUser3, zipCode)
            result = "USER'S NAME"
        'Error Handling
        Case Else
            result = "NOT FOUND"
    End Select
    
    getDSM = result
    
End Function
