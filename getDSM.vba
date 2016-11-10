'User Defined Function that takes in a 5 digit zip code as the first argument and a 2 letter state abbreviation
'as the second argument and returns the name of a person who covers that territory.
'NOTES:
'-Zip codes must be 5 digits
'-State abbreviation cannot contain periods (Add function to remove special chars)
'-State abbreviation must be all caps (Function to make string all caps, or convert arguments to all caps)

Function getDSM(zipCode As String, state As String) As String
    
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
