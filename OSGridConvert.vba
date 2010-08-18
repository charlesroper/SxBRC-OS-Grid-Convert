Option Explicit
Dim intGridLetterX(25) As Integer
Dim intGridLetterY(25) As Integer
Dim intDintyX(25) As Integer
Dim intDintyY(25) As Integer

Function EASTING(strGridIn As String) As Long
    EASTING = GridRefToCoord(strGridIn, True)
End Function

Function NORTHING(strGridIn As String) As Long
    NORTHING = GridRefToCoord(strGridIn, False)
End Function

Function EASTING_C(strGridIn As String) As Long
    EASTING_C = GridRefToCoord(strGridIn, True, True)
End Function

Function NORTHING_C(strGridIn As String) As Long
    NORTHING_C = GridRefToCoord(strGridIn, False, True)
End Function

Private Function GridRefToCoord(strGridRef As String, blnEorN As Boolean, Optional blnCentre As Boolean) As Long
    Dim strChar As String, _
    strNumbers As String, _
    strEastings As String, _
    strGridType As String, _
    strE As String, _
    strN As String, _
    i As Byte, _
    bytHalfLen As Byte, _
    bytIndex As Byte, _
    lngOSGridE As Long, _
    lngOSGridN As Long
    
    ' Clean the gridref
    strGridRef = UCase$(Trim$(strGridRef))
    strGridRef = CReplace(strGridRef, " ", "")
    
    ' Validate gridref and determine type
    strGridType = DetermineGridType(strGridRef)
    
    ' Resolve non-standard gridref types
    If strGridType <> "standard" Then
        Select Case strGridType
        Case "tetrad"
            strGridRef = ConvertDINTY(strGridRef)
        Case "5km"
        Case "invalid"
            Exit Function
        End Select
    End If
        
    ' Centre the gridref?
    If blnCentre Then
        strGridRef = CentreGridRef(strGridRef, strGridType)
    Else
        ' If we don't want the ref centred, pad it instead
        strGridRef = PadGridRef(strGridRef)
    End If
    
    ' Create the OS grid number arrays
    Call CreateOSGrid
    
    ' Split the numbers into Eastings and Northings
    bytHalfLen = Len(Mid$(strGridRef, 3)) / 2
    strE = Mid$(strGridRef, 3, bytHalfLen)
    strN = Mid$(strGridRef, bytHalfLen + 3)
    
    ' Convert first OS letter into E and N
    bytIndex = Asc(Left$(strGridRef, 1)) - 65
    lngOSGridE = CLng(intGridLetterX(bytIndex)) * 5000
    lngOSGridN = CLng(intGridLetterY(bytIndex)) * 5000
    
    ' Convert second OS letter into E and N
    bytIndex = Asc(Mid$(strGridRef, 2, 1)) - 65
    lngOSGridE = lngOSGridE + CLng(intGridLetterX(bytIndex)) * 1000
    lngOSGridN = lngOSGridN + CLng(intGridLetterY(bytIndex)) * 1000
    
    lngOSGridE = lngOSGridE + Val(strE) - 1000000
    lngOSGridN = lngOSGridN + Val(strN) - 500000
    
    If blnEorN Then
        GridRefToCoord = lngOSGridE
    Else
        GridRefToCoord = lngOSGridN
    End If
    
End Function

Private Function DetermineGridType(strGridIn As String) As String
' Function to determine the type of gridref (standard, DINTY, etc.)

    Dim strLetters As String, _
    strNumbers As String, _
    strChar As String, _
    strNum As String, _
    i As Byte
    
    ' Check for invalid characters
    For i = 1 To Len(strGridIn)
        strChar = Mid$(strGridIn, i, 1)
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890", strChar) = 0 Then
            DetermineGridType = "invalid"
            Exit Function
        End If
    Next
    
    ' Check to see if OS letters are valid
    strLetters = Left$(strGridIn, 2)
    If InStr("I0123456789", Left$(strLetters, 1)) Or _
        InStr("I0123456789", Right$(strLetters, 1)) Then
        DetermineGridType = "invalid"
        Exit Function
    End If
    
    ' Check to see if numbers are valid
    For i = 1 To Len(strGridIn)
        strNum = Mid$(strGridIn, i, 1)
        If InStr("0123456789", strNum) Then
            strNumbers = strNumbers & strNum
        End If
    Next
    If Len(strNumbers) Mod 2 Then ' No of digits uneven?
        DetermineGridType = "invalid"
        Exit Function
    End If
    
    ' Check letters on end of gridref (DINTY etc.)
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Right$(strGridIn, 1)) Then
        Select Case Len(strGridIn)
        Case 5
            If Right$(strGridIn, 1) <> "O" Then
                DetermineGridType = "tetrad"
                Exit Function
            End If
            DetermineGridType = "invalid"
            Exit Function
        Case 6
            Select Case Right(strGridIn, 2)
            Case "SW", "SE", "NW", "NE"
                DetermineGridType = "5km"
                Exit Function
            Case Else
                DetermineGridType = "invalid"
                Exit Function
            End Select
        Case Else
            DetermineGridType = "invalid"
            Exit Function
        End Select
    End If
    
    DetermineGridType = "standard"
End Function

Private Sub CreateOSGrid()
    Dim bytChar As Byte, _
    intHorizontal As Integer, _
    intVertical As Integer
    
    bytChar = Asc("A") - 65
    For intVertical = 400 To 0 Step -100
        For intHorizontal = 0 To 400 Step 100
            intGridLetterY(bytChar) = intVertical
            intGridLetterX(bytChar) = intHorizontal
            bytChar = bytChar + 1
            If bytChar = (Asc("I") - 65) Then bytChar = bytChar + 1
        Next
    Next
End Sub

Private Function ConvertDINTY(strDintyIn As String)
' Convert the DINTY letter to relevant coordinates

    Dim strE As String, _
    strN As String, _
    strLetter As String, _
    strOut As String
        
    Call CreateDINTYGrid
    
    strE = Mid$(strDintyIn, 3, 1)
    strN = Mid$(strDintyIn, 4, 1)
    strLetter = Right$(strDintyIn, 1)
    
    strOut = strE & intDintyX(Asc(strLetter) - 65) & _
        strN & intDintyY(Asc(strLetter) - 65)
    
    ConvertDINTY = Left$(strDintyIn, 2) & strOut
End Function

Private Sub CreateDINTYGrid()
    Dim bytChar As Byte, _
    intHorizontal As Integer, _
    intVertical As Integer
    
    bytChar = Asc("A") - 65
    For intHorizontal = 0 To 8 Step 2
        For intVertical = 0 To 8 Step 2
            intDintyX(bytChar) = intHorizontal
            intDintyY(bytChar) = intVertical
            bytChar = bytChar + 1
            If bytChar = (Asc("O") - 65) Then bytChar = bytChar + 1
        Next
    Next
End Sub

Private Function CentreGridRef(strGridIn As String, strGridType As String) As String
    Dim strE As String, _
    strN As String, _
    grLen As Integer, _
    offset As Integer
    
    grLen = Len(strGridIn)
    
    ' Calculate offset value
    Select Case grLen
    Case 12
        offset = 0
    Case 10
        offset = 5
    Case 8
        offset = 50
    Case 6
        offset = 500
    Case 5
        offset = 1000
    Case 4
        offset = 5000
    Case 2
        offset = 50000
    Case Else
        offset = 0
    End Select
    
        
    ' Split the numbers from the gridref into eastings and northings
    strE = Mid$(strGridIn, 3, (grLen - 2) / 2)
    strN = Right$(strGridIn, (grLen - 2) / 2)
    
    Select Case strGridType
    Case "standard"
    ' Pad eastings and northings with fives upto a total of 5 digits
        strE = Left$(strE & offset, 5)
        strN = Left$(strN & offset, 5)
    Case "tetrad"
        strE = Left$(strE, 1) & Val(Right$(strE, 1)) + 1
        strE = Left$(strE & "00000", 5)
        strN = Left$(strN, 1) & Val(Right$(strN, 1)) + 1
        strN = Left$(strN & "00000", 5)
    End Select
    
    CentreGridRef = Left$(strGridIn, 2) & strE & strN
End Function
Private Function PadGridRef(strGridIn As String) As String
    Dim strE As String, _
    strN As String
    
    ' Split the numbers from the gridref into eastings and northings
    strE = Mid$(strGridIn, 3, (Len(strGridIn) - 2) / 2)
    strN = Right$(strGridIn, (Len(strGridIn) - 2) / 2)

    ' Pad the eastings and northings with zeros upto a total of 5 digits
    strE = Left$(strE & "00000", 5)
    strN = Left$(strN & "00000", 5)
    
    ' Re-combine the letters, eastings and northings
    PadGridRef = Left$(strGridIn, 2) & strE & strN
End Function

Function CReplace(strIn As String, strToReplace As String, strReplaceWith As String) As String
    Dim strOut As String
    Dim intPos As Integer
    strOut = Trim$(strIn)
    Do While InStr(strOut, strToReplace)
        intPos = InStr(strOut, strToReplace)
        strOut = Left$(strOut, intPos - 1) & strReplaceWith & _
            Right$(strOut, Len(strOut) - intPos)
    Loop
    CReplace = strOut
End Function
