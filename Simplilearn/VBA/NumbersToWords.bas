Attribute VB_Name = "Module2"
Function NumbersToWords(ByVal MyNumber As Currency) As String
    Dim DecimalPlace As Integer
    Dim Count As Integer
    Dim DecimalSeparator As String
    Dim Temp As String
    Dim Dollars As String
    Dim Cents As String
    Dim DecimalWord As String
    
    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "
    
    DecimalSeparator = "."
    
    ' Convert MyNumber to String, trimming extra spaces.
    Temp = Trim(CStr(MyNumber))
    
    ' Find position of decimal place.
    DecimalPlace = InStr(Temp, DecimalSeparator)
    
    ' Extract dollars and cents.
    If DecimalPlace > 0 Then
        Dollars = Left(Temp, DecimalPlace - 1)
        Cents = Mid(Temp, DecimalPlace + 1)
    Else
        Dollars = Temp
        Cents = "00"
    End If
    
    ' Convert dollars and set MyNumber to cents amount.
    NumbersToWords = GetWords(Dollars) & " Dollars"
    
    ' Convert cents and add "and Cents" if there are cents.
    If Cents <> "00" Then
        DecimalWord = GetWords(Cents)
        If DecimalWord <> "" Then
            NumbersToWords = NumbersToWords & " and " & DecimalWord & " Cents"
        End If
    End If
End Function

Private Function GetWords(ByVal MyNumber As String) As String
    Dim Units As String
    Dim SubUnits As String
    Dim DecimalPlace As Integer
    Dim Count As Integer
    Dim DecimalSeparator As String
    Dim UnitName As String
    Dim SubUnitName As String
    DecimalSeparator = "."
    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "
    ' Convert MyNumber to String, trimming extra spaces.
    MyNumber = Trim(CStr(MyNumber))
    ' Find position of decimal place.
    DecimalPlace = InStr(MyNumber, DecimalSeparator)
    ' If we find decimal place...
    If DecimalPlace > 0 Then
        ' Convert SubUnits and set MyNumber to Units amount.
        SubUnits = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    Count = 1
    Do While MyNumber <> ""
        TempStr = GetHundreds(Right(MyNumber, 3))
        If TempStr <> "" Then Units = TempStr & Place(Count) & Units
        If Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop
    GetWords = Units & SubUnits
End Function

Private Function GetHundreds(ByVal MyNumber As String) As String
    Dim Result As String
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    ' Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
    End If
    ' Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & GetTens(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid(MyNumber, 3))
    End If
    GetHundreds = Result
End Function

Private Function GetTens(TensText As String) As String
    Dim Result As String
    Result = ""           ' Null out the temporary function value.
    If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...
        Select Case Val(TensText)
            Case 10: Result = "Ten"
            Case 11: Result = "Eleven"
            Case 12: Result = "Twelve"
            Case 13: Result = "Thirteen"
            Case 14: Result = "Fourteen"
            Case 15: Result = "Fifteen"
            Case 16: Result = "Sixteen"
            Case 17: Result = "Seventeen"
            Case 18: Result = "Eighteen"
            Case 19: Result = "Nineteen"
            Case Else
        End Select
    Else                                 ' If value between 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: Result = "Twenty "
            Case 3: Result = "Thirty "
            Case 4: Result = "Forty "
            Case 5: Result = "Fifty "
            Case 6: Result = "Sixty "
            Case 7: Result = "Seventy "
            Case 8: Result = "Eighty "
            Case 9: Result = "Ninety "
            Case Else
        End Select
        Result = Result & GetDigit(Right(TensText, 1))   ' Retrieve ones place.
    End If
    GetTens = Result
End Function

Private Function GetDigit(Digit As String) As String
    Select Case Val(Digit)
        Case 1: GetDigit = "One"
        Case 2: GetDigit = "Two"
        Case 3: GetDigit = "Three"
        Case 4: GetDigit = "Four"
        Case 5: GetDigit = "Five"
        Case 6: GetDigit = "Six"
        Case 7: GetDigit = "Seven"
        Case 8: GetDigit = "Eight"
        Case 9: GetDigit = "Nine"
        Case Else: GetDigit = ""
    End Select
End Function

