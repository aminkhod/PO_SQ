Attribute VB_Name = "Module1"
Function NumberstoWords(ByVal pNumber)
'Updateby20140220
Dim Dollars
arr = Array("", "", " Thousand ", " Million ", " Billion ", " Trillion ")
pNumber = Trim(Str(pNumber))
xDecimal = InStr(pNumber, ".")
If xDecimal > 0 Then
pNumber = Trim(Left(pNumber, xDecimal - 1))
End If
xIndex = 1
Do While pNumber <> ""
xHundred = ""
xValue = Right(pNumber, 3)
If Val(xValue) <> 0 Then
xValue = Right("000" & xValue, 3)
If Mid(xValue, 1, 1) <> "0" Then
xHundred = GetDigit(Mid(xValue, 1, 1)) & " Hundred "
End If
If Mid(xValue, 2, 1) <> "0" Then
xHundred = xHundred & GetTens(Mid(xValue, 2))
Else
xHundred = xHundred & GetDigit(Mid(xValue, 3))
End If
End If
If xHundred <> "" Then
Dollars = xHundred & arr(xIndex) & Dollars
End If
If Len(pNumber) > 3 Then
pNumber = Left(pNumber, Len(pNumber) - 3)
Else
pNumber = ""
End If
xIndex = xIndex + 1
Loop
NumberstoWords = Dollars
End Function
Function GetTens(pTens)
Dim Result As String
Result = ""
If Val(Left(pTens, 1)) = 1 Then
Select Case Val(pTens)
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
Else
Select Case Val(Left(pTens, 1))
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
Result = Result & GetDigit(Right(pTens, 1))
End If
GetTens = Result
End Function
Function GetDigit(pDigit)
Select Case Val(pDigit)
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
