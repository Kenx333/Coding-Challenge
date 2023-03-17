# Coding-Challenge
Option Explicit

Sub romanNumeralsConverter()

'Declaring ariables
Dim romanNumeralInput As String
Dim decimalInput As Integer
Dim decision1 As Boolean
Dim decision2 As Boolean
Dim decimalResult As Integer



'Recieving input from user
romanNumeralInput = InputBox("Please type a roman numeral you wish to convert into a decimal")

'Calling romanToDecimal function
decimalResult = romanToDecimal(romanNumeralInput)
    
    'Input validation and output
    If romanToDecimal(romanNumeralInput) = -1 Then
    MsgBox ("Invalid Input")
    Else
    MsgBox (romanNumeralInput & " is " & decimalResult)
    End If
Debug.Print
End Sub

'Beginning of numeral to decimal converter function
Function romanToDecimal(romanStr As String) As Integer

'Declaring variables
Dim counter As Integer
Dim currentValue As Integer
Dim prevValue As Integer
Dim result As Integer
Dim maxValue As Integer

    'Initializing maxValue
    maxValue = 0
    prevValue = 0
    result = 0
    
    'Beginning of for loop that executes the conversion
    'Takes the input and records each letter and does the calculation
    For counter = Len(romanStr) To 1 Step -1
    'Calls the romanNumeralValue function and sets it = to currentValue
    currentValue = romanNumeralValue(Mid(romanStr, counter, 1))
    'Executes the calculation and validates input to make sure that it doesn't violate rule #4
    If Len(romanStr) Mod 2 = 0 Then
        If currentValue < prevValue * 10 + 1 Then
            romanToDecimal = -1
            Exit Function
        
        ElseIf currentValue < prevValue Then
            result = result - currentValue
        ElseIf currentValue >= maxValue Then
            result = result + currentValue
            maxValue = currentValue
        Else
            romanToDecimal = -1
            Exit Function
        End If
    Else
        If currentValue < prevValue Then
            result = result - currentValue
        ElseIf currentValue >= maxValue Then
            result = result + currentValue
            maxValue = currentValue
        Else
            romanToDecimal = -1
            Exit Function
        End If
     End If
     
       
     'Sets prevValue equal to currentValue before returning back up to the beginning of the loop
     prevValue = currentValue
     
    'Increments to the next letter
    Next counter
    
 'Returns result
romanToDecimal = result

'End of romanToDecimal function
End Function

'Beginning of romanNumeralValue function to assign each letter to a number
Function romanNumeralValue(numeralStr As String) As Integer
  
  'Selelct case that assigns every letter to their respective numbers
  Select Case numeralStr
    Case "I": romanNumeralValue = 1
    Case "V": romanNumeralValue = 5
    Case "X": romanNumeralValue = 10
    Case "L": romanNumeralValue = 50
    Case "C": romanNumeralValue = 100
    Case "D": romanNumeralValue = 500
    Case "M": romanNumeralValue = 1000
  End Select

'End of romanNumeralValue function
End Function


'NOTES:
    'I recognize that my code is not perfect when it comes to complying with rule #4 but this is the best I could come up with in 4 hours.
    'I was not able to finish the complete solution that converts a decimal number to a roman numeral.
    'I could have used chatGPT but I would have considered that to be dishonest since I would have been unable to come up with a solution within 4 hours.
    'I do although have a rough idea of what I could have done to complete the second half of the solution.
    'First I would create a few question boxes that asked the user whether or not they wanted to convert roman numerals to decimal numbers or from decimal numbers to roman numerals.
    'This could easily be done through a few if/else statements
    'Once the user has decided, an input box would then show up.
    'If converting decimals to roman numerals was selected then it would procede with a function similar to the function I currently have made.
    'The function would start out similarly to the first function but would then procede to a series of if statements that would take the input and then divide first by 1000 and then by 100 aso.
    'Then if the remainder would be either 9, 5 or 4 then I would set the result the respective letters.
    '9,5 and 4 are in each if statement that tackle hundreds, tens and ones since they are the exceptions to the formatting of roman numerals.
    'I would then connect all the letters and then return the result as output.
    'I spent 4 hours on this chalange

'OUTCOMES:
    'This challange stretched the limits of my abillities, so I turned to the internet to better understand concepts to be able to approach the project more confidently.
    'I got a deeper understanding of functions and I was able to implement them succesfully.
    'Because my experince is limited this challenge was among the most complexs and difficult that I have taken on, but I am passionate about coding and cannot wait to see what I learn next.
    
