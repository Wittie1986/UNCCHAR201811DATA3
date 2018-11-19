Sub SentenceBreaker()

    ' Retrieve the user sentence and store in variable
    ' <YOUR CODE GOES HERE>
    Dim words() as String
    words = Split(Range("B1").value," ")



    ' Retrieve the user word numbers and store in variables 
    ' <YOUR CODE GOES HERE>
    Dim num1 As Integer
    Dim num2 As Integer
    Dim num3 As Integer

    num1 = Cells(4, 1).Value
    num2 = Cells(5, 1).Value
    num3 = Cells(6, 1).Value

    Cells(4, 2).Value = words(num1 - 1)
    Cells(5, 2).Value = words(num2 - 1)
    Cells(6, 2).Value = words(num3 - 1)




    ' Split the user's sentence into words
    ' <YOUR CODE GOES HERE>



    ' Use the word numbers to retrieve the specific words in the sentence
    ' Remember to offset by the 0 index
    ' <YOUR CODE GOES HERE>


    
End Sub