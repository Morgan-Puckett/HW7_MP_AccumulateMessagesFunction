'Morgan Puckett
'Rcet0265
'Spring 2022
'Accumulate messages
'https://github.com/Morgan-Puckett/HW7_MP_AccumulateMessagesFunction.git



Option Explicit On
Option Strict On
Option Compare Text

Module HW7_MP_AccumulateMessagesFunction
    Sub Main()
        Dim userInput As String
        Do
            Do
                'this displays the fuctionality of UserMessages Function
                Console.WriteLine("If you press ""Q"" the program will end")
                userInput = Console.ReadLine()
                UserMessages(userInput, False)
                Console.WriteLine(UserMessages("", False))
            Loop While userInput <> "Q"
            UserMessages(CStr(vbEmpty), True)
        Loop Until userInput = "Q"
        Console.ReadLine()
    End Sub


    'this function will accumulate the messages by adding the latest message on a new line
    Function UserMessages(ByVal newMessage As String, ByVal clear As Boolean) As String
        Static messages As String
        If clear = True Then
            messages = ""
        ElseIf newMessage <> "" Then
            messages &= newMessage & vbNewLine

        End If
        Return messages
    End Function
End Module
