Module Mastermind

    'Everything that is in Sub Main() right now should probably be moved to another sub.
    Sub Main()

        'Causes random number generation to use a unique seed each time the program is run
        Randomize()
        'Will be used to check if the game mode is set to custom or standard. When the program is run for the first time
        ' the user's setting will be saved in an ini file (or cookie if web based I suppose?)
        Dim game_is_custom As Boolean
        'Set to false as a temporary measure, because I'm too lazy to remember how to read files at the moment.
        game_is_custom = False
        'Holds the number of different coloured pins to be used in the game. User will be able to set a custom value,
        ' perhaps remember their previous setting in ini file?
        Dim number_of_colours As Byte
        'Calls the function to get the number of colours to be used, sends back the default value of 6 unless the game
        ' is set to custom mode, in which case the user will be able to enter their own value.
        number_of_colours = get_num_colours(game_is_custom)
        'This is the array that will hold the hidden sequence that the user must guess to win the game.
        Dim seq() As Byte
        'Calls the function that determines the length of the hidden sequence, same idea as with the colours.
        ReDim seq(0 To seq_size(game_is_custom))
        'Loop counter variable, woooo
        Dim x As Byte
        'Runs through the index values of the seq array, calling the random number function to fill each one.
        For x = LBound(seq) To UBound(seq)
            seq(x) = get_random_num(number_of_colours)
            'Shows the value that was just generated and stored, testing purposes only obviously
            MsgBox(CStr(seq(x)))
        Next
        'Again for testing purposes, shows all the values that went into seq
        MsgBox(CStr(seq(0)) & ", " & CStr(seq(1)) & ", " & CStr(seq(2)) & ", " & CStr(seq(3)) & ".")

    End Sub

    'Generates the random numbers, takes an upper bound and a lower bound to determine the range for *rnd
    Function get_random_num(ByVal ub As Byte, Optional ByVal lb As Byte = 0) As Byte
        Dim temp As Byte

        'Generates random number within the range, then corrects the number as necessary (e.g. if you want 2 to 7 it will
        ' get a number between 0 and 5, then add 2 to ensure it's within the range 2 to 7); turns the number into a byte
        ' as well to deal with decimals, but seriously everything is already bytes in this function what the fuck am I doing?
        temp = CByte(Rnd() * (ub - lb))
        temp = temp + lb

        'Returns the random number
        get_random_num = temp

    End Function

    'Determines length of the hdiden sequence
    Function seq_size(ByVal custom As Boolean) As Byte
        If custom = False Then
            seq_size = 3
        Else
            'This is where the code to let the user set the size will go. An extra elseif will have to be added for ini file
            ' I suppose.
            seq_size = 4
        End If
    End Function

    'Determines number of colours to be used
    Function get_num_colours(ByVal custom As Boolean) As Byte
        If custom = False Then
            get_num_colours = 6
        Else
            'Same deal as the sequence function. I wonder if I could combine the two functions in some way, or if doing that much
            ' in a single function would defeat the purpose of it.
            get_num_colours = 7
        End If
    End Function

End Module
