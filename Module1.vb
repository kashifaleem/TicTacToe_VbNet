Module Module1

    'developed in the Visual Studio 2008 (Console Mode) programming environment (VB.NET)

    Dim Board(3, 3) As Char
    Dim PlayerOneName As String
    Dim PlayerTwoName As String
    Dim PlayerOneScore As Single
    Dim PlayerTwoScore As Single
    Dim XCoord As Integer
    Dim YCoord As Integer
    Dim NoOfMoves As Integer
    Dim ValidMove As Boolean
    Dim GameHasBeenWon As Boolean
    Dim GameHasBeenDrawn As Boolean
    Dim CurrentSymbol As Char
    Dim StartSymbol As Char
    Dim PlayerOneSymbol As Char
    Dim PlayerTwoSymbol As Char
    Dim Answer As Char

    Sub Main()
        Randomize()
        Console.Write("What is the name of player one? ")
        PlayerOneName = Console.ReadLine()
        Console.Write("What is the name of player two? ")
        PlayerTwoName = Console.ReadLine()
        Console.WriteLine()
        PlayerOneScore = 0
        PlayerTwoScore = 0
        Do 'Choose player one’s symbol
            Console.Write(PlayerOneName & " what symbol do you wish to use, X or O? ")
            PlayerOneSymbol = Console.ReadLine()
            Console.WriteLine()
            If Not (PlayerOneSymbol = "X" Or PlayerOneSymbol = "O") Then
                Console.WriteLine("Symbol to play must be uppercase X or O")
                Console.WriteLine()
            End If
        Loop Until PlayerOneSymbol = "X" Or PlayerOneSymbol = "O"
        If PlayerOneSymbol = "X" Then
            PlayerTwoSymbol = "O"
        Else
            PlayerTwoSymbol = "X"
        End If
        StartSymbol = GetWhoStarts()
        Do 'Play a game
            NoOfMoves = 0
            GameHasBeenDrawn = False
            GameHasBeenWon = False
            ClearBoard(Board)
            Console.WriteLine()
            DisplayBoard(Board)
            If StartSymbol = PlayerOneSymbol Then
                Console.WriteLine(PlayerOneName & " starts playing " & StartSymbol)
            Else
                Console.WriteLine(PlayerTwoName & " starts playing " & StartSymbol)
            End If
            Console.WriteLine()
            CurrentSymbol = StartSymbol
            Do 'Play until a player wins or the game is drawn
                Do 'Get a valid move
                    GetMoveCoordinates(XCoord, YCoord)
                    ValidMove = CheckValidMove(XCoord, YCoord, Board)
                    If Not ValidMove Then
                        Console.WriteLine("Coordinates invalid, please try again")
                    Else
                        'Check if move is already taken
                        ValidMove = IsEmpty(XCoord, YCoord, Board)
                        If Not ValidMove Then Console.WriteLine("Coordinates are already marked, please try again")
                    End If
                Loop Until ValidMove
                Board(XCoord, YCoord) = CurrentSymbol
                DisplayBoard(Board)
                GameHasBeenWon = CheckXOrOHasWon(Board)
                NoOfMoves = NoOfMoves + 1
                If Not GameHasBeenWon Then
                    If NoOfMoves = 9 Then 'Check if maximum number of allowed moves has been reached
                        GameHasBeenDrawn = True
                    Else
                        If CurrentSymbol = "X" Then
                            CurrentSymbol = "O"
                        Else
                            CurrentSymbol = "X"
                        End If
                    End If
                End If
            Loop Until GameHasBeenWon Or GameHasBeenDrawn
            If GameHasBeenWon Then ' Update scores and display result
                If PlayerOneSymbol = CurrentSymbol Then
                    Console.WriteLine(PlayerOneName & " congratulations you win!")
                    PlayerOneScore = PlayerOneScore + 1
                Else
                    Console.WriteLine(PlayerTwoName & " congratulations you win!")
                    PlayerTwoScore = PlayerTwoScore + 1
                End If
            Else
                Console.WriteLine("A draw this time!")
            End If
            Console.WriteLine()
            Console.WriteLine(PlayerOneName & " your score is: " & PlayerOneScore)
            Console.WriteLine(PlayerTwoName & " your score is: " & PlayerTwoScore)
            Console.WriteLine()
            If StartSymbol = PlayerOneSymbol Then
                StartSymbol = PlayerTwoSymbol
            Else
                StartSymbol = PlayerOneSymbol
            End If
            Console.Write("Another game Y/N?")
            Answer = Console.ReadLine()
        Loop Until Answer = "N" Or Answer = "n"
    End Sub

    Sub DisplayBoard(ByVal Board(,) As Char)
        Dim Row As Integer
        Dim Column As Integer
        Console.WriteLine("  | 1 2 3 ")
        Console.WriteLine("--+-------")
        For Row = 1 To 3
            Console.Write(Row & " | ")
            For Column = 1 To 3
                Console.Write(Board(Column, Row) & " ")
            Next
            Console.WriteLine()
        Next
        Console.WriteLine()
    End Sub

    Sub ClearBoard(ByRef Board(,) As Char)
        Dim Row As Integer
        Dim Column As Integer
        For Row = 1 To 3
            For Column = 1 To 3
                Board(Column, Row) = " "
            Next
        Next
    End Sub

    Sub GetMoveCoordinates(ByRef XCoordinate As Integer, ByRef YCoordinate As Integer)
        Try
            Console.Write("Enter x coordinate: ")
            XCoordinate = Console.ReadLine()
            Console.Write("Enter y coordinate: ")
            YCoordinate = Console.ReadLine()
            Console.WriteLine()
        Catch ex As Exception
            Console.WriteLine()
            Console.WriteLine("Coordinates invalid, please try again")
            GetMoveCoordinates(XCoordinate, YCoordinate)
        End Try
    End Sub

    Function CheckValidMove(ByVal XCoordinate As Integer, ByVal YCoordinate As Integer, _
ByVal Board(,) As Char)
        Dim ValidMove As Boolean
        ValidMove = True

        'Check x coordinate is valid
        If XCoordinate < 1 Or XCoordinate > 3 Then
            ValidMove = False
            'Check y coordinate is valid
        ElseIf YCoordinate < 1 Or YCoordinate > 3 Then
            ValidMove = False
        End If
        CheckValidMove = ValidMove
    End Function

    Function IsEmpty(ByVal XCoordinate As Integer, ByVal YCoordinate As Integer, _
ByVal Board(,) As Char)
        Dim ValidMove As Boolean
        ValidMove = True
        'Check coordinate is already marked
        If Board(XCoordinate, YCoordinate) <> " " Then
            ValidMove = False
        End If
        IsEmpty = ValidMove
    End Function

    Function CheckXOrOHasWon(ByVal Board(,) As Char) As Boolean
        Dim Row As Integer
        Dim Column As Integer
        Dim XOrOHasWon As Boolean
        XOrOHasWon = False
        For Column = 1 To 3
            If Board(Column, 1) = Board(Column, 2) And Board(Column, 2) = Board(Column, 3) _
                                    And Board(Column, 2) <> " " Then XOrOHasWon = True
        Next
        For Row = 1 To 3
            If Board(1, Row) = Board(2, Row) And Board(2, Row) = Board(3, Row) _
                            And Board(2, Row) <> " " Then XOrOHasWon = True
        Next
        If Board(1, 1) = Board(2, 2) And Board(2, 2) = Board(3, 3) _
                        And Board(2, 2) <> " " Then XOrOHasWon = True
        If Board(1, 3) = Board(2, 2) And Board(2, 2) = Board(3, 1) _
                        And Board(2, 2) <> " " Then XOrOHasWon = True
        CheckXOrOHasWon = XOrOHasWon
    End Function

    Function GetWhoStarts() As Char
        Dim RandomNo As Integer
        Dim WhoStarts As Char
        RandomNo = Rnd() * 100
        If RandomNo Mod 2 = 0 Then
            WhoStarts = "X"
        Else
            WhoStarts = "O"
        End If
        GetWhoStarts = WhoStarts
    End Function
End Module

