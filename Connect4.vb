Imports System
Imports System.Collections.Generic
Imports System.Text

Namespace Connect4Graphical

    Public Class Computer
        Implements IPlayer

        Private firstLevel As Integer = 0


        Public CutLevel As Integer = 7
        Public Weights As Integer() = New Integer() {1, 5, 100, 10000, 2, 6, 200, 15000}




#Region "Search algorithm based On MiniMax With Alpha Beta Pruning"

        Private Function AlphaBetaSearch(state As StateType) As Integer

            state.Succesors = Successors(state) 'Retrieve successors for current state.
            firstLevel = state.MoveCount 'Remember start state for cut-off function
            Dim Recursion As Integer = MaxValue(state, Integer.MinValue, Integer.MaxValue) 'Start recursive traversel Of statespace
            'Pick which successor is the correct recursion
            For Each a As Integer In state.Succesors.Keys
                If (state.Succesors(a).V = Recursion) Then
                    Return a
                End If
            Next
            Return -1
        End Function

        Private Function MaxValue(state As StateType, Alpha As Integer, Beta As Integer) As Integer

            If (CutOffTest(state)) Then Return state.V  'Check For Cutoff.
            state.V = Integer.MinValue
            Dim succ As New Dictionary(Of Integer, StateType)

            ' If (Not state.Succesors Is Nothing) Then
            If (state.Succesors.Count <> 0) Then '818
                succ = state.Succesors
            Else
                succ = Successors(state)
            End If
            For Each a As Integer In succ.Keys

                state.V = Math.Max(state.V, MinValue(succ(a), Alpha, Beta))
                If (state.V >= Beta) Then
                    Return state.V
                End If
                Alpha = Math.Max(Alpha, state.V)
            Next
            Return state.V
        End Function

        Private Function MinValue(state As StateType, Alpha As Integer, Beta As Integer) As Integer

            If (CutOffTest(state)) Then Return Eval(state.Values)
            state.V = Integer.MaxValue
            Dim succ As New Dictionary(Of Integer, StateType)
            ' If (Not state.Succesors Is Nothing) Then
            If (state.Succesors.Count <> 0) Then '818
                succ = state.Succesors
            Else
                succ = Successors(state)
            End If

            For Each a As Integer In succ.Keys

                state.V = Math.Min(state.V, MaxValue(succ(a), Alpha, Beta))
                If (state.V <= Alpha) Then
                    Return state.V
                End If
                Beta = Math.Min(Beta, state.V)
            Next
            Return state.V
        End Function

        Private Function Successors(state As StateType) As Dictionary(Of Integer, StateType)
            Dim succ As New Dictionary(Of Integer, StateType)(7)
            For i As Integer = 1 To 7

                Dim st As StateType = state.Move(i) 'If move Is valid
                If (Not st Is Nothing) Then succ.Add(i, st)
            Next
            Return succ
        End Function

        Private Function CutOffTest(state As StateType) As Boolean

            state.V = Eval(state.Values) 'Check for Win situation
            If (Math.Abs(state.V) > 5000) Then Return True 'We have a terminal state
            If ((state.MoveCount - firstLevel) > CutLevel) Then Return True
            Return False
        End Function


#End Region

#Region "Heuristic Evaluation"

        'Private int CheckLine(params int[] vals)
        Private Function CheckLine(ParamArray vals() As Integer) As Integer

            Dim score As Integer = 0
            For i As Integer = 0 To vals.Length - 4

                'Examine each opportunity
                Dim ComputerPieces As Integer = 0
                Dim PlayerPieces As Integer = 0
                For j As Integer = 0 To 3
                    If (vals(i + j) = number) Then
                        ComputerPieces += 1 'TODO Make sure that it looks at it's own identity
                    Else
                        If (vals(i + j) <> 0) Then PlayerPieces += 1
                    End If
                Next

                If ((ComputerPieces > 0) And (PlayerPieces = 0)) Then
                    'Computer opportunity
                    If (ComputerPieces = 4) Then Return Weights(3) 'Win
                    score += ((ComputerPieces / 3) * Weights(2)) + ((ComputerPieces / 2) * Weights(1)) + Weights(0)

                ElseIf ((ComputerPieces = 0) And (PlayerPieces > 0)) Then
                    'Player opportunity
                    If (PlayerPieces = 4) Then Return -1 * Weights(7) 'Win
                    score -= ((PlayerPieces / 3) * Weights(6)) + ((PlayerPieces / 2) * Weights(5)) + Weights(4)
                End If
            Next
            Return score
        End Function

        Private Function Eval(state(,) As Integer) As Integer


            Dim score As Integer = 0
            'Eval Horizontal
            For i As Integer = 0 To 5
                score += CheckLine(state(0, i), state(1, i), state(2, i), state(3, i), state(4, i), state(5, i), state(6, i))
            Next

            'Eval Vertical
            For i As Integer = 0 To 6
                score += CheckLine(state(i, 0), state(i, 1), state(i, 2), state(i, 3), state(i, 4), state(i, 5))
            Next

            'Eval Diagonal /
            score += CheckLine(state(0, 2), state(1, 3), state(2, 4), state(3, 5))
            score += CheckLine(state(0, 1), state(1, 2), state(2, 3), state(3, 4), state(4, 5))
            score += CheckLine(state(0, 0), state(1, 1), state(2, 2), state(3, 3), state(4, 4), state(5, 5))
            score += CheckLine(state(1, 0), state(2, 1), state(3, 2), state(4, 3), state(5, 4), state(6, 5))
            score += CheckLine(state(2, 0), state(3, 1), state(4, 2), state(5, 3), state(6, 4))
            score += CheckLine(state(3, 0), state(4, 1), state(5, 2), state(6, 3))

            'Eval Diagonal \
            score += CheckLine(state(3, 0), state(2, 1), state(1, 2), state(0, 3))
            score += CheckLine(state(4, 0), state(3, 1), state(2, 2), state(1, 3), state(0, 4))
            score += CheckLine(state(5, 0), state(4, 1), state(3, 2), state(2, 3), state(1, 4), state(0, 5))
            score += CheckLine(state(6, 0), state(5, 1), state(4, 2), state(3, 3), state(2, 4), state(1, 5))
            score += CheckLine(state(6, 1), state(5, 2), state(4, 3), state(3, 4), state(2, 5))
            score += CheckLine(state(6, 2), state(5, 3), state(4, 4), state(3, 5))
            Return score
        End Function

#End Region

        Public Function PlayerTurn(state As StateType) As Integer Implements IPlayer.PlayerTurn
            Return AlphaBetaSearch(state) 'Perform search To find best action
        End Function

        Private number As Integer
        Public Sub New(PlayerNumber As Integer) 'Computer(PlayerNumber As Integer)
            number = PlayerNumber
        End Sub


    End Class

End Namespace

Imports System
Imports System.Collections.Generic
Imports System.Text

Namespace Connect4Graphical

    Public Class Human
        Implements IPlayer



        Public Function PlayerTurn(state As StateType) As Integer Implements IPlayer.PlayerTurn

            Game.ClearImages()

            Game.DrawImages(state)




            Do Until Game.DontWait
                Application.DoEvents()
                Trace.WriteLine("Waiting for player input")
            Loop
            Game.FrmMain.DontWait = False
            Dim CheckedColumn As Integer = Game.FrmMain.ButtonIndex
            Return CheckedColumn
        End Function

        Public Sub New()

        End Sub

    End Class
End Namespace

Imports System
Imports System.Collections.Generic
Imports System.Text

Namespace Connect4Graphical

    Public Interface IPlayer
        Function PlayerTurn(state As StateType) As Integer
    End Interface
End Namespace


Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Threading

Namespace Connect4Graphical

    Public Class Connect4

        Public State As StateType 'Current game state
        Public Winner As Boolean  'Has a winner been found

        Public Player1 As IPlayer 'Points To Player 1
        Public Player2 As IPlayer 'Points To Player 2

        Public Function Turn(Move As Integer) As Boolean

            Dim st As StateType = State.Move(Move)
            If (st Is Nothing) Then Return False
            State = st
            Return True
        End Function

        Public Sub PlayGame()

            Dim ww As Integer = 0    'ww=0 => no winner yet, ww=1 => Player 1 wins, ww=2 => Player 2 wins, ww=-1 => Tie
            Dim WinnerMenuResult As DialogResult



            Do While Not Winner

                UpdateInterfaceInformation(3, Nothing)


                If (State.Turn = 1) Then
                    Turn(Player1.PlayerTurn(State))
                Else
                    Turn(Player2.PlayerTurn(State)) ' No Error handling yet
                End If

                UpdateInterfaceInformation(1, Nothing)

                Game.Refresh()

                ww = State.CheckWin() 'ww=0 => no winner yet, ww=1 => Player 1 wins, ww=2 => Player 2 wins, ww=-1 => Tie
                Winner = (ww <> 0)


                Game.ClearImages()

                Game.DrawImages(State)

                Game.Refresh()

                Game.PushSaveGame(State)


            Loop

            Game.timer1ms.Stop()  'ending timer so it displays full game time
            Game.totalTimeStopwatch.Stop()

            Game.ClearImages()

            Game.FrmMain.DrawImages(State)

            UpdateInterfaceInformation(2, ww)

            If (ww = -1) Then

                WinnerMenuResult = MsgBox("Its a Tie!!  Would you like to play again?", MessageBoxButtons.YesNo, "Winner : Draw")
            Else

                If (ww = 1) Then

                    WinnerMenuResult = MsgBox("Player 1 Wins!  Would you like to play again?", MessageBoxButtons.YesNo, "Winner : Player 1")

                Else

                    WinnerMenuResult = MsgBox("Player 2 Wins!  Would you like to play again?", MessageBoxButtons.YesNo, "Winner : Player 2")

                End If
            End If

            If WinnerMenuResult = DialogResult.Yes Then
                Game.Close()
                NewGame.Show()
            End If

        End Sub

        Public Sub New() ' Connect4()

            'Initialize
            Winner = False
            State = New StateType()
            State.Turn = 1
        End Sub

        Private Sub UpdateInterfaceInformation(value As Integer, winner As Integer) 'Changing labels to reflect game
            If value = 1 Then
                If State.MoveCount = 1 Then
                    Game.totalTimeStopwatch.Start()       'starting game timings 
                    Game.currentTurnTimeStopwatch.Start()
                    Game.timer1ms.Start()
                Else
                    Game.currentTurnTimeStopwatch.Restart()
                End If

                Dim whosTurn As String = ""
                If State.Turn = 1 Then
                    whosTurn = "Player1 Turn"
                ElseIf State.Turn = 2 Then
                    whosTurn = "Player2 Turn"
                End If

                Game.lbl_gameStatusDisplay.Text = "Thinking..."
                Game.lbl_moveCountDisplay.Text = State.MoveCount
                Game.lbl_whosTurnDisplay.Text = whosTurn
            ElseIf value = 2 Then
                Dim gameStatus As String = ""
                If winner = 1 Then
                    gameStatus = "Player 1 Wins"
                ElseIf winner = 2 Then
                    gameStatus = "Player 2 wins"
                End If
                Game.lbl_gameStatusDisplay.Text = gameStatus
            ElseIf value = 3 Then
                Game.lbl_gameStatusDisplay.Text = "Waiting on you"
            End If

        End Sub
    End Class
End Namespace


Imports System
Imports System.Collections.Generic
Imports System.Text

Namespace Connect4Graphical

    Class Program

        Shared Sub PlayervComputer(StartingState As StateType)


            Dim c4 As New Connect4()
            If Not StartingState Is Nothing Then
                c4.State = StartingState
            End If

            c4.Player1 = New Human()
            c4.Player2 = New Computer(2)

            c4.PlayGame()

        End Sub

        Shared Sub PlayervPlayer(StartingState As StateType)

            Dim c4 As New Connect4
            If Not StartingState Is Nothing Then
                c4.State = StartingState
            End If
            c4.Player1 = New Human
            c4.Player2 = New Human

            c4.PlayGame()
        End Sub

        Shared Sub ComputervComputer(StartingState As StateType)

            Dim c4 As New Connect4

            c4.Player1 = New Computer(1)
            c4.Player2 = New Computer(2)

            TryCast(c4.Player2, Computer).CutLevel = 7
            TryCast(c4.Player2, Computer).Weights = New Integer() {1, 5, 100, 10000, 1, 5, 200, 15000}

            c4.PlayGame()

        End Sub
    End Class
End Namespace



Imports System
Imports System.Collections.Generic
Imports System.Text

Namespace Connect4Graphical

    Public Class StateType

        Public Values(,) As Integer 'Board: horizontal, vertical

        Public RowCounts() As Integer  'Number Of occupied fields In Each column.
        Public Turn As Integer  'Current Player
        Public V As Integer 'State Value
        Public MoveCount As Integer  'Moves made
        Public Succesors As New Dictionary(Of Integer, StateType) 'Cache Action-Successors pairs storage For first node.
        Public GameType As Integer        'Type of game being played


        Public Function Move(MM As Integer) As StateType

            Dim st As StateType = New StateType(Me) 'Clone current state
            Dim m As Integer = MM - 1
            If (st.RowCounts(m) = 6) Then Return Nothing 'If column Is full, return null. Illegal action
            st.RowCounts(m) += 1
            st.Values(m, (6 - st.RowCounts(m))) = st.Turn 'Update board



            If (st.Turn = 1) Then 'Change turn
                st.Turn = 2
            Else
                st.Turn = 1
            End If

            st.MoveCount += 1
            Return st
        End Function


#Region "Check For Win situation"

        Private Function CheckLine(ParamArray vals() As Integer) As Integer

            Dim last As Integer = 0
            Dim lastcount As Integer = 0
            For i As Integer = 0 To vals.Length - 1


                If (vals(i) = last) Then
                    lastcount += 1
                Else
                    last = vals(i)
                    lastcount = 1
                End If

                If ((lastcount = 4) And (last > 0)) Then Return last
            Next
            Return 0
        End Function


        Public Function CheckWin() As Integer

            Dim r As Integer = 0
            'Check Horizontal
            For i As Integer = 0 To 5
                r = CheckLine(Values(0, i), Values(1, i), Values(2, i), Values(3, i), Values(4, i), Values(5, i), Values(6, i))
                If (r <> 0) Then Return r
            Next

            'Check Vertical
            For i As Integer = 0 To 6
                r = CheckLine(Values(i, 0), Values(i, 1), Values(i, 2), Values(i, 3), Values(i, 4), Values(i, 5))
                If (r <> 0) Then Return r
            Next

            'Check Diagonal
            r = CheckLine(Values(0, 2), Values(1, 3), Values(2, 4), Values(3, 5))
            If (r <> 0) Then Return r
            r = CheckLine(Values(0, 1), Values(1, 2), Values(2, 3), Values(3, 4), Values(4, 5))
            If (r <> 0) Then Return r
            r = CheckLine(Values(0, 0), Values(1, 1), Values(2, 2), Values(3, 3), Values(4, 4), Values(5, 5))
            If (r <> 0) Then Return r
            r = CheckLine(Values(1, 0), Values(2, 1), Values(3, 2), Values(4, 3), Values(5, 4), Values(6, 5))
            If (r <> 0) Then Return r
            r = CheckLine(Values(2, 0), Values(3, 1), Values(4, 2), Values(5, 3), Values(6, 4))
            If (r <> 0) Then Return r
            r = CheckLine(Values(3, 0), Values(4, 1), Values(5, 2), Values(6, 3))
            If (r <> 0) Then Return r
            'Check Diagonal 2
            r = CheckLine(Values(3, 0), Values(2, 1), Values(1, 2), Values(0, 3))
            If (r <> 0) Then Return r
            r = CheckLine(Values(4, 0), Values(3, 1), Values(2, 2), Values(1, 3), Values(0, 4))
            If (r <> 0) Then Return r
            r = CheckLine(Values(5, 0), Values(4, 1), Values(3, 2), Values(2, 3), Values(1, 4), Values(0, 5))
            If (r <> 0) Then Return r
            r = CheckLine(Values(6, 0), Values(5, 1), Values(4, 2), Values(3, 3), Values(2, 4), Values(1, 5))
            If (r <> 0) Then Return r
            r = CheckLine(Values(6, 1), Values(5, 2), Values(4, 3), Values(3, 4), Values(2, 5))
            If (r <> 0) Then Return r
            r = CheckLine(Values(6, 2), Values(5, 3), Values(4, 4), Values(3, 5))
            If (r <> 0) Then Return r

            'Check for full (=tie)
            Dim empty As Boolean = False
            For i As Integer = 0 To RowCounts.Length - 1
                If (RowCounts(i) < 6) Then
                    empty = True
                    Exit For
                End If
            Next

            If (Not empty) Then Return -1
            Return 0
        End Function
#End Region

        Public Overrides Function ToString() As String

            Dim sb As New StringBuilder()
            'sb.Append("\t1\t2\t3\t4\t5\t6\t7" & vbCrLf & vbCrLf)
            sb.Append(Chr(9) & "1" & Chr(9) & "2" & Chr(9) & "3" & Chr(9) & "4" & Chr(9) & "5" & Chr(9) & "6" & Chr(9) & "7" & vbCrLf & vbCrLf)
            For j As Integer = 0 To 5

                For i As Integer = 0 To 6
                    Dim s As String = ""
                    Dim s1 As String

                    If (Values(i, j) = 1) Then
                        s1 = "1"
                    Else
                        s1 = "2"
                    End If

                    If (Values(i, j) = 0) Then
                        s = "*"
                    Else
                        s = s1
                    End If


                    sb.Append(Chr(9) & s)
                Next
                'sb.Append("\n\n")
                sb.Append(vbCrLf & vbCrLf)
            Next
            'sb.Append("\n\nMoves: " + MoveCount.ToString() + "\nNext Turn: Player " + Turn.ToString() + "\n")
            sb.Append(vbCrLf & vbCrLf & "Moves: " & MoveCount.ToString() & vbCrLf & "Next Turn: Player " & Turn.ToString() & vbCrLf)

            Return sb.ToString()
        End Function


        Public Sub New(orig As StateType)

            Values = CType(orig.Values.Clone(), Integer(,))
            RowCounts = CType(orig.RowCounts.Clone(), Integer())
            Turn = orig.Turn
            MoveCount = orig.MoveCount
        End Sub

        Public Sub New()
            ReDim Values(7, 6)
            ReDim RowCounts(7)
            MoveCount = 0
        End Sub
    End Class
End Namespace


Imports Connect4Graphical.Connect4Graphical
Imports System.IO
Public Class LoadGame

    Private LoadFileState As StateType

    Private Sub LoadGame_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim OpenGame As New OpenFileDialog
        OpenGame.Filter = "Text Files|*.txt"
        OpenGame.InitialDirectory = "C:"
        OpenGame.Title = "Select a game file to Open"

        If OpenGame.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim FileName As String = OpenGame.FileName
            Dim FileContents As List(Of Integer) = LoadGameFile(FileName)
            LoadFileState = OrganiseData(FileContents)

            txt_fileNameDisplay.Text = OpenGame.FileName.ToString          '}
            lbl_moveCountDisplay.Text = DisplayMoveCount(LoadFileState)    '}visual
            lbl_currentTurnDisplay.Text = DisplayTurn(LoadFileState)       '}labels
            lbl_gametypeDisplay.Text = DisplayGametype(LoadFileState)      '}

            txt_gameBoard.Text = DisplayBoard(LoadFileState) 'visual board

        End If
    End Sub

    Public Function DisplayMoveCount(state As StateType) As String
        Return state.MoveCount
    End Function

    Public Function DisplayTurn(state As StateType) As String
        Dim Turn As String
        If state.Turn = 1 Then
            Turn = "Player 1's Turn"
        Else
            Turn = "Player 2's Turn"
        End If
        Return Turn
    End Function

    Public Function DisplayGametype(state As StateType) As String
        Dim GameType As String
        If state.GameType = 1 Then
            GameType = "Player v Player"
        Else
            GameType = "Player v Computer"
        End If
        Return GameType
    End Function

    Public Function DisplayBoard(State As StateType) As String
        Dim Str As String = ""
        For j As Integer = 0 To 5
            For i As Integer = 0 To 6

                Str &= State.Values(i, j) & " "
            Next
            Str &= vbCrLf
        Next
        Return Str
    End Function

    Public Function ConstructBoard(RawData As List(Of Integer)) As Integer(,)  'makes board ready for loading
        Dim Board(6, 5) As Integer
        Dim k As Integer = 0

        For i = 0 To 6
            For j = 0 To 5
                Board(i, j) = RawData(k)
                k += 1
            Next
        Next
        Return Board
    End Function

    Private Function ConstructRowCounts(RawData As List(Of Integer)) As Integer()   'makes rowcounts ready for loading
        Dim RowCounts(6) As Integer
        For i As Integer = 0 To 6
            RowCounts(i) = RawData(42 + i)
        Next
        Return RowCounts
    End Function


    Private Function OrganiseData(RawData As List(Of Integer)) As StateType        'creates the statetype to be loaded from
        Dim Board(,) As Integer
        Dim RowCounts() As Integer
        Dim MoveCount As Integer
        Dim Turn As Integer
        Dim GameType As Integer

        Board = ConstructBoard(RawData)
        RowCounts = ConstructRowCounts(RawData)
        MoveCount = RawData(RawData.Count - 3)
        Turn = RawData(RawData.Count - 2)
        GameType = RawData(RawData.Count - 1)

        'validate()

        Dim st As New StateType
        For i = 0 To 6
            For j = 0 To 5
                st.Values(i, j) = Board(i, j)
            Next
        Next
        For i = 0 To 6
            st.RowCounts(i) = RowCounts(i)
        Next
        st.MoveCount = MoveCount
        st.Turn = Turn
        st.GameType = GameType

        Return st
    End Function

    Private Function LoadGameFile(FileName As String) As List(Of Integer)
        If Not FileName Is Nothing Then
            Dim fStream As New FileStream(FileName, FileMode.Open)   'opens file to read
            Dim sReader As New StreamReader(fStream)

            Dim FileList As New List(Of Integer)         'writes contents of file to a list
            Do While sReader.Peek >= 0
                FileList.Add(sReader.ReadLine)
            Loop


            fStream.Close()          'closes file
            sReader.Close()

            Return FileList
        End If
        Return Nothing
    End Function

    Private Sub btn_beginLoad_Click(sender As Object, e As EventArgs) Handles btn_beginLoad.Click
        Game.Show()
        If LoadFileState.GameType = 1 Then
            Program.PlayervPlayer(LoadFileState)  'gametype = pvp
        Else
            Program.PlayervComputer(LoadFileState) 'gametype = pvc
        End If
        Me.Close()
    End Sub
End Class


Imports Connect4Graphical.Connect4Graphical

Public Class NewGame

    Private Sub btn_playervPlayer_Click(sender As Object, e As EventArgs) Handles btn_playervPlayer.Click
        Game.Show()
        Me.Close()
        Game.PushSaveGamemode(1)
        Program.PlayervPlayer(Nothing)
    End Sub

    Private Sub btn_playervComputer_Click(sender As Object, e As EventArgs) Handles btn_playervComputer.Click
        Game.Show()
        Me.Close()
        Game.PushSaveGamemode(2)
        Program.PlayervComputer(Nothing)
    End Sub

    Private Sub btn_computervComputer_Click(sender As Object, e As EventArgs) Handles btn_computervComputer.Click
        Game.Show()
        Me.Close()
        Program.ComputervComputer(Nothing)
    End Sub

    Private Sub btn_goBack_Click(sender As Object, e As EventArgs) Handles btn_goBack.Click
        Me.Close()
        MainMenu.Show()
    End Sub

    Private Sub btn_exit_Click(sender As Object, e As EventArgs) Handles btn_exit.Click
        Me.Close()
    End Sub
End Class


Imports Connect4Graphical
Imports Connect4Graphical.Connect4Graphical

Public Class Game
    Public Shared FrmMain As Game
    Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        FrmMain = Me
        For Each lbl As Control In gbox_gameplayingarea.Controls
            If TypeOf lbl Is Label And lbl.Name.Contains("A") = False And lbl.Name.Contains("B") = False And lbl.Name.Contains("Num") = False Then
                Dim name As String = lbl.Name
                Dim index As Double = name.Replace("Label", "").Trim
                Dim row As Integer = -1

                For i As Integer = 0 To 5
                    If (index / 7 > i And index / 7 <= i + 1) Then row = i
                Next

                Dim column As Integer = -1

                column = (7 * (row + 1) - index) Mod 7
                lbl.Tag = row & "--" & column

            End If
        Next

    End Sub

    Friend Sub DrawImages(state As StateType)
        Dim sep(0) As String
        sep(0) = vbCrLf & vbCrLf
        Dim lines() As String = state.ToString.Split(sep, StringSplitOptions.RemoveEmptyEntries)
        For i As Integer = 1 To 6
            Dim line As String = lines(i)
            Dim row As Integer = 6 - i

            Dim sepp(0) As String
            sepp(0) = Chr(9)
            Dim column As Integer = -1
            Dim cells() As String = line.Split(sepp, StringSplitOptions.RemoveEmptyEntries)
            For j = 0 To 6
                Dim CellValue As String = cells(j)
                column = j
                SetLabelImage(row, column, CellValue)
            Next
            ' Debugger.Break()
        Next
    End Sub

    Private Sub SetLabelImage(row As Integer, column As Integer, cellValue As String)
        Dim label As Label
        For Each lbl As Control In gbox_gameplayingarea.Controls
            If TypeOf lbl Is Label Then
                If lbl.Tag = row & "--" & column Then
                    label = lbl
                    Exit For
                End If
            End If
        Next

        Select Case cellValue
            Case "1" : label.Image = lblA.Image ' label.BackColor = Color.Red
            Case "2" : label.Image = lblB.Image  'label.BackColor = Color.Yellow
            Case "*" : label.Image = Nothing 'label.BackColor = Color.White
        End Select
    End Sub

    Public Sub ClearImages()
        For Each lbl As Control In gbox_gameplayingarea.Controls
            If TypeOf lbl Is Label Then
                If lbl.Tag <> "" Then CType(lbl, Label).Image = Nothing
            End If
        Next
    End Sub

    Public ButtonIndex As Integer
    Public Running As Boolean
    Public DontWait As Boolean = False

    Private Sub btn_dropcolumn_Click(sender As Object, e As EventArgs) Handles btn_dropcolumn1.Click, btn_dropcolumn2.Click, btn_dropcolumn3.Click, btn_dropcolumn4.Click, btn_dropcolumn5.Click, btn_dropcolumn6.Click, btn_dropcolumn7.Click
        Dim Value As Integer = Val(CType(sender, Button).Text)
        DontWait = True
        ButtonIndex = Value
    End Sub


    Public totalTimeStopwatch As New Stopwatch
    Public currentTurnTimeStopwatch As New Stopwatch


    Private Sub totalGameTime_Tick(sender As Object, e As EventArgs) Handles timer1ms.Tick
        Dim elapsed As TimeSpan = Me.totalTimeStopwatch.Elapsed
        lbl_totalGameTimeDisplay.Text = String.Format("{0:00}:{1:00}:{2:00}",
                                                 Math.Floor(elapsed.TotalHours),
                                                elapsed.Minutes, elapsed.Seconds)
    End Sub

    Private Sub currentTurnTime_Tick(sender As Object, e As EventArgs) Handles timer1ms.Tick
        Dim elapsed As TimeSpan = Me.currentTurnTimeStopwatch.Elapsed
        lbl_currentTurnTimeDisplay.Text = String.Format("{0:00}:{1:00}:{2:00}",
                                                 Math.Floor(elapsed.TotalHours),
                                                  elapsed.Minutes, elapsed.Seconds)
    End Sub

#Region " Save Game "
    Private CurrentState As StateType     'clone of game state
    Private CurrentGameType As Integer

    Public Sub PushSaveGame(State As StateType)
        CurrentState = State
    End Sub

    Public Sub PushSaveGamemode(gamemode As Integer)
        CurrentGameType = gamemode        '1 = playervplayer   2=playervcomputer
    End Sub

    Private Sub btn_saveGameAs_Click(sender As Object, e As EventArgs) Handles btn_saveGameAs.Click
        totalTimeStopwatch.Stop()              'stop timers for the game
        currentTurnTimeStopwatch.Stop()
        Dim SaveFile As New SaveFileDialog
        SaveFile.Filter = "Text File|*.txt"
        SaveFile.Title = "Save Game"
        SaveFile.ShowDialog()

        If SaveFile.FileName <> "" Then
            FileOpen(1, SaveFile.FileName, OpenMode.Output)       'Create and open file
            For i As Integer = 0 To 6
                For j As Integer = 0 To 5
                    WriteLine(1, CurrentState.Values(i, j))       'write board to file
                Next
            Next
            For i As Integer = 0 To CurrentState.RowCounts.Length - 1  'write rowcounts to file
                WriteLine(1, CurrentState.RowCounts(i))
            Next
            WriteLine(1, CurrentState.MoveCount)
            WriteLine(1, CurrentState.Turn)
            WriteLine(1, CurrentGameType)

            FileClose(1)        'close file 
        End If
    End Sub

#End Region

#Region " Load Game "

#End Region
End Class


Public Class MainMenu


    Private Sub btn_newGame_Click(sender As Object, e As EventArgs) Handles btn_newGame.Click
        NewGame.Show()
        Me.Close()
    End Sub

    Private Sub btn_loadGame_Click(sender As Object, e As EventArgs) Handles btn_loadGame.Click
        LoadGame.Show()
    End Sub

    Private Sub btn_howToPlay_Click(sender As Object, e As EventArgs) Handles btn_howToPlay.Click
        Help.Show()
    End Sub

    Private Sub btn_about_Click(sender As Object, e As EventArgs) Handles btn_about.Click
        About.Show()
    End Sub

    Private Sub btn_exit_Click(sender As Object, e As EventArgs) Handles btn_exit.Click
        Me.Close()
    End Sub
End Class




Public Class MainTaskbar


#Region "New / Load / Save as / Exit"
    Private Sub NewGameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewGameToolStripMenuItem.Click
        NewGame.Show()
        Game.Close()
    End Sub

    Private Sub LoadGameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadGameToolStripMenuItem.Click
        LoadGame.Show() b
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        MainMenu.Show()
        Game.Close()
    End Sub

#End Region
#Region "How to play / About"

    Private Sub HowToPlayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HowToPlayToolStripMenuItem.Click
        Help.Show()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        About.Show()
    End Sub

#End Region

End Class

