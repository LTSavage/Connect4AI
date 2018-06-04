Function PlayGame

PWinner = 0

Loop While Winner = False
    If Turn =1
        Player1 Turn
    Else
        Player2 Turn
    End If
    
    UpdateGame
    
    CheckWin
    If Winner = True
        WinnerFound
    End If
    
End Loop



Function SetUpBoard

For each Label
    For 0 to 5
        For 0 to 6
        Label = "i,j"
        End For
    End For
End For

Function Move

State = Current State
If State.M = 6
    Return Nothing
End
Increment State.M

If Turn = 1
    Turn = 2
Else
    Turn = 1
End If

Increment Move Count

Return State


Function CheckWin
Score=0
for i=0 to 5
    R=CheckLine(row)
End
for i=0 to 6
    R=CheckLine(Column)
End
for i=0 to 4
    R=CheckLine(diagonal1)
End
for i=0 to 4
    R=CheckLine(diagonal2)
End
If r is not 0
    Return 0
End
If Board is Full 
    Return -1
End
Return 0



Function CheckLine(Vals)

Last = 0
LCount = 0
For i = 0 to Vals.Length
    If Vals = last 
        Increment LCount
    Else
        Last = Vals
        LCount = 1
    End
    If LCount = 4 and Last > 0
        Return Last
    End
End
Return 0



Function CheckLine

Score = 0
For i = 0 to Length
    P = 0
    C = 0
    For j = 0 to Length
    If PlayerPiece
        Increment P
    Else
        Increment C
    End
    End
End

If P > 0 and C = 0
    Decrease Score
Else
    Increase Score
End
Return Score


Function SaveGame
SaveFile as new SaveFileDialog
Open File
Write Board
Write Meta
Close File



Function LoadGame
LoadFile as new OpenFileDialog
Open File
Read Board
Read Meta
Close File
Return (Board & Meta)

