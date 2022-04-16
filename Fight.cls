VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Instance of a Fight in the tournament

'Fully justified player image profile names
Dim Player1 As String
Dim Player2 As String

'Bracket label (A, B, C, D, Semi-Final, Final)
Dim groupLabel As String


'Returns the name of the player based on whether this is the first fighter (1) or second fighter (2)
Function GetPlayer(playerNum As Integer) As String

    If playerNum = 1 Then
        GetPlayer = Player1
    ElseIf playerNum = 2 Then
        GetPlayer = Player2
    Else
        MsgBox "Invalid. Player # must be 1 or 2"
    End If

End Function

'Sets the name of the players based on the Round 1 Fight ID
'Names of the player are the fully justified file paths to the player's image files
'e.g. The first two fighters in round 1 are 1a and 1b. The second two fighters in round 1 are 2a and 2b.
Sub BuildPlayers(fightID As Integer, filePath As String)
    
    Player1 = filePath & fightID & "a.png"
    Player2 = filePath & fightID & "b.png"

End Sub

'Sets the name of the players based on the winner filename that is provided and which player
Sub SetPlayer(playerName As String, playerNum As Integer)
    If playerNum = 1 Then
        Player1 = playerName
    ElseIf playerNum = 2 Then
        Player2 = playerName
    Else
        MsgBox "Invalid. Player # must be 1 or 2"
    End If

End Sub

'Sets the braket label based on the number of fighters and the round
'A = 1, B = 2, C = 3, D = 4, SemiFinals - 1, Finals = 0
Sub SetGroupLabel(groupID As Integer)

    'Map the groupID to the final display version of the group label
    Select Case groupID
        Case -1
            groupLabel = "SEMIFINALS"
        Case 0
             groupLabel = "FINALS"
        Case 1
            groupLabel = "BLOCK A"
        Case 2
            groupLabel = "BLOCK B"
        Case 3
            groupLabel = "BLOCK C"
        Case 4
            groupLabel = "BLOCK D"
        Case Else
            MsgBox groupdID & " is invalid group ID"
    End Select
    
    
End Sub

'Returns the current bracket label
Function GetGroupLabel() As String
    
    GetGroupLabel = groupLabel

End Function


