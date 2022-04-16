Attribute VB_Name = "Controller"
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CONTROLLER - handles the flow and logic of the game
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


'GLOBAL VARIABLES
Dim imagefilePath As String
Dim numFighters As Integer

Dim fightCtr As Integer
Dim bracket() As Fight

Dim loadingScreenSlide As Slide
Dim fightSlide As Slide
Dim victorySlide As Slide
Dim seatingSlide As Slide

Dim tournamentWinner As String


'Initializes the tournament: set number of fighters, specify the fighter picture filepaths
'Builds the bracket array and builds the first round of fighters
'Fight counter initialized to 0
Sub InitializeTournament()

    'VALID NUMBERS: 8, 16, 32
    numFighters = 32
    
    'image file path should be in a folder called 'Participants' located in the same directory as this presentation
    imagefilePath = ActivePresentation.Path & "\Participants\"
    
    'Retrieve references to the slides
    Set loadingScreenSlide = ActivePresentation.Slides("LoadingScreen")
    Set fightSlide = ActivePresentation.Slides("Fight")
    Set victorySlide = ActivePresentation.Slides("Victory")
    Set seatingSlide = ActivePresentation.Slides("Seating")
    
    'Make sure there is no text on any of the picture holders
    fightSlide.Shapes("1").TextFrame.TextRange.Text = ""
    fightSlide.Shapes("2").TextFrame.TextRange.Text = ""
    victorySlide.Shapes("Winner").TextFrame.TextRange.Text = ""
    
    
    'Create the bracket for the fighters, where # of fights is 1 less than the total number of fighters
    'VBA arrays specify index, not length (i.e. arr(4) creates an array of 0 to 4 of length 5
    'To make # of fights (numFighters - 1), specified index should be (numFighters - 2)
    ReDim bracket(numFighters - 2)
    
    'Initialize the first round of fighter names, which is 1/2 of the bracket
    'In other rounds, the fighter names will be recored when the victor is decided
    For i = 0 To UBound(bracket)
        
        Set bracket(i) = New Fight
        
        'Initialize the first round of fighter names, which is 1/2 of the bracket
        'Player names will be fully jusitified file path to their profile image
        'In other rounds, the fighter names will be recorded when the victor is decided
        If i <= Floor(UBound(bracket) / 2) Then
            
            bracket(i).BuildPlayers CInt(i + 1), imagefilePath
            
        End If
        
        'Set group labels for the entire tournament
        bracket(i).SetGroupLabel (GetGroupID(numFighters, CInt(i + 1)))
    Next i
    
    
    fightCtr = 0
    
    MsgBox "Tournament Initialization Complete."
End Sub

Sub RunMatch()

    InitializeTournament
    MsgBox "Click anywhere to go to seating slide."
    'ActivePresentation.SlideShowWindow.View.GotoSlide (seatingSlide.SlideIndex)

End Sub


'Advance the fight to the current fight, set up the display for Round and Group
'Go to the loading screen
Sub LoadLoadingScreen()

    fightCtr = fightCtr + 1

    With loadingScreenSlide
        .Shapes("RoundID").TextFrame.TextRange.Text = Utility.GetRoundID(numFighters, fightCtr)
        .Shapes("GroupID").TextFrame.TextRange.Text = bracket(fightCtr - 1).GetGroupLabel
    End With
    
    ActivePresentation.SlideShowWindow.View.GotoSlide (loadingScreenSlide.SlideIndex)

End Sub

'Prepare the display for the fight and go to the fight screen
Sub LoadFight()

    With fightSlide

        'Set the background of the shapes
        .Shapes("1").Fill.UserPicture bracket(fightCtr - 1).GetPlayer(1)
        .Shapes("2").Fill.UserPicture bracket(fightCtr - 1).GetPlayer(2)
        
        'Make sure all shapes on the fight slide are visible
        .Shapes("1").Visible = msoTrue
        .Shapes("2").Visible = msoTrue
        .Shapes("1Cover").Visible = msoTrue
        .Shapes("2Cover").Visible = msoTrue
        
        'Remove any text that might be on the placeholders
        .Shapes("1").TextFrame.TextRange.Text = ""
        .Shapes("2").TextFrame.TextRange.Text = ""
        
        'Set the opacity on the concealer shapes depending on the round
        If Utility.GetRoundID(numFighters, fightCtr) = 1 Then
            .Shapes("1Cover").Fill.Transparency = 0
            .Shapes("2Cover").Fill.Transparency = 0
        Else
            .Shapes("1Cover").Fill.Transparency = 1
            .Shapes("2Cover").Fill.Transparency = 1
        End If
    End With

    ActivePresentation.SlideShowWindow.View.GotoSlide (fightSlide.SlideIndex)

End Sub

'Record the winner (fighter that was clicked) as a player in their next fight
Sub SelectWinner(oShape As Shape)

    'Extract which player this is (1 or 2)
    Dim playerNum As Integer
    playerNum = CInt(oShape.Name)
    
    RecordWinner (playerNum)
    
End Sub

'Records the winner based on the player that was selected
Sub RecordWinner(playerNum As Integer)

    'Extract the name of the player
    Dim playerName As String
    playerName = bracket(fightCtr - 1).GetPlayer(playerNum)

    'If this is the finals, there are no more fights to record
    If fightCtr < numFighters - 1 Then
    
        'Get the ID for the next fight for the winner
        Dim winnerFightID As Integer
        winnerFightID = Utility.GetWinnerFightID(numFighters, fightCtr)
        
        'Get the player number for the next fight for the winner
        Dim winnerPlayerNum As Integer
        winnerPlayerNum = Utility.GetWinnerPlayerNum(fightCtr)
        
        'Set the appropriate player name
        bracket(winnerFightID - 1).SetPlayer playerName, winnerPlayerNum
    Else
        'Set the tournament winner
        tournamentWinner = playerName
        
    End If

End Sub


'Checks if this is the last fight
Function IsTournamentDone() As Boolean
    IsTournamentDone = (fightCtr = numFighters - 1)
End Function

'Go to the VictorySlide if this is the last fight
Sub LoadVictory()
    
    With victorySlide
        .Shapes("Winner").Fill.UserPicture tournamentWinner
    End With
    
    ActivePresentation.SlideShowWindow.View.GotoSlide (victorySlide.SlideIndex)

End Sub

'Reset the presentation so that users can understand what shapes are
Sub EndTournament()

    With fightSlide
        .Shapes("1").TextFrame.TextRange.Text = "Player 1"
        .Shapes("2").TextFrame.TextRange.Text = "Player 2"
        
        .Shapes("1").Visible = msoTrue
        .Shapes("2").Visible = msoTrue
        .Shapes("1Cover").Visible = msoTrue
        .Shapes("2Cover").Visible = msoTrue
        
        .Shapes("1Cover").Fill.Transparency = 0.5
        .Shapes("2Cover").Fill.Transparency = 0.5
    End With

    'On the loading slide, reset the round and group displays
        With loadingScreenSlide
        .Shapes("RoundID").TextFrame.TextRange.Text = "[ROUND#]"
        .Shapes("GroupID").TextFrame.TextRange.Text = "[BLOCKLABEL]"
    End With
    
    'On the victory slide, reset the shapes
    With victorySlide
        .Shapes("Winner").Fill.ForeColor.RGB = RGB(128, 0, 0)
        .Shapes("Winner").TextFrame.TextRange.Text = "Winner"
    End With
            
    ActivePresentation.SlideShowWindow.View.Exit
    
End Sub

'Re-covers the match
Sub RecoverMatch()

    InitializeTournament
    
    'Collect the fight information
    Dim targetFight As Integer
    targetFight = CInt(InputBox("Enter Fight # to rejoin on:"))
    
    'Loop through all the fights up to that point
    For i = 0 To targetFight - 2
    
        'Simulate Loading Screen
        fightCtr = fightCtr + 1
    
        Recovery.Controls("Fight").Caption = fightCtr
        Recovery.Repaint
        Recovery.Show
        
        'The user will select Player 1 or Player 2 and the button action updates the winner
    Next i
    
    MsgBox "Bracket has been recovered. Click OK to proceed to the loading screen"
    LoadLoadingScreen

End Sub

Sub ReturnToController()

    If Controller.IsTournamentDone Then
        Controller.LoadVictory
    Else
        Controller.LoadLoadingScreen
    End If

End Sub




