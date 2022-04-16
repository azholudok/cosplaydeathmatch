Attribute VB_Name = "Utility"
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'UTILITIES - Functions for calculating roundID, groupID, winnerFightID, Ceiling, and Floor
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


'Returns the current round based on the total number of fighters and the ID of the fight
Function GetRoundID(numFighters As Integer, fightID As Integer) As Integer

    'CEILING (log2(numFighters) - log2(numFighters - fightID))
    'Account for an edge case where fight 8 yields a value to pass of 1, but due to VBA data type conversion or approximiation,
    'does not equate to 1, which results in Ceiling(1) = 2
    Dim valueToPass As Double
    valueToPass = (Math.Log(numFighters) / Math.Log(2)) - (Math.Log(numFighters - fightID) / Math.Log(2))
    
    If fightID = 16 Or fightID = 24 Then
        valueToPass = valueToPass - 0.01
    End If
    
    GetRoundID = Int(Ceiling(valueToPass, 1))

End Function

'Returns the string label for the current group
Function GetGroupID(numFighters As Integer, fightID As Integer) As Integer

    Dim totalRounds As Integer
    totalRounds = Log(numFighters) / Log(2)
    
    
    Dim round As Integer
    round = GetRoundID(numFighters, fightID)

    Dim groupID As Integer
    
    'If this is the semi-finals, set to -1
    'If this is the finals, set to 0
    'Otherwise, calculate the groupID
    If round >= totalRounds - 1 Then
        groupID = round - totalRounds
    Else

        'Total number of fights as of last round
        Dim fightsAsOfLastRound As Integer
        fightsAsOfLastRound = numFighters * (1 - (0.5 ^ (round - 1)))
    
        Dim groupSize As Integer
        groupSize = 2 ^ (totalRounds - round - 2)
        
        groupID = Int(Ceiling((fightID - fightsAsOfLastRound) / groupSize))
    End If
        
    GetGroupID = groupID
        
End Function

'Returns the next fight ID for the winner of the current fight
Function GetWinnerFightID(numFighters As Integer, fightID As Integer) As Integer
    
    GetWinnerFightID = Ceiling(fightID / 2) + (numFighters / 2)

End Function

'Returns the next player number for the winner of the current fight
Function GetWinnerPlayerNum(fightID As Integer) As Integer
    
    If fightID Mod 2 = 1 Then
        GetWinnerPlayerNum = 1
    Else
        GetWinnerPlayerNum = 2
    End If
    
End Function



'Ceiling function pulled from http://www.tek-tips.com/faqs.cfm?fid=5031
Function Ceiling(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' is the multiple to which you want to round
    Ceiling = (Int(X / Factor) - (X / Factor - Int(X / Factor) > 0)) * Factor
End Function

Function Floor(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' is the multiple to which you want to round
    Floor = Int(X / Factor) * Factor
End Function
