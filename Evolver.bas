Attribute VB_Name = "Evolver"
Function MaxEvolves(CandyCost As Integer, CandyCount As Integer) As Integer

    While CandyCount >= CandyCost
        MaxEvolves = MaxEvolves + Int(CandyCount / CandyCost)
        CandyCount = CandyCount - Int(CandyCount / CandyCost) * (CandyCost - 1)
    Wend
    
End Function
