Attribute VB_Name = "IVCalculator"
Function IVSolution(BaseHP As Integer, BaseAtk As Integer, BaseDef As Integer, rminHP As Integer, rmaxHP As Integer, minADS As Double, maxADS As Double, AppraisalSum As String, AppraisalHP As Integer, AppraisalAtk As Integer, AppraisalDef As Integer, AppraisalBest As String, ProjectedBaseHP As Integer, ProjectedBaseAtk As Integer, ProjectedBaseDef As Integer) As Variant
    
    ' Initalize IV bounds
    Dim aminIVSum As Integer, amaxIVSum As Integer, aminHP As Integer, amaxHP As Integer, aminAtk As Integer, amaxAtk As Integer, aminDef As Integer, amaxDef As Integer, aminIV As Integer, amaxIV As Integer
    aminHP = rminHP
    amaxHP = rmaxHP
    aminAtk = 0
    amaxAtk = 15
    aminDef = 0
    amaxDef = 15
    
    ' Implement appraisal restrictions
    If AppraisalBest = "A" Then
        aminIV = Range("minIVA")
        amaxIV = Range("maxIVA")
    ElseIf AppraisalBest = "B" Then
        aminIV = Range("minIVB")
        amaxIV = Range("maxIVB")
    ElseIf AppraisalBest = "C" Then
        aminIV = Range("minIVC")
        amaxIV = Range("maxIVC")
    ElseIf AppraisalBest = "D" Then
        aminIV = Range("minIVD")
        amaxIV = Range("maxIVD")
    End If
    
    If AppraisalHP = 1 Then
        aminHP = Application.WorksheetFunction.Max(aminIV, aminHP)
        amaxHP = Application.WorksheetFunction.Min(amaxIV, amaxHP)
    End If
    
    If AppraisalAtk = 1 Then
        aminAtk = aminIV
        amaxAtk = amaxIV
    End If
    
    If AppraisalDef = 1 Then
        aminDef = aminIV
        amaxDef = amaxIV
    End If
    
    aminIVSum = aminHP + aminAtk + aminDef
    amaxIVSum = amaxHP + amaxAtk + amaxDef
    
    If AppraisalSum = "A" Then
        aminIVSum = Application.WorksheetFunction.Max(Range("minIVSumA"), aminIVSum)
        amaxIVSum = Application.WorksheetFunction.Min(Range("maxIVSumA"), amaxIVSum)
    ElseIf AppraisalSum = "B" Then
        aminIVSum = Application.WorksheetFunction.Max(Range("minIVSumB"), aminIVSum)
        amaxIVSum = Application.WorksheetFunction.Min(Range("maxIVSumB"), amaxIVSum)
    ElseIf AppraisalSum = "C" Then
        aminIVSum = Application.WorksheetFunction.Max(Range("minIVSumC"), aminIVSum)
        amaxIVSum = Application.WorksheetFunction.Min(Range("maxIVSumC"), amaxIVSum)
    ElseIf AppraisalSum = "D" Then
        aminIVSum = Application.WorksheetFunction.Max(Range("minIVSumD"), aminIVSum)
        amaxIVSum = Application.WorksheetFunction.Min(Range("maxIVSumD"), amaxIVSum)
    End If
    
    ' Define results
    Dim Solutions As Integer, minIVSum As Integer, maxIVSum As Integer, minHP As Integer, maxHP As Integer, minAtk As Integer, maxAtk As Integer, minDef As Integer, maxDef As Integer, IVSum As Integer, aHPAtk As Integer, aHPDef As Integer, aAtkDef As Integer

    ' Define storage of intermediate calculation for CP checking
    Dim ADS As Double, ProjectedADS As Double, ProjectedMinADS As Double, ProjectedMaxADS As Double
    
    ' Initialize minimum and maximum possible results
    minIVSum = 46
    maxIVSum = -1
    minHP = 16
    maxHP = -1
    minAtk = 16
    maxHP = -1
    minDef = 16
    maxDef = -1
    ProjectedMaxADS = ProjectedBaseAtk ^ 2 * ProjectedBaseDef * ProjectedBaseHP
    ProjectedMinADS = (ProjectedBaseAtk + 15) ^ 2 * (ProjectedBaseDef + 15) * (ProjectedBaseHP + 15)
    
    ' Create multipliers for appraisal logic
    aHPAtk = AppraisalHP - AppraisalAtk
    aHPDef = AppraisalHP - AppraisalDef
    aAtkDef = AppraisalAtk - AppraisalDef
    
    ' Iterate through all possible IV combinations
    For HP = aminHP To amaxHP
        For Atk = aminAtk To amaxAtk
            For Def = aminDef To amaxDef
                ADS = (BaseAtk + Atk) ^ 2 * (BaseDef + Def) * (BaseHP + HP)
                IVSum = HP + Atk + Def
                If HP >= rminHP And HP <= rmaxHP And ADS >= minADS And ADS <= maxADS And IVSum >= aminIVSum And IVSum <= amaxIVSum And (aHPAtk = 0 Or aHPAtk * HP > aHPAtk * Atk) And (aHPDef = 0 Or aHPDef * HP > aHPDef * Def) And (aAtkDef = 0 Or aAtkDef * Atk > aAtkDef * Def) Then
                    Solutions = Solutions + 1
                    minIVSum = Application.WorksheetFunction.Min(minIVSum, IVSum)
                    maxIVSum = Application.WorksheetFunction.Max(maxIVSum, IVSum)
                    minHP = Application.WorksheetFunction.Min(minHP, HP)
                    maxHP = Application.WorksheetFunction.Max(maxHP, HP)
                    minAtk = Application.WorksheetFunction.Min(minAtk, Atk)
                    maxAtk = Application.WorksheetFunction.Max(maxAtk, Atk)
                    minDef = Application.WorksheetFunction.Min(minDef, Def)
                    maxDef = Application.WorksheetFunction.Max(maxDef, Def)
                    ProjectedADS = (ProjectedBaseAtk + Atk) ^ 2 * (ProjectedBaseDef + Def) * (ProjectedBaseHP + HP)
                    ProjectedMinADS = Application.WorksheetFunction.Min(ProjectedMinADS, ProjectedADS)
                    ProjectedMaxADS = Application.WorksheetFunction.Max(ProjectedMaxADS, ProjectedADS)
                End If
            Next
        Next
    Next
    
    ' Display results
    Dim Result(12) As Variant
    
    If Solutions = 0 Then
    
        Result(0) = Solutions
        Result(1) = ""
        Result(2) = ""
        Result(3) = ""
        Result(4) = ""
        Result(5) = ""
        Result(6) = ""
        Result(7) = ""
        Result(8) = ""
        Result(9) = ""
        Result(10) = ""
        Result(11) = ""
        Result(12) = ""

    Else

        Result(0) = Solutions
        Result(1) = minIVSum
        Result(2) = maxIVSum
        Result(3) = minIVSum / 45
        Result(4) = maxIVSum / 45
        Result(5) = minHP
        Result(6) = maxHP
        Result(7) = minAtk
        Result(8) = maxAtk
        Result(9) = minDef
        Result(10) = maxDef
        Result(11) = ProjectedMinADS
        Result(12) = ProjectedMaxADS
        
    End If
        
    IVSolution = Result
        
End Function

