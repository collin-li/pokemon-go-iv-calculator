Attribute VB_Name = "IVSolution"
Function IVSolution(BaseHP As Integer, BaseAtk As Integer, BaseDef As Integer, rminHP As Integer, rmaxHP As Integer, minADS As Double, maxADS As Double, AppraisalSum As String, AppraisalHP As Boolean, AppraisalAtk As Boolean, AppraisalDef As Boolean, AppraisalBest As String) As Variant
    
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
    
    If AppraisalHP Then
        aminHP = Application.WorksheetFunction.Max(aminIV, aminHP)
        amaxHP = Application.WorksheetFunction.Min(amaxIV, amaxHP)
    End If
    
    If AppraisalAtk Then
        aminAtk = aminIV
        amaxAtk = amaxIV
    End If
    
    If AppraisalDef Then
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
    Dim Solutions As Integer, minIVSum As Integer, maxIVSum As Integer, minHP As Integer, maxHP As Integer, minAtk As Integer, maxAtk As Integer, minDef As Integer, maxDef As Integer, IVSum As Integer

    ' Define storage of intermediate calculation for CP checking
    Dim ADS As Double
    
    ' Initialize minimum and maximum possible results
    minIVSum = 46
    maxIVSum = -1
    minHP = 16
    maxHP = -1
    minAtk = 16
    maxHP = -1
    minDef = 16
    maxDef = -1
    
    ' Iterate through all possible IV combinations
    For HP = aminHP To amaxHP
        For Atk = aminAtk To amaxAtk
            For Def = aminDef To amaxDef
                ADS = (BaseAtk + Atk) ^ 2 * (BaseDef + Def) * (BaseHP + HP)
                IVSum = HP + Atk + Def
                If HP >= rminHP And HP <= rmaxHP And ADS >= minADS And ADS <= maxADS And IVSum >= aminIVSum And IVSum <= amaxIVSum Then
                    minIVSum = Application.WorksheetFunction.Min(minIVSum, IVSum)
                    maxIVSum = Application.WorksheetFunction.Max(maxIVSum, IVSum)
                    minHP = Application.WorksheetFunction.Min(minHP, HP)
                    maxHP = Application.WorksheetFunction.Max(maxHP, HP)
                    minAtk = Application.WorksheetFunction.Min(minAtk, Atk)
                    maxAtk = Application.WorksheetFunction.Max(maxAtk, Atk)
                    minDef = Application.WorksheetFunction.Min(minDef, Def)
                    maxDef = Application.WorksheetFunction.Max(maxDef, Def)
                    Solutions = Solutions + 1
                End If
            Next
        Next
    Next
    
    ' Display results
    Dim Result(10) As Variant
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
    IVSolution = Result
End Function
