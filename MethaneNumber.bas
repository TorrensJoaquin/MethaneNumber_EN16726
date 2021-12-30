Option Explicit
Const OptimalAmountsOfComponentsRepresentedInTheTernary As Byte = 3
Dim a(1 To 20, 0 To 7, 0 To 6) As Double
Dim MinVmaxOverVSum(1 To 11, 1 To 18) As Double
Dim TernaryComponents(1 To 18, 1 To 11) As Boolean 'HotOne matrix of components inside a specific ternary
Dim xyzOfTernary(1 To 18, 1 To 3) As Byte
Dim CompensationForShortTernary(1 To 18) As Byte
Dim xMax(1 To 20) As Byte
Dim xMin(1 To 20) As Byte
Dim yMax(1 To 20) As Byte
Dim yMin(1 To 20) As Byte
Dim zMax(1 To 20) As Byte
Dim zMin(1 To 20) As Byte
Dim StandardDeviationOfTheSolver As Single
Dim CheckIfAnImprovementIsDoneInTheLastXMovements As Boolean
Function DebuggerShowTheSelectedTernary(Methane As Double, Ethane As Double, Propane As Double, iButane As Double, nButane As Double, ipentane As Double, npentane As Double, Hexanes As Double, Nitrogen As Double, CarbonDioxide As Double, Hydrogen As Double, CarbonMonoxide As Double, Butadiene As Double, Butylene As Double, Ethylene As Double, Propylene As Double, HydrogenSulphide As Double) As Variant
    Dim SimplifiedChromatografy As Variant
    Dim MethaneNumberMWMWithoutInerts As Variant
    Dim IsThisComponentPresentHotOnes(1 To 11) As Boolean
    Dim IsThisComponentPresentInThisTernaryHotOnes(1 To 11, 1 To 18) As Boolean
    Dim HowManyComponentsAreRepresentedInThisTernary(1 To 18) As Byte
    Dim HowManyTimesIsTheComponentRepresented(1 To 11) As Byte
    Dim AffinitiesOfEachTernary(1 To 18) As Double
    Dim WillWeBeUsingThisTernaryHotOnes(1 To 18) As Boolean
    Dim NAji(1 To 11, 1 To 18) As Single
    Dim VAji(1 To 11, 1 To 18) As Single
    Dim NumberOfVariablesToChange As Byte
    Dim CalculatedMethaneNumbers() As Variant
    Dim SumOfNAjiComponentsInTheTernary(1 To 18) As Single
    Call UploadTheCoefficients
    SimplifiedChromatografy = SimplifyChromatografy(Methane, Ethane, Propane, iButane, nButane, ipentane, npentane, Hexanes, Nitrogen, CarbonDioxide, Hydrogen, CarbonMonoxide, Butadiene, Butylene, Ethylene, Propylene, HydrogenSulphide)
    Call CalculateIsThisComponentPresentHotOnes(SimplifiedChromatografy, IsThisComponentPresentHotOnes)
    Call CalculateHowManyComponentsAreRepresentedInThisTernary(IsThisComponentPresentHotOnes, HowManyComponentsAreRepresentedInThisTernary)
    Call CalculateAffinitiesOfEachTernary(SimplifiedChromatografy, AffinitiesOfEachTernary)
    Call CalculateHowManyTimesIsTheComponentRepresented(HowManyComponentsAreRepresentedInThisTernary, AffinitiesOfEachTernary, IsThisComponentPresentHotOnes, HowManyTimesIsTheComponentRepresented, WillWeBeUsingThisTernaryHotOnes, NAji, VAji)
    DebuggerShowTheSelectedTernary = WorksheetFunction.Transpose(WillWeBeUsingThisTernaryHotOnes)
End Function
Function MethaneNumberMWM(Methane As Double, Ethane As Double, Propane As Double, iButane As Double, nButane As Double, ipentane As Double, npentane As Double, Hexanes As Double, Nitrogen As Double, CarbonDioxide As Double, Hydrogen As Double, CarbonMonoxide As Double, Butadiene As Double, Butylene As Double, Ethylene As Double, Propylene As Double, HydrogenSulphide As Double) As Variant
    Dim SimplifiedChromatografy As Variant
    Dim MethaneNumberMWMWithoutInerts As Variant
    Call UploadTheCoefficients
    SimplifiedChromatografy = SimplifyChromatografy(Methane, Ethane, Propane, iButane, nButane, ipentane, npentane, Hexanes, Nitrogen, CarbonDioxide, Hydrogen, CarbonMonoxide, Butadiene, Butylene, Ethylene, Propylene, HydrogenSulphide)
    MethaneNumberMWMWithoutInerts = CalculateMethaneNumberMWM(SimplifiedChromatografy)
    MethaneNumberMWM = MethaneNumberMWMWithoutInerts + CorrectingMethaneNumberWithInerts(Methane, Ethane, Propane, iButane, nButane, ipentane, npentane, Hexanes, Nitrogen, CarbonDioxide, Hydrogen, CarbonMonoxide, Butadiene, Butylene, Ethylene, Propylene, HydrogenSulphide) - 100.0003
End Function
Private Function CorrectingMethaneNumberWithInerts(Methane As Double, Ethane As Double, Propane As Double, iButane As Double, nButane As Double, ipentane As Double, npentane As Double, Hexanes As Double, Nitrogen As Double, CarbonDioxide As Double, Hydrogen As Double, CarbonMonoxide As Double, Butadiene As Double, Butylene As Double, Ethylene As Double, Propylene As Double, HydrogenSulphide As Double)
    Dim NewMethaneContent As Double
    Dim SumOfComponents As Double
    Dim i As Byte
    Dim j As Byte
    NewMethaneContent = Methane + Ethane + Propane + iButane + nButane + ipentane + npentane + Hexanes + Hydrogen + CarbonMonoxide + Butadiene + Butylene + Ethylene + Propylene + HydrogenSulphide
    SumOfComponents = NewMethaneContent + CarbonDioxide + Nitrogen
    NewMethaneContent = NewMethaneContent * 100 / (SumOfComponents - Nitrogen)
    CarbonDioxide = CarbonDioxide * 100 / (SumOfComponents - Nitrogen)
'    NewMethaneContent = NewMethaneContent * SumOfComponents / (NewMethaneContent + CarbonDioxide)
'    CarbonDioxide = CarbonDioxide * SumOfComponents / (NewMethaneContent + CarbonDioxide)
    For i = 0 To 7
        For j = 0 To 6
            CorrectingMethaneNumberWithInerts = CorrectingMethaneNumberWithInerts + a(20, i, j) * NewMethaneContent ^ i * CarbonDioxide ^ j
        Next j
    Next i
End Function
Private Function CalculateMethaneNumberMWM(SimplifiedChromatografy As Variant) As Variant
    Dim IsThisComponentPresentHotOnes(1 To 11) As Boolean
    Dim IsThisComponentPresentInThisTernaryHotOnes As Variant
    Dim HowManyComponentsAreRepresentedInThisTernary(1 To 18) As Byte
    Dim HowManyTimesIsTheComponentRepresented(1 To 11) As Byte
    Dim AffinitiesOfEachTernary(1 To 18) As Double
    Dim WillWeBeUsingThisTernaryHotOnes(1 To 18) As Boolean
    Dim NAji(1 To 11, 1 To 18) As Single
    Dim VAji(1 To 11, 1 To 18) As Single
    Dim MinimumNAji(1 To 11, 1 To 18) As Single
    Dim NumberOfVariablesToChange As Byte
    Dim CalculatedMethaneNumbers() As Variant
    Dim SumOfNAjiComponentsInTheTernary(1 To 18) As Single
    '
    Call CalculateIsThisComponentPresentHotOnes(SimplifiedChromatografy, IsThisComponentPresentHotOnes)
    Call CalculateHowManyComponentsAreRepresentedInThisTernary(IsThisComponentPresentHotOnes, HowManyComponentsAreRepresentedInThisTernary)
    Call CalculateAffinitiesOfEachTernary(SimplifiedChromatografy, AffinitiesOfEachTernary)
    Call CalculateHowManyTimesIsTheComponentRepresented(HowManyComponentsAreRepresentedInThisTernary, AffinitiesOfEachTernary, IsThisComponentPresentHotOnes, HowManyTimesIsTheComponentRepresented, WillWeBeUsingThisTernaryHotOnes, NAji, VAji)
    '
    Dim i As Byte
    Dim j As Byte
    Dim x As Long
    Dim RangeMinMaxAvgValueOfTheResult(1 To 4) As Single
    Dim ActualMinimumRangeOfTheResultAchieved  As Single
    Dim ActualBestCalculatedMethaneNumber As Single
    Dim WhichCalculatedMethaneNumber As Byte
    WhichCalculatedMethaneNumber = 1
    IsThisComponentPresentInThisTernaryHotOnes = GetMeIsThisComponentPresentInThisTernaryHotOnes(IsThisComponentPresentHotOnes, WillWeBeUsingThisTernaryHotOnes, TernaryComponents)
    For x = 1 To 10000
        Call CalculateVAji(IsThisComponentPresentInThisTernaryHotOnes, SimplifiedChromatografy, VAji, SumOfNAjiComponentsInTheTernary, NAji, x, MinimumNAji)
        If IsThisCompositionInsideBoundarys(VAji) Then
            Call CalculateMethaneNumber(WillWeBeUsingThisTernaryHotOnes, VAji, CalculatedMethaneNumbers, RangeMinMaxAvgValueOfTheResult)
            If ActualMinimumRangeOfTheResultAchieved = 0 Or RangeMinMaxAvgValueOfTheResult(1) < ActualMinimumRangeOfTheResultAchieved Then
                CalculateMethaneNumberMWM = 0
                ActualMinimumRangeOfTheResultAchieved = RangeMinMaxAvgValueOfTheResult(1)
                CheckIfAnImprovementIsDoneInTheLastXMovements = True
                For i = 1 To 18
                    If WillWeBeUsingThisTernaryHotOnes(i) Then
                        CalculateMethaneNumberMWM = CalculateMethaneNumberMWM + CalculatedMethaneNumbers(WhichCalculatedMethaneNumber) * SumOfNAjiComponentsInTheTernary(i) / 100
                        WhichCalculatedMethaneNumber = WhichCalculatedMethaneNumber + 1
                        For j = 1 To 11
                            MinimumNAji(j, i) = NAji(j, i)
                        Next j
                    End If
                Next i
                'Debug.Print "Iteracion N°: " & x & " : " & Int(RangeMinMaxAvgValueOfTheResult(2)) & " " & Int(RangeMinMaxAvgValueOfTheResult(4)) & " " & Int(RangeMinMaxAvgValueOfTheResult(3)) & "  MN: " & CalculateMethaneNumberMWM & "  Rango: " & RangeMinMaxAvgValueOfTheResult(1) & " " & StandardDeviationOfTheSolver
                WhichCalculatedMethaneNumber = 1
                If RangeMinMaxAvgValueOfTheResult(1) < 0.1 Then
                    Exit For
                End If
            End If
        End If
    Next x
End Function
Private Sub CalculateMethaneNumber(WillWeBeUsingThisTernaryHotOnes() As Boolean, VAji() As Single, CalculatedMethaneNumbers As Variant, RangeMinMaxAvgValueOfTheResult() As Single)
    Dim i As Byte
    ReDim CalculatedMethaneNumbers(1 To 1)
    CalculatedMethaneNumbers(1) = 0
    For i = 1 To 18
        If WillWeBeUsingThisTernaryHotOnes(i) Then
            If CalculatedMethaneNumbers(1) <> 0 Then
                ReDim Preserve CalculatedMethaneNumbers(1 To UBound(CalculatedMethaneNumbers) + 1)
            End If
            CalculatedMethaneNumbers(UBound(CalculatedMethaneNumbers)) = FunctionA3(i, VAji)
        End If
    Next i
    'This value is important. Is going to be the Objective Function
    RangeMinMaxAvgValueOfTheResult(2) = WorksheetFunction.Min(CalculatedMethaneNumbers)
    RangeMinMaxAvgValueOfTheResult(3) = WorksheetFunction.Max(CalculatedMethaneNumbers)
    RangeMinMaxAvgValueOfTheResult(4) = WorksheetFunction.Average(CalculatedMethaneNumbers)
    RangeMinMaxAvgValueOfTheResult(1) = RangeMinMaxAvgValueOfTheResult(3) - RangeMinMaxAvgValueOfTheResult(2)
End Sub
Private Sub CalculateVAji(IsThisComponentPresentInThisTernaryHotOnes As Variant, SimplifiedChromatografy As Variant, VAji() As Single, SumOfNAjiComponentsInTheTernary() As Single, NAji() As Single, x As Long, MinimumNAji() As Single)
    'Create The first stage of FractionOfComponentInsideTernary.
    Dim FractionOfComponentInsideTernary(1 To 11, 1 To 18) As Single
    Dim RelationshipBetweenRandomNumbersAndTotalVolume(1 To 11) As Single
    Dim i As Byte
    Dim j As Byte
    'This is the pathfinding solver changing the precision.
    If x Mod 500 = 0 Then
        If CheckIfAnImprovementIsDoneInTheLastXMovements = False Then
            StandardDeviationOfTheSolver = StandardDeviationOfTheSolver * 0.5
        End If
        CheckIfAnImprovementIsDoneInTheLastXMovements = False
    End If
    '
    For i = 1 To 18
        SumOfNAjiComponentsInTheTernary(i) = 0
    Next i
    For i = 1 To 18
        For j = 1 To 11
            If IsThisComponentPresentInThisTernaryHotOnes(j, i) Then
                FractionOfComponentInsideTernary(j, i) = RandomizedNumberWithEvolutiveApproach(x, MinimumNAji, j, i)
                RelationshipBetweenRandomNumbersAndTotalVolume(j) = RelationshipBetweenRandomNumbersAndTotalVolume(j) + FractionOfComponentInsideTernary(j, i)
            End If
        Next j
    Next i
    'Create The second stage of FractionOfComponentInsideTernary -> NAji.
    For i = 1 To 18
        For j = 1 To 11
            If IsThisComponentPresentInThisTernaryHotOnes(j, i) Then
                NAji(j, i) = FractionOfComponentInsideTernary(j, i) * SimplifiedChromatografy(j) / RelationshipBetweenRandomNumbersAndTotalVolume(j)
                SumOfNAjiComponentsInTheTernary(i) = SumOfNAjiComponentsInTheTernary(i) + NAji(j, i)
            End If
        Next j
    Next i
    'Calculate VAji
    For i = 1 To 18
        For j = 1 To 11
            If NAji(j, i) <> 0 Then
                VAji(j, i) = NAji(j, i) * 100 / SumOfNAjiComponentsInTheTernary(i)
            End If
        Next j
    Next i
End Sub
Private Function RandomizedNumberWithEvolutiveApproach(x As Long, MinimumNAji() As Single, j As Byte, i As Byte)
    If x \ 1000 = 0 Then
        RandomizedNumberWithEvolutiveApproach = Rnd()
    Else
        'RandomizedNumberWithEvolutiveApproach = Rnd() * StandardDeviationOfTheSolver + MinimumNAji(j, i)
        RandomizedNumberWithEvolutiveApproach = Abs(WorksheetFunction.Norm_Inv(Rnd(), MinimumNAji(j, i), MinimumNAji(j, i) * StandardDeviationOfTheSolver))
    End If
End Function
Private Sub CalculateHowManyTimesIsTheComponentRepresented(HowManyComponentsAreRepresentedInThisTernary() As Byte, AffinitiesOfEachTernary() As Double, IsThisComponentPresentHotOnes() As Boolean, HowManyTimesIsTheComponentRepresented() As Byte, WillWeBeUsingThisTernaryHotOnes() As Boolean, NAji() As Single, VAji() As Single)
    'Inputs: HowManyComponentsAreRepresentedInThisTernary() as byte,AffinitiesOfEachTernary() as Single,IsThisComponentPresentHotOnes () as byte
    'Outputs: HowManyTimesIsTheComponentRepresented() as byte,WillWeBeUsingThisTernaryHotOnes() as boolean,IsThisComponentPresentInThisTernaryHotOnes() as boolean
    Dim RunAgainTheTernarySelectionAnalysis As Byte
    Dim CurrentComponentInAnalisys As Byte
    Dim TernaryCoveredInTheLastIteration(1 To 18) As Boolean
    Dim ActualTernarySelected As Byte
    Dim MinimumAmmountOfAceptableTernaryMixtures As Variant
    For RunAgainTheTernarySelectionAnalysis = 1 To 5
        Erase TernaryCoveredInTheLastIteration
        For CurrentComponentInAnalisys = 1 To 11
            MinimumAmmountOfAceptableTernaryMixtures = DoIAlreadyHaveTheMinimumAmmountOfAceptableTernaryMixtures(IsThisComponentPresentHotOnes, TernaryComponents, WillWeBeUsingThisTernaryHotOnes)
            If MinimumAmmountOfAceptableTernaryMixtures(0) Then
                Exit For
            End If
            If MinimumAmmountOfAceptableTernaryMixtures(CurrentComponentInAnalisys) = False And DoIveAlreadyCoveredThisComponentDuringThisIteration(CurrentComponentInAnalisys, TernaryCoveredInTheLastIteration) = False Then
                ActualTernarySelected = FindTheNextTernaryToBeSelected(CurrentComponentInAnalisys, HowManyComponentsAreRepresentedInThisTernary, AffinitiesOfEachTernary, WillWeBeUsingThisTernaryHotOnes)
                If ActualTernarySelected <> 0 Then
                    WillWeBeUsingThisTernaryHotOnes(ActualTernarySelected) = True
                    TernaryCoveredInTheLastIteration(ActualTernarySelected) = True
                    ActualTernarySelected = 0
                End If
            End If
            If MinimumAmmountOfAceptableTernaryMixtures(0) Then
                Exit For
            End If
        Next CurrentComponentInAnalisys
        If MinimumAmmountOfAceptableTernaryMixtures(0) Then
            Exit For
        End If
    Next RunAgainTheTernarySelectionAnalysis
End Sub
Private Function DoIveAlreadyCoveredThisComponentDuringThisIteration(CurrentComponentInAnalisys As Byte, TernaryCoveredInTheLastIteration() As Boolean) As Boolean
    Dim i As Byte
    For i = 1 To 18
        If TernaryCoveredInTheLastIteration(i) And MinVmaxOverVSum(CurrentComponentInAnalisys, i) <> 0 Then
            DoIveAlreadyCoveredThisComponentDuringThisIteration = True
        End If
    Next i
End Function
Private Function FindTheNextTernaryToBeSelected(CurrentComponentInAnalisys As Byte, HowManyComponentsAreRepresentedInThisTernary() As Byte, AffinitiesOfEachTernary() As Double, WillWeBeUsingThisTernaryHotOnes() As Boolean) As Byte
    Dim LowDownMyExpectationOnTheComponentsRepresentedInTheTernary As Byte
    Dim ActualAffinityOfTheTernarySelected As Single
    Dim ActualTernarySelected As Byte
    Dim i As Byte
    For LowDownMyExpectationOnTheComponentsRepresentedInTheTernary = 0 To OptimalAmountsOfComponentsRepresentedInTheTernary - 1
        ActualAffinityOfTheTernarySelected = 0
        ActualTernarySelected = 0
        For i = 1 To 18
            If MinVmaxOverVSum(CurrentComponentInAnalisys, i) > 0 And HowManyComponentsAreRepresentedInThisTernary(i) + CompensationForShortTernary(i) = OptimalAmountsOfComponentsRepresentedInTheTernary - LowDownMyExpectationOnTheComponentsRepresentedInTheTernary And WillWeBeUsingThisTernaryHotOnes(i) = False Then
                If AffinitiesOfEachTernary(i) > ActualAffinityOfTheTernarySelected Then
                    ActualAffinityOfTheTernarySelected = AffinitiesOfEachTernary(i)
                    ActualTernarySelected = i
                End If
            End If
        Next i
        If ActualTernarySelected <> 0 Then
            Exit For
        End If
    Next LowDownMyExpectationOnTheComponentsRepresentedInTheTernary
    FindTheNextTernaryToBeSelected = ActualTernarySelected
End Function
Private Function GetMeIsThisComponentPresentInThisTernaryHotOnes(IsThisComponentPresentHotOnes() As Boolean, WillWeBeUsingThisTernaryHotOnes() As Boolean, TernaryComponents() As Boolean) As Variant
    Dim IsThisComponentPresentInThisTernaryHotOnes(1 To 11, 1 To 18) As Boolean
    Dim i As Byte
    Dim j As Byte
    For i = 1 To 18
        For j = 1 To 11
        If IsThisComponentPresentHotOnes(j) And WillWeBeUsingThisTernaryHotOnes(i) And TernaryComponents(i, j) Then
            IsThisComponentPresentInThisTernaryHotOnes(j, i) = True
        End If
        Next j
    Next i
    GetMeIsThisComponentPresentInThisTernaryHotOnes = IsThisComponentPresentInThisTernaryHotOnes
End Function
Private Function DoIAlreadyHaveTheMinimumAmmountOfAceptableTernaryMixtures(IsThisComponentPresentHotOnes() As Boolean, TernaryComponents() As Boolean, WillWeBeUsingThisTernaryHotOnes() As Boolean) As Variant
    Dim i As Byte
    Dim j As Byte
    Dim HowManyComponentsAreRepresentedInThisTernary(1 To 18) As Byte
    Dim HowManyTimesIsTheComponentRepresented(1 To 11) As Byte
    Dim MinimumAmmountOfAceptableTernaryMixtures(0 To 11) As Boolean
    For j = 0 To 11
        MinimumAmmountOfAceptableTernaryMixtures(j) = True
    Next j
    For i = 1 To 18
        For j = 1 To 11
            If IsThisComponentPresentHotOnes(j) And WillWeBeUsingThisTernaryHotOnes(i) And TernaryComponents(i, j) Then
                HowManyTimesIsTheComponentRepresented(j) = HowManyTimesIsTheComponentRepresented(j) + 1
            End If
        Next j
    Next i
    For j = 1 To 11
        If IsThisComponentPresentHotOnes(j) And HowManyTimesIsTheComponentRepresented(j) < 2 Then
                MinimumAmmountOfAceptableTernaryMixtures(0) = False
                MinimumAmmountOfAceptableTernaryMixtures(j) = False
        End If
    Next j
    DoIAlreadyHaveTheMinimumAmmountOfAceptableTernaryMixtures = MinimumAmmountOfAceptableTernaryMixtures
End Function
Private Sub CalculateAffinitiesOfEachTernary(SimplifiedChromatografy As Variant, AffinitiesOfEachTernary() As Double)
    Dim i As Byte
    Dim j As Byte
    For i = 1 To 11
        For j = 1 To 18
            AffinitiesOfEachTernary(j) = AffinitiesOfEachTernary(j) + SimplifiedChromatografy(i) * MinVmaxOverVSum(i, j)
        Next j
    Next i
End Sub
Private Sub CalculateHowManyComponentsAreRepresentedInThisTernary(IsThisComponentPresentHotOnes() As Boolean, HowManyComponentsAreRepresentedInThisTernary() As Byte)
    Dim i As Byte
    Dim j As Byte
    For i = 1 To 18
        For j = 1 To 11
            If IsThisComponentPresentHotOnes(j) And TernaryComponents(i, j) Then
                HowManyComponentsAreRepresentedInThisTernary(i) = HowManyComponentsAreRepresentedInThisTernary(i) + 1
            End If
        Next j
    Next i
End Sub
Private Sub CalculateIsThisComponentPresentHotOnes(SimplifiedChromatografy As Variant, IsThisComponentPresentHotOnes() As Boolean)
    Dim j As Byte
    For j = 1 To 11
        If SimplifiedChromatografy(j) > 0.05 Then
            IsThisComponentPresentHotOnes(j) = True
        End If
    Next j
End Sub
Private Function IsThisCompositionInsideBoundarys(VAji() As Single) As Boolean
    IsThisCompositionInsideBoundarys = True
    Dim i As Byte
    Dim j As Byte
    For i = 1 To 18
        If xyzOfTernary(i, 1) <> 0 Then
            If VAji(xyzOfTernary(i, 1), i) > xMax(i) And VAji(xyzOfTernary(i, 1), i) < xMin(i) Then
                IsThisCompositionInsideBoundarys = False
                Exit For
            End If
        End If
        
        If xyzOfTernary(i, 2) <> 0 Then
            If VAji(xyzOfTernary(i, 2), i) > yMax(i) And VAji(xyzOfTernary(i, 2), i) < yMin(i) Then
                IsThisCompositionInsideBoundarys = False
                Exit For
            End If
        End If
        
        If xyzOfTernary(i, 3) <> 0 Then
            If VAji(xyzOfTernary(i, 3), i) > zMax(i) And VAji(xyzOfTernary(i, 3), i) < zMin(i) Then
                IsThisCompositionInsideBoundarys = False
                Exit For
            End If
        End If
    Next i
End Function
Private Function SimplifyChromatografy(Methane As Double, Ethane As Double, Propane As Double, iButane As Double, nButane As Double, ipentane As Double, npentane As Double, Hexanes As Double, Nitrogen As Double, CarbonDioxide As Double, Hydrogen As Double, CarbonMonoxide As Double, Butadiene As Double, Butylene As Double, Ethylene As Double, Propylene As Double, HydrogenSulphide) As Variant
    Dim Result(1 To 11) As Variant
    Dim SumOfComponents As Double
    Dim i As Byte
    Result(1) = CarbonMonoxide
    Result(2) = 0 'Butadiene
    Result(3) = 0 'Butylene
    Result(4) = Ethylene
    Result(5) = Propylene
    Result(6) = HydrogenSulphide
    Result(7) = Hydrogen
    Result(8) = Propane
    Result(9) = Ethane
    Result(10) = (iButane + nButane) + (ipentane + npentane) * 2.3 + Hexanes * 5.3 + Butadiene + Butylene
    Result(11) = Methane
    For i = 1 To 11
        SumOfComponents = SumOfComponents + Result(i)
    Next i
    For i = 1 To 11
        Result(i) = Result(i) * 100 / SumOfComponents
    Next i
    SimplifyChromatografy = Result
End Function
Private Sub UploadTheCoefficients()
    'Solver Coefficients
    StandardDeviationOfTheSolver = 1
    '
    CompensationForShortTernary(12) = 1
    CompensationForShortTernary(13) = 1
    CompensationForShortTernary(14) = 1
    CompensationForShortTernary(15) = 1
    CompensationForShortTernary(16) = 1
    CompensationForShortTernary(17) = 2
    CompensationForShortTernary(18) = 2
    'XYZ Boundarys
    xMax(1) = 100
    xMax(2) = 100
    xMax(3) = 100
    xMax(4) = 100
    xMax(5) = 100
    xMax(6) = 100
    xMax(7) = 100
    xMax(8) = 100
    xMax(9) = 100
    xMax(10) = 100
    xMax(11) = 100
    xMax(12) = 100
    xMax(13) = 100
    xMax(14) = 100
    xMax(15) = 100
    xMax(16) = 100
    xMax(17) = 100
    xMax(18) = 100
    xMax(20) = 100
    xMin(9) = 60
    xMin(10) = 60
    xMin(11) = 60
    xMin(17) = 100
    xMin(18) = 100
    xMin(20) = 35
    yMax(1) = 100
    yMax(2) = 100
    yMax(3) = 100
    yMax(4) = 100
    yMax(5) = 100
    yMax(6) = 100
    yMax(7) = 100
    yMax(8) = 100
    yMax(9) = 40
    yMax(10) = 40
    yMax(11) = 40
    yMax(12) = 100
    yMax(13) = 100
    yMax(14) = 100
    yMax(15) = 100
    yMax(16) = 100
    yMax(20) = 100
    zMax(1) = 100
    zMax(2) = 100
    zMax(3) = 100
    zMax(4) = 100
    zMax(5) = 100
    zMax(6) = 100
    zMax(7) = 100
    zMax(8) = 100
    zMax(9) = 40
    zMax(10) = 40
    zMax(11) = 40
    zMax(20) = 65
    'XYZ of Ternary components
    xyzOfTernary(1, 1) = 11
    xyzOfTernary(1, 2) = 7
    xyzOfTernary(1, 3) = 9
    xyzOfTernary(2, 1) = 8
    xyzOfTernary(2, 2) = 9
    xyzOfTernary(2, 3) = 10
    xyzOfTernary(3, 1) = 7
    xyzOfTernary(3, 2) = 8
    xyzOfTernary(3, 3) = 5
    xyzOfTernary(4, 1) = 11
    xyzOfTernary(4, 2) = 9
    xyzOfTernary(4, 3) = 8
    xyzOfTernary(5, 1) = 11
    xyzOfTernary(5, 2) = 7
    xyzOfTernary(5, 3) = 8
    xyzOfTernary(6, 1) = 11
    xyzOfTernary(6, 2) = 7
    xyzOfTernary(6, 3) = 10
    xyzOfTernary(7, 1) = 11
    xyzOfTernary(7, 2) = 8
    xyzOfTernary(7, 3) = 10
    xyzOfTernary(8, 1) = 11
    xyzOfTernary(8, 2) = 9
    xyzOfTernary(8, 3) = 10
    xyzOfTernary(9, 1) = 11
    xyzOfTernary(9, 2) = 4
    xyzOfTernary(9, 3) = 10
    xyzOfTernary(10, 1) = 11
    xyzOfTernary(10, 2) = 6
    xyzOfTernary(10, 3) = 10
    xyzOfTernary(11, 1) = 11
    xyzOfTernary(11, 2) = 9
    xyzOfTernary(11, 3) = 6
    xyzOfTernary(12, 1) = 11
    xyzOfTernary(12, 2) = 5
    xyzOfTernary(13, 1) = 9
    xyzOfTernary(13, 2) = 5
    xyzOfTernary(14, 1) = 1
    xyzOfTernary(14, 2) = 7
    xyzOfTernary(15, 1) = 9
    xyzOfTernary(15, 2) = 4
    xyzOfTernary(16, 1) = 8
    xyzOfTernary(16, 2) = 4
    xyzOfTernary(17, 1) = 2
    xyzOfTernary(18, 1) = 3
    TernaryComponents(1, 7) = True
    TernaryComponents(1, 9) = True
    TernaryComponents(1, 11) = True
    TernaryComponents(2, 8) = True
    TernaryComponents(2, 9) = True
    TernaryComponents(2, 10) = True
    TernaryComponents(3, 5) = True
    TernaryComponents(3, 7) = True
    TernaryComponents(3, 8) = True
    TernaryComponents(4, 8) = True
    TernaryComponents(4, 9) = True
    TernaryComponents(4, 11) = True
    TernaryComponents(5, 7) = True
    TernaryComponents(5, 8) = True
    TernaryComponents(5, 11) = True
    TernaryComponents(6, 7) = True
    TernaryComponents(6, 10) = True
    TernaryComponents(6, 11) = True
    TernaryComponents(7, 8) = True
    TernaryComponents(7, 10) = True
    TernaryComponents(7, 11) = True
    TernaryComponents(8, 9) = True
    TernaryComponents(8, 10) = True
    TernaryComponents(8, 11) = True
    TernaryComponents(9, 4) = True
    TernaryComponents(9, 10) = True
    TernaryComponents(9, 11) = True
    TernaryComponents(10, 6) = True
    TernaryComponents(10, 10) = True
    TernaryComponents(10, 11) = True
    TernaryComponents(11, 6) = True
    TernaryComponents(11, 9) = True
    TernaryComponents(11, 11) = True
    TernaryComponents(12, 5) = True
    TernaryComponents(12, 11) = True
    TernaryComponents(13, 5) = True
    TernaryComponents(13, 9) = True
    TernaryComponents(14, 1) = True
    TernaryComponents(14, 7) = True
    TernaryComponents(15, 4) = True
    TernaryComponents(15, 9) = True
    TernaryComponents(16, 4) = True
    TernaryComponents(16, 8) = True
    TernaryComponents(17, 2) = True
    TernaryComponents(18, 3) = True
    ''MinVmaxOverVSum
    MinVmaxOverVSum(1, 14) = 1
    MinVmaxOverVSum(2, 17) = 1
    MinVmaxOverVSum(3, 18) = 1
    MinVmaxOverVSum(4, 9) = 0.166666667
    MinVmaxOverVSum(4, 15) = 0.416666667
    MinVmaxOverVSum(4, 16) = 0.416666667
    MinVmaxOverVSum(5, 3) = 0.333333333
    MinVmaxOverVSum(5, 12) = 0.333333333
    MinVmaxOverVSum(5, 13) = 0.333333333
    MinVmaxOverVSum(6, 10) = 0.5
    MinVmaxOverVSum(6, 11) = 0.5
    MinVmaxOverVSum(7, 1) = 0.2
    MinVmaxOverVSum(7, 3) = 0.2
    MinVmaxOverVSum(7, 5) = 0.2
    MinVmaxOverVSum(7, 6) = 0.2
    MinVmaxOverVSum(7, 14) = 0.2
    MinVmaxOverVSum(8, 2) = 0.166666667
    MinVmaxOverVSum(8, 3) = 0.166666667
    MinVmaxOverVSum(8, 4) = 0.166666667
    MinVmaxOverVSum(8, 5) = 0.166666667
    MinVmaxOverVSum(8, 7) = 0.166666667
    MinVmaxOverVSum(8, 16) = 0.166666667
    MinVmaxOverVSum(9, 1) = 0.15625
    MinVmaxOverVSum(9, 2) = 0.15625
    MinVmaxOverVSum(9, 4) = 0.15625
    MinVmaxOverVSum(9, 8) = 0.15625
    MinVmaxOverVSum(9, 11) = 0.0625
    MinVmaxOverVSum(9, 13) = 0.15625
    MinVmaxOverVSum(9, 15) = 0.15625
    MinVmaxOverVSum(10, 2) = 0.208333333
    MinVmaxOverVSum(10, 6) = 0.208333333
    MinVmaxOverVSum(10, 7) = 0.208333333
    MinVmaxOverVSum(10, 8) = 0.208333333
    MinVmaxOverVSum(10, 9) = 0.083333333
    MinVmaxOverVSum(10, 10) = 0.083333333
    MinVmaxOverVSum(11, 1) = 0.1
    MinVmaxOverVSum(11, 4) = 0.1
    MinVmaxOverVSum(11, 5) = 0.1
    MinVmaxOverVSum(11, 6) = 0.1
    MinVmaxOverVSum(11, 7) = 0.1
    MinVmaxOverVSum(11, 8) = 0.1
    MinVmaxOverVSum(11, 9) = 0.1
    MinVmaxOverVSum(11, 10) = 0.1
    MinVmaxOverVSum(11, 11) = 0.1
    MinVmaxOverVSum(11, 12) = 0.1
''a coefficients
    a(1, 0, 0) = 43.62819
    a(1, 1, 0) = -0.09250887
    a(1, 0, 1) = -0.01048858
    a(1, 2, 0) = 0.01644927
    a(1, 1, 1) = -0.002500773
    a(1, 0, 2) = -0.004320274
    a(1, 3, 0) = -0.0003119169
    a(1, 2, 1) = -0.00006048696
    a(1, 1, 2) = -0.00005352801
    a(1, 0, 3) = 0.00006850742
    a(1, 4, 0) = 0.000002122334
    a(1, 3, 1) = 0.00000219937
    a(1, 2, 2) = 0.000001210969
    a(1, 1, 3) = 0.0000002970658
    a(1, 0, 4) = -0.0000006713802
    a(2, 0, 0) = 10.24513
    a(2, 1, 0) = 0.08590661
    a(2, 0, 1) = 0.1498213
    a(2, 2, 0) = 0.007384396
    a(2, 1, 1) = 0.009570504
    a(2, 0, 2) = 0.005136971
    a(2, 3, 0) = -0.0001003662
    a(2, 2, 1) = -0.0002020327
    a(2, 1, 2) = -0.00004580277
    a(2, 0, 3) = -0.00005685615
    a(2, 4, 0) = 0.0000004127305
    a(2, 3, 1) = 0.000001251138
    a(2, 2, 2) = 0.0000003114703
    a(2, 1, 3) = -0.0000003140157
    a(2, 0, 4) = 0.0000002403948
    a(3, 0, 0) = 18.62794
    a(3, 1, 0) = -0.1203581
    a(3, 0, 1) = 0.1087109
    a(3, 2, 0) = 0.01929801
    a(3, 1, 1) = -0.001305063
    a(3, 0, 2) = 0.0017985
    a(3, 3, 0) = -0.001301808
    a(3, 2, 1) = 0.00002990447
    a(3, 1, 2) = 0.00008561376
    a(3, 0, 3) = -0.00002583667
    a(3, 4, 0) = 0.00004169295
    a(3, 3, 1) = 0.0000002001124
    a(3, 2, 2) = -0.0000006854646
    a(3, 1, 3) = -0.0000006262613
    a(3, 0, 4) = 0.0000001198789
    a(3, 5, 0) = -0.0000006952638
    a(3, 6, 0) = 0.000000005798984
    a(3, 7, 0) = -1.913374E-11
    a(4, 0, 0) = 33.53909
    a(4, 1, 0) = -0.1028224
    a(4, 0, 1) = 0.2068375
    a(4, 2, 0) = 0.02398141
    a(4, 1, 1) = 0.003316137
    a(4, 0, 2) = -0.003553689
    a(4, 3, 0) = -0.0009584746
    a(4, 2, 1) = -0.0002409604
    a(4, 1, 2) = 0.0000394184
    a(4, 0, 3) = 0.00005001856
    a(4, 4, 0) = 0.00002005288
    a(4, 3, 1) = 0.00000345851
    a(4, 2, 2) = 0.0000008036454
    a(4, 1, 3) = -0.0000004333876
    a(4, 0, 4) = -0.0000002504256
    a(4, 5, 0) = -0.0000002115417
    a(4, 6, 0) = 0.000000000905402
    a(5, 0, 0) = 34.75804
    a(5, 1, 0) = -0.5194905
    a(5, 0, 1) = 0.05473705
    a(5, 2, 0) = 0.04405446
    a(5, 1, 1) = 0.02642531
    a(5, 0, 2) = -0.01056781
    a(5, 3, 0) = -0.0008743329
    a(5, 2, 1) = -0.001084645
    a(5, 1, 2) = -0.0003555327
    a(5, 0, 3) = 0.0002289769
    a(5, 4, 0) = 0.000005476742
    a(5, 3, 1) = 0.0000113098
    a(5, 2, 2) = 0.000007987488
    a(5, 1, 3) = 0.0000007486085
    a(5, 0, 4) = -0.000001634024
    a(6, 0, 0) = 12.29902
    a(6, 1, 0) = -0.7518207
    a(6, 0, 1) = -0.451037
    a(6, 2, 0) = 0.05143333
    a(6, 1, 1) = 0.05126147
    a(6, 0, 2) = 0.0178663
    a(6, 3, 0) = -0.001024159
    a(6, 2, 1) = -0.001640652
    a(6, 1, 2) = -0.00100224
    a(6, 0, 3) = -0.0001427912
    a(6, 4, 0) = 0.000006699563
    a(6, 3, 1) = 0.00001566121
    a(6, 2, 2) = 0.00001576306
    a(6, 1, 3) = 0.000005249888
    a(7, 0, 0) = 10.16914
    a(7, 1, 0) = 0.4366612
    a(7, 0, 1) = 0.03817096
    a(7, 2, 0) = -0.08726454
    a(7, 1, 1) = -0.007947864
    a(7, 0, 2) = 0.01036501
    a(7, 3, 0) = 0.005939795
    a(7, 2, 1) = 0.0003267886
    a(7, 1, 2) = 0.0002371491
    a(7, 0, 3) = -0.0001615215
    a(7, 4, 0) = -0.0001854127
    a(7, 3, 1) = -0.0000003308586
    a(7, 2, 2) = -0.000004975863
    a(7, 1, 3) = -0.0000008782291
    a(7, 0, 4) = 0.000000774084
    a(7, 5, 0) = 0.000002956598
    a(7, 6, 0) = -0.00000002337074
    a(7, 7, 0) = 7.322348E-11
    a(8, 0, 0) = 10.77761
    a(8, 1, 0) = 0.164749
    a(8, 0, 1) = -0.1405007
    a(8, 2, 0) = -0.0519873
    a(8, 1, 1) = -0.007044869
    a(8, 0, 2) = 0.01615437
    a(8, 3, 0) = 0.003991315
    a(8, 2, 1) = 0.0001479482
    a(8, 1, 2) = 0.0003384803
    a(8, 0, 3) = -0.000175467
    a(8, 4, 0) = -0.0001277487
    a(8, 3, 1) = 0.000002756444
    a(8, 2, 2) = -0.000004041667
    a(8, 1, 3) = -0.000001971021
    a(8, 0, 4) = 0.0000006075213
    a(8, 5, 0) = 0.000002015703
    a(8, 6, 0) = -0.00000001558017
    a(8, 7, 0) = 4.797693E-11
    a(9, 0, 0) = -124085.7
    a(9, 1, 0) = 11938.458
    a(9, 0, 1) = -199.62282
    a(9, 2, 0) = -485.74811
    a(9, 1, 1) = 7.8748002
    a(9, 0, 2) = 2.5929804
    a(9, 3, 0) = 10.855881
    a(9, 2, 1) = -0.10266703
    a(9, 1, 2) = -0.069109752
    a(9, 0, 3) = -0.0145046
    a(9, 4, 0) = -0.1441712
    a(9, 3, 1) = 0.00044431373
    a(9, 2, 2) = 0.00045679208
    a(9, 1, 3) = 0.0001987161
    a(9, 0, 4) = 0.000026937182
    a(9, 5, 0) = 0.001139533
    a(9, 6, 0) = -0.0000049703336
    a(9, 7, 0) = 9.2406348E-09
    a(10, 0, 0) = 183885.06
    a(10, 1, 0) = -15396.773
    a(10, 0, 1) = -14.160386
    a(10, 2, 0) = 541.58924
    a(10, 1, 1) = 0.56775484
    a(10, 0, 2) = 1.1942148
    a(10, 3, 0) = -10.358971
    a(10, 2, 1) = -0.0077071033
    a(10, 1, 2) = -0.024873835
    a(10, 0, 3) = -0.031209902
    a(10, 4, 0) = 0.11603083
    a(10, 3, 1) = 0.000033083382
    a(10, 2, 2) = 0.00017311782
    a(10, 1, 3) = 0.000004175449
    a(10, 0, 4) = 0.0015364226
    a(10, 5, 0) = -0.00075743018
    a(10, 6, 0) = 0.0000026462473
    a(10, 7, 0) = -3.7606039E-09
    a(10, 0, 5) = -0.00003565003
    a(10, 0, 6) = 0.00000030668448
    a(11, 0, 0) = -117884.66
    a(11, 1, 0) = 11251.043
    a(11, 0, 1) = -267.12519
    a(11, 2, 0) = -454.92745
    a(11, 1, 1) = 10.645736
    a(11, 0, 2) = 3.6669421
    a(11, 3, 0) = 10.120505
    a(11, 2, 1) = -0.13986048
    a(11, 1, 2) = -0.097497566
    a(11, 0, 3) = -0.024662769
    a(11, 4, 0) = -0.13401172
    a(11, 3, 1) = 0.00060764355
    a(11, 2, 2) = 0.00064613035
    a(11, 1, 3) = 0.00031927693
    a(11, 0, 4) = 0.000076292913
    a(11, 5, 0) = 0.001057975
    a(11, 6, 0) = -0.0000046175613
    a(11, 7, 0) = 8.6063163E-09
    a(12, 0, 0) = 59.095515
    a(12, 1, 0) = 0.10602705
    a(12, 0, 1) = -3.406924
    a(12, 2, 0) = -0.003188483
    a(12, 0, 2) = 0.15370325
    a(12, 3, 0) = -0.0001080121
    a(12, 0, 3) = -0.00367487
    a(12, 4, 0) = 0.00000845993
    a(12, 0, 4) = 0.000046273625
    a(12, 5, 0) = -0.00000013928745
    a(12, 6, 0) = 0.000000000716383
    a(12, 0, 5) = -0.0000002905423
    a(12, 0, 6) = 0.000000000716383
    a(13, 0, 0) = 31.5507
    a(13, 1, 0) = 0.0797494
    a(13, 0, 1) = -0.17706875
    a(13, 2, 0) = 0.00048659675
    a(13, 0, 2) = 0.00048659675
    a(14, 1, 0) = 1.5
    a(14, 2, 0) = -0.0075
    a(14, 1, 1) = -0.0075
    a(15, 0, 0) = 29.655595
    a(15, 1, 0) = 0.17064685
    a(15, 0, 1) = -0.12344405
    a(15, 2, 0) = -0.000236014
    a(15, 0, 2) = -0.000236014
    a(16, 0, 0) = 24.494755
    a(16, 1, 0) = 0.13676575
    a(16, 0, 1) = -0.0545979
    a(16, 2, 0) = -0.00041083915
    a(16, 0, 2) = -0.00041083915
    a(17, 0, 0) = 12
    a(18, 0, 0) = 20
    a(20, 0, 0) = 299.1743
    a(20, 1, 0) = -15.11958
    a(20, 0, 1) = -0.3115636
    a(20, 2, 0) = 0.7635948
    a(20, 1, 1) = 0.04548069
    a(20, 0, 2) = 0.01123041
    a(20, 3, 0) = -0.02376263
    a(20, 2, 1) = -0.0007856294
    a(20, 1, 2) = 0.0006555709
    a(20, 0, 3) = -0.002146855
    a(20, 4, 0) = 0.0004355494
    a(20, 3, 1) = 0.000003860668
    a(20, 2, 2) = -0.000001381699   'Typo found by the readme of MWM
    a(20, 1, 3) = -0.000007933902
    a(20, 0, 4) = 0.00006699364
    a(20, 5, 0) = -0.000004607726
    a(20, 6, 0) = 0.0000000261057
    a(20, 7, 0) = -6.143914E-11
    a(20, 0, 5) = -0.0000008369387
    a(20, 0, 6) = 0.000000003928073
End Sub
Private Function FunctionA3(t As Byte, VAji() As Single) As Single
    Dim i As Byte
    Dim j As Byte
    For i = 0 To 7
        For j = 0 To 6
            FunctionA3 = FunctionA3 + a(t, i, j) * VAji(xyzOfTernary(t, 1), t) ^ i * VAji(xyzOfTernary(t, 2), t) ^ j
        Next j
    Next i
End Function
