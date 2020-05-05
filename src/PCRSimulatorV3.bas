Attribute VB_Name = "PCRSimulatorV3"
'Version 3 of PCRSimulator
'Generalizes the approach in PCR_V2 to allow multiple templates/no templates by simulating dissociation and annealing process
'Based on Professor Anderson's NewPCRSimulator class in ConstructionFileSimulator
'Works well in regular PCR, but does not handle edge cases (IPCR, etc.) well - also extremely slow due to data structure constraints
Public Function PCR_V3(ByVal primer1 As String, ByVal primer2 As String, ByVal template As String) As String
    'Check that inputs are not empty
    If Not (Len(primer1) > 0 And Len(primer2) > 0) Then
        PCR_V3 = "Error: One or more primers does not exist."
        Exit Function
    End If
    
    'Convert inputs to uppercase and remote whitespace
    primer1 = SeqUtil.ProcessDna(primer1)
    primer2 = SeqUtil.ProcessDna(primer2)
    template = SeqUtil.ProcessDna(template)
    
    'Check that inputs are valid DNA sequences
    If Not (SeqUtil.IsValidDna(primer1) And SeqUtil.IsValidDna(primer2) And SeqUtil.IsValidDna(template)) Then
        PCR_V3 = "Error: One or more inputs not valid DNA sequence."
        Exit Function
    End If
    
    'Get reverse complement of template
    Dim reverseTemplate As String
    reverseTemplate = SeqUtil.ReverseComplement(template)
                        
    'Simulate dissociation of templates (if they exist, assume circular)
    Dim singleStrands As New Collection
    If Len(template) > 0 Then
        Call CollectionUtil.Add(singleStrands, template & template)
        Call CollectionUtil.Add(singleStrands, reverseTemplate & reverseTemplate)
    End If
    
    'Create another collection for storing previous round strands
    Dim oldStrands As New Collection
    
    'Simulate thermocycling
    Do While True
        'Add primers then anneal and polymerize
        Call CollectionUtil.Add(singleStrands, primer1)
        Call CollectionUtil.Add(singleStrands, primer2)
        Set singleStrands = SimulateAnneal(singleStrands)

        'If single strands is unchanged between rounds, exit loop
        If CollectionUtil.Equals(oldStrands, singleStrands) Then
            Exit Do
        End If
        
        'Update old strands
        Set oldStrands = New Collection
        Call CollectionUtil.Copy(singleStrands, oldStrands)
    Loop
    
    'Reverse complement the second primer
    Dim reverse2 As String
    reverse2 = SeqUtil.ReverseComplement(primer2)
    
    'Minimize the product length to prevent duplicate template (to handle circularity)
    Dim minLength As Integer
    minLength = Len(template) * 2 + Len(primer1) + Len(primer2)
    
    'Set default output
    PCR_V3 = "No PCR product generated."
    
    'Determine which of the species is the PCR product
    Dim seq As Variant
    For Each seq In singleStrands
        If StrComp(Left(seq, Len(primer1)), primer1) = 0 And StrComp(Right(seq, Len(reverse2)), reverse2) = 0 Then
            If Len(seq) < minLength Then
                PCR_V3 = seq
                minLength = Len(seq)
            End If
        End If
    Next
    
End Function

'Simulate annealing given a set of single strands
Private Function SimulateAnneal(ByVal singleStrands As Collection) As Collection
    Dim newStrands As New Collection
    Dim oligoA As Variant
    Dim oligoB As Variant
    Dim pdt As String
    
    'Iterate through each possible pair of strands and simulate annealing
    For Each oligoA In singleStrands
        For Each oligoB In singleStrands
            pdt = Anneal(oligoA, oligoB)
            If Not pdt = "" Then
                Call CollectionUtil.Add(newStrands, pdt)
            End If
        Next
    Next
    
    Set SimulateAnneal = newStrands
        
End Function

'Simulate annealing between two oligos
Private Function Anneal(ByVal oligoA As String, ByVal oligoB As String)
    'Reverse complement oligoB
    Dim reverseB As String
    reverseB = SeqUtil.ReverseComplement(oligoB)
    
    'If oligoA and oligoB are identical, terminate
    If StrComp(oligoA, oligoB) = 0 Then
        Anneal = ""
        Exit Function
    End If

    'If oligoA and oligoB are partner strands, terminate
    If StrComp(oligoA, reverseB) = 0 Then
        Anneal = ""
        Exit Function
    End If
    
    'Grab the last 6bp of oligoA
    Dim sixBP As String
    sixBP = Right(oligoA, 6)
    
    'Scan through and find the best annealing index if any
    Dim index As Integer
    index = 12 'index + 6 = 17, the first site it could be
    Dim bestSimilarity As Integer
    bestSimilarity = 17 '17 is one short of the cutoff of 18
    Dim bestIndex As Integer
    bestIndex = 0
    
    Do While True
        'Find the next 6bp match, exit loop if there are no more
        index = InStr(index + 1, reverseB, sixBP)
        If index = 0 Then
            Exit Do
        End If
        
        'Find the 30bp (or less) region of oligoA
        Dim startA As Integer
        startA = 1
        If Len(oligoA) > 30 Then
            startA = Len(oligoA) - 30 + 1
        End If
        Dim annealA As String
        annealA = Right(oligoA, Len(oligoA) - startA + 1)
        
        'Find the 30bp (or less) region ending in index+6
        Dim startB As Integer
        startB = index + 6 - 30
        If startB < 1 Then
            startB = 1
        End If
        Dim annealB As String
        annealB = Mid(reverseB, startB, index + 6 - startB)
        
        'Find the edit distance, taking into account that degenerate bases should not be counted as matches
        Dim distance As Integer
        distance = SeqUtil.EditDistance(annealA, annealB)
        
        'Calculate similarity based on edit distance
        Dim similarity As Integer
        similarity = 30 - distance
        
        'Find highest index with best similarity score, leading to shortest strand (to prevent duplicate templates from circularity)
        If similarity > bestSimilarity Then
            bestIndex = index
            bestSimilarity = similarity
        ElseIf similarity = bestSimilarity And index > bestIndex Then
            bestIndex = index
        End If
    Loop
    
    'If bestIndex was not updated, no 15+ annealing site was found
    If bestIndex < 1 Then
        Anneal = ""
        Exit Function
    End If
    
    'Construct extension product
    Anneal = oligoA & Right(reverseB, Len(reverseB) - (bestIndex + 6) + 1)
End Function

