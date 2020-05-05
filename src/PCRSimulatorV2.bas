Attribute VB_Name = "PCRSimulatorV2"
'Version 2 of PCRSimulator
'Simulates annealing by finding 6bp matches, and building up
'Based on Professor Anderson's PCRSimulator class in ConstructionFileSimulator
'Works well with most cases (but needs template)
Public Function PCR_V2(ByVal primer1 As String, ByVal primer2 As String, ByVal template As String) As String
    'Run simulation with both possible template directions
    PCR_V2 = DirectionalPCR(primer1, primer2, template, False)
    If Left(PCR_V2, 5) = "Error" Then
        PCR_V2 = DirectionalPCR(primer1, primer2, template, True)
    End If
End Function

'Simulates PCR product given a template direction
Private Function DirectionalPCR(ByVal oligo1 As String, ByVal oligo2 As String, ByVal template As String, ByVal templateRC As Boolean) As String

    'Define constants
    MAXINT = (2 ^ 15) - 1

    'Check that primer inputs are not empty
    If Not (Len(oligo1) > 0 And Len(oligo2) > 0) Then
        DirectionalPCR = "Error: One or more primers does not exist."
        Exit Function
    End If
    
    'If template is empty, should use PCR instead
    If Len(template) = 0 Then
        DirectionalPCR = "Error: Template does not exist. Use PCR_V3 instead of PCR_V2."
        Exit Function
    End If
    
    'Convert inputs to uppercase and remote whitespace
    oligo1 = SeqUtil.ProcessDna(oligo1)
    oligo2 = SeqUtil.ProcessDna(oligo2)
    template = SeqUtil.ProcessDna(template)
    
    'Check that inputs are valid DNA sequences
    If Not (SeqUtil.IsValidDna(oligo1) And SeqUtil.IsValidDna(oligo2) And SeqUtil.IsValidDna(template)) Then
        DirectionalPCR = "Error: One or more inputs not valid DNA sequence."
        Exit Function
    End If
    
    'If templateRC is true, reverse complement template
    If templateRC Then
        template = SeqUtil.ReverseComplement(template)
    End If
    
    'Duplicate template to take simulate circulate template
    Dim circularTemplate As String
    circularTemplate = template & template
    
    'Find indices of sequences starting with the last 6 bases 5' to 3' of oligo1 (or the first 6 bases 3' to 5' rev comp of the sequence)
    Dim lastSixOligo1Seq As String
    lastSixOligo1Seq = Right(oligo1, 6) 'Last six bases of 5' to 3' oligo1
    
    Dim revFirstSixOligo1Seq As String
    revFirstSixOligo1Seq = Left(SeqUtil.ReverseComplement(oligo1), 6) 'First six bases of 3' to 5' revcompOligo1
    
    Dim forwardIndicesOligo1 As New Collection 'Last indices of 5' to 3' matching sequences
    Dim revIndicesOligo1 As New Collection 'First indices of 3' to 5' reverse complement matching sequences
    
    'If the six bp matching sequence is found starting at index i, add the index of the last 5' to 3' bp (i+6)
    Dim i As Integer
    i = InStr(circularTemplate, lastSixOligo1Seq)
    
    Do While i <> 0
        Call CollectionUtil.Add(forwardIndicesOligo1, i + 6)
        i = InStr(i + 1, circularTemplate, lastSixOligo1Seq)
    Loop
    
    'If the six bp matching revcomp sequence is found starting at index i, add the index of the first 3' to 5' bp (i)
    i = InStr(circularTemplate, revFirstSixOligo1Seq)
    
    Do While i <> 0
        Call CollectionUtil.Add(revIndicesOligo1, i)
        i = InStr(i + 1, circularTemplate, revFirstSixOligo1Seq)
    Loop
    
    'Check that a 6 bp match was found
    If forwardIndicesOligo1.Count = 0 And revIndicesOligo1.Count = 0 Then
        DirectionalPCR = "Error: There was no matching 6 bp sequence for the provided oligo1."
        Exit Function
    End If
    
    'forwardIndicesOligo1 contains the indices (on the 5' to 3' template sequence) of the END of 5' to 3' sequences potentially matching oligo1
    'revIndicesOligo1 contains the indices (on the 5' to 3' template sequence) of the FIRST index of 3' to 5' sequences potentially matching revOligo1
    
    'Will be assigned to the END index (on the 5' to 3' template seq) of the 5' to 3' sequence that best matches oligo1
    Dim lowestForwardEditDistanceOligo1SeqIndex As Integer
    lowestForwardEditDistanceOligo1SeqIndex = 1
    
    Dim lowestForwardEditDistanceOligo1 As Integer
    lowestForwardEditDistanceOligo1 = MAXINT
    
    Dim j As Variant
    For Each j In forwardIndicesOligo1
        offset = j - Len(oligo1) 'Offset assigned to the index of the FIRST index of the 5' to 3' sequence potentially matching oligo1
        tempRotatedTemplateSeq = SeqUtil.RotateString(circularTemplate, offset) 'Rotate the string such that the first bp of the 5' to 3' matching seq is at index 0
        potentialMatchSeq = Mid(tempRotatedTemplateSeq, Application.WorksheetFunction.Max(1, Len(oligo1) - 24), 25) 'Select the last 25 bases of the potential matching seq, or all bases if less than 25 bp long
        annealingRegionOligo1 = Right(oligo1, 25) 'Select the last 25 bases of oligo1, or all bases if less than 25 bp long
        tempEditDistance = SeqUtil.EditDistance(potentialMatchSeq, annealingRegionOligo1)
        If tempEditDistance < lowestForwardEditDistanceOligo1 Then
            lowestForwardEditDistanceOligo1SeqIndex = j
            lowestForwardEditDistanceOligo1 = tempEditDistance
        End If
    Next
    
    'Will be assigned to the FIRST index (on the 5' to 3' template seq) of the 3' to 5' sequence that best matches oligo1
    Dim lowestRevEditDistanceOligo1SeqIndex As Integer
    lowestRevEditDistanceOligo1SeqIndex = 1
    
    Dim lowestRevEditDistanceOligo1 As Integer
    lowestRevEditDistanceOligo1 = MAXINT
    
    For Each j In revIndicesOligo1
        tempRotatedTemplateSeq = SeqUtil.RotateString(circularTemplate, j) 'Rotate such that the matching 3' to 5' seq head is at index 0
        potentialMatchSeq = Mid(tempRotatedTemplateSeq, Application.WorksheetFunction.Max(1, Len(oligo1) - 24), 25) 'Select the last 25 bases, or all bases if less than 25 bp long
        revPotentialMatchSeq = SeqUtil.ReverseComplement(potentialMatchSeq) 'Reverse complement the 5' to 3' sequence to find the potentially matching 3' to 5' seq
        revOligo1Seq = SeqUtil.ReverseComplement(oligo1) 'Reverse complement oligo1
        revAnnealingRegionOligo1 = Left(revOligo1Seq, 25) 'Select the last 25 bases of rev comp oligo1, or all bases if less than 25 bp long
        tempEditDistance = SeqUtil.EditDistance(revPotentialMatchSeq, revAnnealingRegionOligo1)
        If tempEditDistance < lowestRevEditDistanceOligo1 Then
            lowestRevEditDistanceOligo1SeqIndex = j
            lowestRevEditDistanceOligo1 = tempEditDistance
        End If
    Next
    
    'lowestForwardEditDistanceOligo1SeqIndex is now assigned to the END index (on the 5' to 3' template sequence) of the 5' to 3' sequence that has the best match with oligo1
    'lowestRevEditDistanceOligo1SeqIndex is now assigned to the FIRST index (on the 5' to 3' template sequence) of the 3' to 5' sequence that has the best match with oligo1

    Dim oligo1Forward As Boolean
    oligo1Forward = True
    
    'If the forward direction has a lower edit distance, rotate the template such that oligo1 is at the very end
    'Otherwise, rotate the template such that revcomp of oligo1 ends at the first index, then rev comp the entire template seq such that the revcomp of oligo1 starts at the very end
    If lowestForwardEditDistanceOligo1 < lowestRevEditDistanceOligo1 Then
        circularTemplate = SeqUtil.RotateString(circularTemplate, lowestForwardEditDistanceOligo1SeqIndex)
    Else
        circularTemplate = SeqUtil.RotateString(circularTemplate, lowestRevEditDistanceOligo1SeqIndex)
        circularTemplate = SeqUtil.ReverseComplement(circularTemplate)
        oligo1Forward = False
    End If
    
    Dim endIndexOligo1 As Integer
    endIndexOligo1 = 1
    
    'After the above steps, the template sequence should now either end with oligo1 or with the reverse complement of oligo1, and the end index of the oligo1 seq is saved as endIndexOligo1
    
    'Find indices of sequences starting with the first 6 bases of oligo2 (or the rev comp of the sequence)
    Dim lastSixOligo2Seq As String
    lastSixOligo2Seq = Right(oligo2, 6) 'Last six bases of 5' to 3' oligo2
    
    Dim revFirstSixOligo2Seq As String
    revFirstSixOligo2Seq = Left(SeqUtil.ReverseComplement(oligo2), 6) 'First six bases of 3' to 5' revcompOligo2
    
    Dim forwardIndicesOligo2 As New Collection 'Last indices of 5' to 3' matching sequences
    Dim revIndicesOligo2 As New Collection 'First indices of 3' to 5' reverse complement matching sequences
    
    ''If the six bp matching sequence is found starting at index i, add the index of the last 5' to 3' bp (i+6)
    i = InStr(circularTemplate, lastSixOligo2Seq)
    
    Do While i <> 0
        Call CollectionUtil.Add(forwardIndicesOligo2, i + 6)
        i = InStr(i + 1, circularTemplate, lastSixOligo2Seq)
    Loop
    
    'If the six bp matching revcomp sequence is found starting at index i, add the index of the first 3' to 5' bp (i)
    i = InStr(circularTemplate, revFirstSixOligo2Seq)
    
    Do While i <> 0
        Call CollectionUtil.Add(revIndicesOligo2, i)
        i = InStr(i + 1, circularTemplate, revFirstSixOligo2Seq)
    Loop
    
    'Check that a 6 bp match was found
    If forwardIndicesOligo2.Count = 0 And revIndicesOligo2.Count = 0 Then
        DirectionalPCR = "Error: There was no matching 6 bp sequence for the provided oligo2."
        Exit Function
    End If
    
    'Will be assigned to the END index (on the 5' to 3' template seq) of the 5' to 3' sequence that best matches oligo2
    Dim lowestForwardEditDistanceOligo2SeqIndex As Integer
    lowestForwardEditDistanceOligo2SeqIndex = 1
    
    Dim lowestForwardEditDistanceOligo2 As Integer
    lowestForwardEditDistanceOligo2 = MAXINT
    
    For Each j In forwardIndicesOligo2
        potentialMatchSeq = Mid(circularTemplate, Application.WorksheetFunction.Max(1, Len(oligo2) - 24), 25) 'Select the last 25 bases of the potential matching seq, or all bases if less than 25 bp long
        annealingRegionOligo2 = Right(oligo2, 25) 'Select the last 25 bases of oligo2, or all bases if less than 25 bp long
        tempEditDistance = SeqUtil.EditDistance(potentialMatchSeq, annealingRegionOligo2)
        If tempEditDistance < lowestForwardEditDistanceOligo2 Then
            lowestForwardEditDistanceOligo2SeqIndex = j
            lowestForwardEditDistanceOligo2 = tempEditDistance
        End If
    Next
    
    'Will be assigned to the FIRST index (on the 5' to 3' template seq) of the 3' to 5' sequence that best matches oligo2
    Dim lowestRevEditDistanceOligo2SeqIndex As Integer
    lowestRevEditDistanceOligo2SeqIndex = 1
    
    Dim lowestRevEditDistanceOligo2 As Integer
    lowestRevEditDistanceOligo2 = MAXINT
    
    For Each j In revIndicesOligo2
        potentialMatchSeq = Mid(circularTemplate, j, 25) 'Select the last 25 bases of the potential matching seq, or all bases if less than 25 bp long
        revOligo2Seq = SeqUtil.ReverseComplement(oligo2)
        revAnnealingRegionOligo2 = Left(revOligo2Seq, 25) 'Select the last 25 bases of rev comp oligo2, or all bases if less than 25 bp long
        tempEditDistance = SeqUtil.EditDistance(potentialMatchSeq, revAnnealingRegionOligo2)
        If tempEditDistance < lowestRevEditDistanceOligo2 Then
            lowestRevEditDistanceOligo2SeqIndex = j
            lowestRevEditDistanceOligo2 = tempEditDistance
        End If
    Next

    Dim oligo2Forward As Boolean
    oligo2Forward = True
    
    Dim startIndexOligo2 As Integer
    
    If lowestForwardEditDistanceOligo2 < lowestRevEditDistanceOligo2 Then
        startIndexOligo2 = lowestForwardEditDistanceOligo2SeqIndex - Len(oligo2)
    Else
        startIndexOligo2 = lowestRevEditDistanceOligo2SeqIndex
        oligo2Forward = False
    End If
    
    'StartIndexOligo2 should now be assigned to the FIRST index (on the 5' to 3' template) of the sequence that best matches oligo2 (whether 5' to 3' OR 3' to 5')
    
    'Check if the oligos will anneal in the same directions
    If oligo1Forward = oligo2Forward Then
        DirectionalPCR = "Error: The provided oligos will both anneal in the same direction."
        Exit Function
    End If
    
    'Region between the two oligos that does not contain bp from either
    Dim flankedRegion As String
    flankedRegion = Mid(circularTemplate, endIndexOligo1, startIndexOligo2 - endIndexOligo1)
    
    DirectionalPCR = oligo1 & flankedRegion & SeqUtil.ReverseComplement(oligo2)
    
End Function
