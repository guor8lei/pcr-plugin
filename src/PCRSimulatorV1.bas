Attribute VB_Name = "PCRSimulatorV1"
'Version 1 of PCRSimulator
'Simulates annealing by finding the longest possible anneal sites with exact match
'Works well with most cases (but needs template), but does not accurately simulate anneal in the correct biological sense
Public Function PCR_V1(ByVal primer1 As String, ByVal primer2 As String, ByVal template As String)

    'Check that primer inputs are not empty
    If Not (Len(primer1) > 0 And Len(primer2) > 0) Then
        PCR_V1 = "Error: One or more primers does not exist."
        Exit Function
    End If
    
    'If template is empty, should use PCR instead
    If Len(template) = 0 Then
        PCR_V1 = "Error: Template does not exist. Use PCR_V3 instead of PCR_V1."
        Exit Function
    End If
    
    'Convert inputs to uppercase and remote whitespace
    primer1 = SeqUtil.ProcessDna(primer1)
    primer2 = SeqUtil.ProcessDna(primer2)
    template = SeqUtil.ProcessDna(template)

    'Check that inputs are valid DNA sequences
    If Not (SeqUtil.IsValidDna(primer1) And SeqUtil.IsValidDna(primer2) And SeqUtil.IsValidDna(template)) Then
        PCR_V1 = "Error: One or more inputs not valid DNA sequence."
        Exit Function
    End If
    
    'Reverse complement primer 2 (reverse oligo)
    primer2 = SeqUtil.ReverseComplement(primer2)
    
    'Find best annealing sites for template
    Dim forwardPrimer1AnnealSite As String
    Dim forwardPrimer2AnnealSite As String
    Dim forwardAnnealTotalLength As Integer
    forwardAnnealTotalLength = FindAnnealSites(primer1, primer2, template, forwardPrimer1AnnealSite, forwardPrimer2AnnealSite)
    
    'Get reverse complement of template
    Dim reverseTemplate As String
    reverseTemplate = SeqUtil.ReverseComplement(template)
    
    'Find best annealing sites for reverse template
    Dim reversePrimer1AnnealSite As String
    Dim reversePrimer2AnnealSite As String
    Dim reverseAnnealTotalLength As Integer
    reverseAnnealTotalLength = FindAnnealSites(primer1, primer2, reverseTemplate, reversePrimer1AnnealSite, reversePrimer2AnnealSite)
    
    'Compare template vs reverse template anneal sites, determine final anneal sites
    Dim finalPrimer1AnnealSite As String
    Dim finalPrimer2AnnealSite As String
    
    If reverseAnnealTotalLength > forwardAnnealTotalLength Then
        template = reverseTemplate
        finalPrimer1AnnealSite = reversePrimer1AnnealSite
        finalPrimer2AnnealSite = reversePrimer2AnnealSite
    Else
        finalPrimer1AnnealSite = forwardPrimer1AnnealSite
        finalPrimer2AnnealSite = forwardPrimer2AnnealSite
    End If
    
    'Find flank region based on final anneal sites
    Dim circularTemplate As String
    circularTemplate = template & template
    template = SeqUtil.RotateStringLeftBy(template, InStr(circularTemplate, finalPrimer1AnnealSite) - 1)
    
    Dim flankRegion As String
    flankRegion = Mid(template, Len(finalPrimer1AnnealSite) + 1, InStr(template, finalPrimer2AnnealSite) - Len(finalPrimer1AnnealSite) - 1)
    
    'Construct final PCR product
    PCR_V1 = primer1 & flankRegion & primer2
End Function

'Used by PCR_V1 to find anneal sites by starting with longest possible annealing site and decreasing length until match is found
Private Function FindAnnealSites(ByVal primer1 As String, ByVal primer2 As String, ByVal template As String, ByRef primer1AnnealSite As String, ByRef primer2AnnealSite As String) As Integer
    Dim minAnnealLength As Integer
    minAnnealLength = 6

    Dim circularTemplate As String
    circularTemplate = template & template
    
    Dim currPrimer1Length As Integer
    Dim currPrimer2Length As Integer
    Dim currPrimer1AnnealSite As String
    Dim currPrimer2AnnealSite As String
    Dim primer1FoundIndex As Integer
    Dim primer2FoundIndex As Integer
    Dim annealTotalLength As Integer
    
    'Find best annealing site for primer1 (forward oligo)
    For currPrimer1Length = Len(primer1) To minAnnealLength Step -1
        If currPrimer1Length + Len(primer2) < annealTotalLength Then
            Exit For
        End If
    
        currPrimer1AnnealSite = Right(primer1, currPrimer1Length)
        primer1FoundIndex = InStr(circularTemplate, currPrimer1AnnealSite)
        If primer1FoundIndex > 0 And Not SeqUtil.IsDegenerateDna(currPrimer1AnnealSite) Then
            template = SeqUtil.RotateStringLeftBy(template, primer1FoundIndex - 1)
            
            'Find best annealing site for primer2 (reverse oligo)
            For currPrimer2Length = Len(primer2) To minAnnealLength Step -1
                currPrimer2AnnealSite = Left(primer2, currPrimer2Length)
                primer2FoundIndex = InStr(template, currPrimer2AnnealSite)
                If primer2FoundIndex > 0 And Not SeqUtil.IsDegenerateDna(currPrimer2AnnealSite) Then
                    If currPrimer1Length + currPrimer2Length > annealTotalLength Then
                        annealTotalLength = currPrimer1Length + currPrimer2Length
                        primer1AnnealSite = currPrimer1AnnealSite
                        primer2AnnealSite = currPrimer2AnnealSite
                    End If
                End If
            Next
        End If
    Next
    
    FindAnnealSites = annealTotalLength
    
End Function

