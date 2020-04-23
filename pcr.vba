Function PCR(primer1 As String, primer2 As String, template As String)

    'Check that inputs are not empty
    If Not (Len(primer1) > 0 And Len(primer2) > 0 And Len(template) > 0) Then
            PCR = CVErr(xlErrValue)
            Exit Function
    End If
    
    'Convert inputs to uppercase and remote whitespace
    primer1 = Replace(Replace(UCase(primer1), Chr(32), ""), Chr(10), "")
    primer2 = Replace(Replace(UCase(primer2), Chr(32), ""), Chr(10), "")
    template = Replace(Replace(UCase(template), Chr(32), ""), Chr(10), "")
    
    'Check that inputs are valid DNA sequences
    If Not (IsValidDna(primer1) And IsValidDna(primer2) And IsValidDna(template)) Then
            PCR = CVErr(xlErrValue)
            Exit Function
    End If
    
    'Reverse complement primer 2 (reverse oligo)
    primer2 = ReverseComplement(primer2)
    
    'Find best annealing sites for template
    Dim forwardPrimer1AnnealSite As String
    Dim forwardPrimer2AnnealSite As String
    Dim forwardAnnealTotalLength As Integer
    forwardAnnealTotalLength = FindAnnealSites(primer1, primer2, template, forwardPrimer1AnnealSite, forwardPrimer2AnnealSite)
    
    'Find best annealing sites for reverse template
    Dim reverseTemplate As String
    reverseTemplate = ReverseComplement(template)
    
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
    template = RotateStringLeft(template, InStr(circularTemplate, finalPrimer1AnnealSite) - 1)
    
    Dim flankRegion As String
    flankRegion = Mid(template, Len(finalPrimer1AnnealSite) + 1, InStr(template, finalPrimer2AnnealSite) - Len(finalPrimer1AnnealSite) - 1)
    
    'Construct final PCR product
    PCR = primer1 & flankRegion & primer2

End Function

Function IsValidDna(inputStr As String) As Boolean
    Dim intPos As Integer
    IsValidDna = True
    For intPos = 1 To Len(inputStr)
        If Not Mid(inputStr, intPos, 1) Like WorksheetFunction.Rept("[ATCGRYSWKMBDHVN]", 1) Then
            IsValidDna = False
            Exit For
        End If
    Next
End Function

Function ReverseComplement(forwardStr As String) As String
    ReverseComplement = ""
    Dim intPos As Integer
    For intPos = 1 To Len(forwardStr)
        Select Case Mid(forwardStr, intPos, 1)
            Case "A"
                ReverseComplement = "T" & ReverseComplement
            Case "T"
                ReverseComplement = "A" & ReverseComplement
            Case "G"
                ReverseComplement = "C" & ReverseComplement
            Case "C"
                ReverseComplement = "G" & ReverseComplement
            Case Else
                ReverseComplement = "N" & ReverseComplement
        End Select
    Next
End Function

'Rotates string left (anti-clockwise) by d elements
Function RotateStringLeft(inputStr As String, d As Integer) As String
    RotateStringLeft = Right(inputStr, Len(inputStr) - d) & Left(inputStr, d)
End Function

Function FindAnnealSites(primer1 As String, primer2 As String, template As String, ByRef primer1AnnealSite As String, ByRef primer2AnnealSite As String) As Integer
    Dim minAnnealLength As Integer
    minAnnealLength = 16

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
        If primer1FoundIndex > 0 Then
            template = RotateStringLeft(template, primer1FoundIndex - 1)
            
            'Find best annealing site for primer2 (reverse oligo)
            For currPrimer2Length = Len(primer2) To minAnnealLength Step -1
                currPrimer2AnnealSite = Left(primer2, currPrimer2Length)
                primer2FoundIndex = InStr(template, currPrimer2AnnealSite)
                If primer2FoundIndex > 0 Then
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
