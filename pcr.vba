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
    
    'Find best annealing site for forward oligo
    Dim circularTemplate As String
    circularTemplate = template & template
    
    Dim currLength As Integer
    Dim bestForwardAnnealSite As String
    Dim currAnnealSite As String
    Dim foundIndex As Integer
    
    For currLength = Len(primer1) To 1 Step -1
        currAnnealSite = Right(primer1, currLength)
        foundIndex = InStr(circularTemplate, currAnnealSite)
        If foundIndex > 0 Then
            bestForwardAnnealSite = currAnnealSite
            template = RotateStringLeft(template, foundIndex - 1)
            Exit For
        End If
    Next
    
    If Len(bestForwardAnnealSite) = 0 Then
            MsgBox "Annealing site not found."
            Exit Function
    End If
    
    PCR = bestForwardAnnealSite
    MsgBox template

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
