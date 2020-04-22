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
    
    PCR = ReverseComplement(primer1)

End Function

Function IsValidDna(inputStr As String) As Boolean
    IsValidDna = inputStr Like WorksheetFunction.Rept("[ATCGRYSWKMBDHVN]", Len(inputStr))
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
