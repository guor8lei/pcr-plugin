Attribute VB_Name = "SeqUtil"
'Convert inputs to uppercase and remote whitespace
Public Function ProcessDna(ByVal inputDna As String) As String
    ProcessDna = Replace(Replace(UCase(inputDna), Chr(32), ""), Chr(10), "")
End Function

'Check that input is a valid DNA sequence
Public Function IsValidDna(ByVal inputStr As String) As Boolean
    Dim intPos As Integer
    inputStr = ProcessDna(inputStr)
    IsValidDna = True
    For intPos = 1 To Len(inputStr)
        If Not Mid(inputStr, intPos, 1) Like WorksheetFunction.Rept("[ATCGRYSWKMBDHVN]", 1) Then
            IsValidDna = False
            Exit For
        End If
    Next
End Function

'Check that input is a degenerate DNA sequence
Public Function IsDegenerateDna(ByVal inputStr As String) As Boolean
    Dim intPos As Integer
    inputStr = ProcessDna(inputStr)
    IsDegenerateDna = False
    For intPos = 1 To Len(inputStr)
        If Not Mid(inputStr, intPos, 1) Like WorksheetFunction.Rept("[ATCG]", 1) Then
            IsDegenerateDna = True
            Exit For
        End If
    Next
End Function

'Count number of degenerate base pairs in a DNA sequence
Public Function CountDegenerateBases(ByVal inputStr As String) As Integer
    Dim intPos As Integer
    inputStr = ProcessDna(inputStr)
    CountDegenerateBases = 0
    For intPos = 1 To Len(inputStr)
        If Not Mid(inputStr, intPos, 1) Like WorksheetFunction.Rept("[ATCG]", 1) Then
            CountDegenerateBases = CountDegenerateBases + 1
        End If
    Next
End Function

'Rotates string left (anti-clockwise) by d elements
Public Function RotateStringLeftBy(ByVal inputStr As String, ByVal d As Integer) As String
    RotateStringLeftBy = Right(inputStr, Len(inputStr) - d) & Left(inputStr, d)
End Function

'Rotate string to desired offset
Public Function RotateString(ByVal inputStr As String, ByVal offset As Integer) As String
    offset = offset Mod Len(inputStr)
    If offset < 1 Then
        offset = offset + Len(inputStr)
    End If
    RotateString = Right(inputStr, Len(inputStr) - offset + 1) & Left(inputStr, offset - 1)
End Function

'Returns reverse complement of DNA sequence
Public Function ReverseComplement(ByVal forwardStr As String) As String
    ReverseComplement = ""
    forwardStr = ProcessDna(forwardStr)
    If Not IsValidDna(forwardStr) Then
        ReverseComplement = CVErr(xlErrValue)
        Exit Function
    End If
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
            Case "B"
                ReverseComplement = "V" & ReverseComplement
            Case "D"
                ReverseComplement = "H" & ReverseComplement
            Case "H"
                ReverseComplement = "D" & ReverseComplement
            Case "K"
                ReverseComplement = "M" & ReverseComplement
            Case "M"
                ReverseComplement = "K" & ReverseComplement
            Case "R"
                ReverseComplement = "Y" & ReverseComplement
            Case "S"
                ReverseComplement = "S" & ReverseComplement
            Case "V"
                ReverseComplement = "B" & ReverseComplement
            Case "W"
                ReverseComplement = "W" & ReverseComplement
            Case "Y"
                ReverseComplement = "R" & ReverseComplement
            Case Else
                ReverseComplement = "N" & ReverseComplement
        End Select
    Next
End Function

'Compute edit distance between two Strings using the Smith-Waterman Algorithm
'Does not count degenerate base pairs
Public Function EditDistance(ByVal s1 As String, ByVal s2 As String) As Integer
    Dim s1Len As Integer
    Dim s2Len As Integer
    
    s1Len = Len(s1)
    s2Len = Len(s2)
    
    Dim DistArray() As Integer
    ReDim DistArray(s1Len, s2Len) As Integer

    Dim i As Integer
    Dim j As Integer

    For i = 0 To s1Len
        For j = 0 To s2Len
            If i = 0 Then
                DistArray(i, j) = j
            ElseIf j = 0 Then
                DistArray(i, j) = i
            ElseIf Mid(s1, i, 1) = Mid(s2, j, 1) And Mid(s1, i, 1) Like WorksheetFunction.Rept("[ATCG]", 1) Then
                DistArray(i, j) = DistArray(i - 1, j - 1)
            Else
                DistArray(i, j) = 1 + Application.WorksheetFunction.Min(DistArray(i, j - 1), DistArray(i - 1, j), DistArray(i - 1, j - 1))
            End If
        Next j
    Next i

    EditDistance = DistArray(s1Len, s2Len)
End Function

