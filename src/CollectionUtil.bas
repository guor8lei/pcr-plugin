Attribute VB_Name = "CollectionUtil"
'Returns true if a col contains val
Public Function Contains(ByVal col As Collection, ByVal val As Variant) As Boolean
    Contains = False
    For Each Item In col
        If Item = val Then
            Contains = True
            Exit Function
        End If
    Next
End Function
'Mutatively adds val to col
Public Sub Add(ByRef col As Collection, ByVal val As Variant)
    If Not Contains(col, val) Then
        col.Add val
    End If
End Sub
'Mutatively removes val from col
Public Sub Remove(ByRef col As Collection, ByVal val As Variant)
    Dim newCol As New Collection
    For Each Item In col
        If Not Item = val Then
            Call Add(newCol, Item)
        End If
    Next
    Set col = newCol
End Sub
'Returns true if col1 and col2 contain the same items
Public Function Equals(ByVal col1 As Collection, ByVal col2 As Collection) As Boolean
    For Each Item In col1
        If Not Contains(col2, Item) Then
            Equals = False
            Exit Function
        End If
    Next

    For Each Item In col2
        If Not Contains(col1, Item) Then
            Equals = False
            Exit Function
        End If
    Next
    Equals = True
End Function
'Make toCol into a copy of fromCol, with all the same items
Public Sub Copy(ByVal fromCol As Collection, ByRef toCol As Collection)
    Set toCol = New Collection
    For Each Item In fromCol
        Call Add(toCol, Item)
    Next
End Sub
