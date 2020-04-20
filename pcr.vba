Function PCR(primer1, primer2, template)

If VarType(primer1) <> vbString Or VarType(primer2) <> vbString Or VarType(template) <> vbString Then
    MsgBox "Wrong input data type; should be String"
End If

PCR = primer1 & " " & primer2

End Function
