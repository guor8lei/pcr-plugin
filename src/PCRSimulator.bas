Attribute VB_Name = "PCRSimulator"
'Working version of PCRSimulator
'Uses a combination of approaches in PCRSimulatorV2 and V4 to achieve the most accurate and efficient way to simulate PCR products

Public Function PCR(ByVal primer1 As String, ByVal primer2 As String, Optional ByVal template As String = "")
    'If template does not exist, use V4, else use V2
    If Len(template) > 0 Then
        PCR = PCR_V2(primer1, primer2, template)
        Exit Function
    Else
        PCR = PCR_V4(primer1, primer2, "")
    End If
End Function
