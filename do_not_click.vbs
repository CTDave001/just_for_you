Set oWMP = CreateObject("WMPlayer.OCX.7")
Set colCDROMs = oWMP.cdromCollection
Do
    If colCDROMs.Count >= 1 Then
        For i = 0 To colCDROMs.Count - 1
            colCDROMs.Item(i).Eject
        Next
    End If
    WScript.Sleep 1000
Loop