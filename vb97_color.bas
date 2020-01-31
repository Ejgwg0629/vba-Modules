Attribute VB_Name = "vb97_color"
Public Function vbaColorToRGB(ByVal lColor As Long) As String
    Dim iR, iG, iB As Long
    iR = (lColor Mod 256)
    iG = (lColor \ 256) Mod 256
    iB = (lColor \ 65536) Mod 256
    vbaColorToRGB = "(" & iR & "," & iG & "," & iB & ")"
End Function
