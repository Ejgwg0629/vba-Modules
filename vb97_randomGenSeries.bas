Attribute VB_Name = "vb97_randomGenSeries"
'********************************************************************************
' Function return a random string
' with the optional length parameter
'********************************************************************************
Function ranStr(Optional ByVal Length As Integer = 1)
    myAry = Array("A", "B", "C", "D", "E", "F", "G", "H", _
        "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
        "S", "T", "U", "V", "W", "X", "Y", "Z")
    Randomize
    tempStr = ""
    For i = 1 To Length
        tempStr = tempStr & myAry(Int((25 * Rnd) + 1))
    Next
    ranStr = tempStr
End Function


'********************************************************************************
' Function return a random time
' with the optional format parameter
'********************************************************************************
Function ranTime(Optional ByVal fmt As String = "hh:mm:ss")
    Randomize
    ranTime = format(CDate(Rnd), fmt)
End Function


'********************************************************************************
' Function return a random date
' with the optional format parameter
'********************************************************************************
Function ranDate(Optional ByVal fmt As String = "yyyyMMdd")
    Randomize
    ranDate = format(Rnd * (CLng(Date) - 1) + 1, fmt)
End Function

Public Function ranColor()
    arr = Array("&hf0f8f8", "&hf2f8f8", "&h3e4849", "&h282828", "&h74dbe6", "&hefd966", "&h7226f9", _
                "&hff81ae", "&h5e7175", "&h1f97fd", "&h69d5ff", "&h2ee2a6", "&h2f9b52", "&hc6c8c5", _
                "&h211f1d", "&hfecb96", "&h60ffa8", "&h62c0e9", "&h6666cc", "&hfec5c6", "&h99cc99", _
                "&hededed", "&h98eef9", "&h85d0da", "&hfeb162", "&hfecb96", "&hfecb96", "&hfec5c6", _
                "&h423a38", "&h776c69", "&ha7a1a0", "&hbc8401", "&hf27840", "&ha426a6", "&h4fa150", _
                "&h4956e4", "&h4312ca", "&h016898", "&h0184c1", "&hfafafa", "&hff6e52", "&h211f1d", _
                "&h2e2a28", "&h413b37", "&h969896", "&hb4b7b4", "&hc6c8c5", "&he0e0e0", "&hffffff", _
                "&h6666cc", "&h5f93de", "&h74c6f0", "&h68bdb5", "&hb7be8a", "&hbea281", "&hbb94b2", _
                "&h5a68a3", "&h362b00", "&h423607", "&h756e58", "&h837b65", "&h969483", "&ha1a193", _
                "&hd5e8ee", "&he3f6fd", "&h0089b5", "&h164bcb", "&h2f32dc", "&h8236d3", "&hc4716c", _
                "&hd28b26", "&h98a12a", "&h009985", "&hefaf5f", "&hffe9cf", "&h9f6f2f", "&hdfcbaa", _
                "&hffffff", "&h000000", "&h3232f9", "&h5049d4", "&he1e1e1", "&hadadad", "&h777777", _
                "&h333333", "&h369d69", "&h177244", "&hb90067", "&h7f0047", "&h1aade9", "&h0f67bc", _
                "&h98a12a", "&hf8f8f8", "&h535353", "&he5e5e5", "&hc0c0c0", "&h969896", "&h333333", _
                "&h211f1d", "&hffd6b0", "&hfecb96", "&hb38600", "&h913618", "&h813e1d", "&h808000", _
                "&ha35d79", "&h173a69", "&h5d1da7", "&heaffea", "&h60ffa8", "&h5ca363", "&h32a555", _
                "&hececff", "&h6666cc", "&h002cbd", "&h1d2ab5", "&h62c0e9", "&h436aed")
    Randomize
    count = UBound(arr)
    ranColor = CLng(arr(Rnd * count))
End Function
