
Sub AcadNusxToUnicode()
    Dim starttime As Date
    Dim endtime As Date
    starttime = Now
    Dim fonts(0 To 2) As String
    fonts(0) = "AcadNusx"
    fonts(1) = "AcadMtavr"
    fonts(2) = "LitNusx"

    Dim eng(0 To 32) As String
    eng(0) = "a"
    eng(1) = "b"
    eng(2) = "g"
    eng(3) = "d"
    eng(4) = "e"
    eng(5) = "v"
    eng(6) = "z"
    eng(7) = "T"
    eng(8) = "i"
    eng(9) = "k"
    eng(10) = "l"
    eng(11) = "m"
    eng(12) = "n"
    eng(13) = "o"
    eng(14) = "p"
    eng(15) = "J"
    eng(16) = "r"
    eng(17) = "s"
    eng(18) = "t"
    eng(19) = "u"
    eng(20) = "f"
    eng(21) = "q"
    eng(22) = "R"
    eng(23) = "y"
    eng(24) = "S"
    eng(25) = "C"
    eng(26) = "c"
    eng(27) = "Z"
    eng(28) = "w"
    eng(29) = "W"
    eng(30) = "x"
    eng(31) = "j"
    eng(32) = "h"

    Dim geo(0 To 32) As Integer
    geo(0) = 4304
    geo(1) = 4305
    geo(2) = 4306
    geo(3) = 4307
    geo(4) = 4308
    geo(5) = 4309
    geo(6) = 4310
    geo(7) = 4311
    geo(8) = 4312
    geo(9) = 4313
    geo(10) = 4314
    geo(11) = 4315
    geo(12) = 4316
    geo(13) = 4317
    geo(14) = 4318
    geo(15) = 4319
    geo(16) = 4320
    geo(17) = 4321
    geo(18) = 4322
    geo(19) = 4323
    geo(20) = 4324
    geo(21) = 4325
    geo(22) = 4326
    geo(23) = 4327
    geo(24) = 4328
    geo(25) = 4329
    geo(26) = 4330
    geo(27) = 4331
    geo(28) = 4332
    geo(29) = 4333
    geo(30) = 4334
    geo(31) = 4335
    geo(32) = 4336

    For k = LBound(fonts) To UBound(fonts)
        For i = LBound(eng) To UBound(eng)
            Selection.Find.ClearFormatting
            Selection.Find.Font.Name = fonts(k)
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = eng(i)
               .Replacement.Text = ChrW(geo(i))
               .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
        Next
    Next
    endtime = Now
    interval = endtime - starttime
    MsgBox ("Converting Completed! It took: " & interval & " seconds")

End Sub


