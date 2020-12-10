REM  *****  BASIC  *****


sub AcadNusxToUnicode

	dim document   as object
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
	
	Dim geo(0 To 32) As String
	geo(0) = "ა"
	geo(1) = "ბ"
	geo(2) = "გ"
	geo(3) = "დ"
	geo(4) = "ე"
	geo(5) = "ვ"
	geo(6) = "ზ"
	geo(7) = "თ"
	geo(8) = "ი"
	geo(9) = "კ"
	geo(10) = "ლ"
	geo(11) = "მ"
	geo(12) = "ნ"
	geo(13) = "ო"
	geo(14) = "პ"
	geo(15) = "ჟ"
	geo(16) = "რ"
	geo(17) = "ს"
	geo(18) = "ტ"
	geo(19) = "უ"
	geo(20) = "ფ"
	geo(21) = "ქ"
	geo(22) = "ღ"
	geo(23) = "ყ"
	geo(24) = "შ"
	geo(25) = "ჩ"
	geo(26) = "ც"
	geo(27) = "ძ"
	geo(28) = "წ"
	geo(29) = "ჭ"
	geo(30) = "ხ"
	geo(31) = "ჯ"
	geo(32) = "ჰ"
	
	document = ThisComponent.CurrentController.Frame
	
	For k = LBound(fonts) To UBound(fonts)
	    For i = LBound(eng) To UBound(eng)
	    	ReplaceChar(fonts(k), eng(i), geo(i))
		Next
	Next
	
	endtime = Now
	interval = endtime - starttime
	MsgBox ("Converting Completed! it took:" & interval & " seconds")

end sub


sub ReplaceChar(optional fontName as string, optional searchChar as string, optional replaceWithChar as string)
	Dim Attributes(0) As New com.sun.star.beans.PropertyValue
	Dim oReplace as Object
	
	Attributes(0).Name = "CharFontName"
	Attributes(0).Value = fontName

	oReplace = ThisComponent.createReplaceDescriptor()
	oReplace.setSearchString( searchChar )
	oReplace.setSearchAttributes(Attributes())
	oReplace.setReplaceString( replaceWithChar )
	oReplace.SearchRegularExpression = False
	oReplace.SearchCaseSensitive = True
	
	Writer_ReplaceAll_Paragraph_Breaks_With_Linefeeds = ThisComponent.replaceAll( oReplace )

end sub
