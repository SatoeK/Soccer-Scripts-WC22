'=====================================================================
' FSN 2008
' FOX Sports VBScript Misc module
'---------------------------------------------------------------------
' Etienne Laneville (Etienne.Laneville@FOX.com)
' September 2003
'=====================================================================

'

'=====================================================================
'	AUDIO
'---------------------------------------------------------------------

' Audio On:
'       Enable sounds
Sub Audio_On()

	' Set volume
	'Ventuz_Update ".Audio.Volume", 100
End Sub

' Audio Off:
'       Disable sounds
Sub Audio_Off()

	' Set volume
	'Ventuz_Update ".Audio.Volume", 0
End Sub

Function FormatOrdinal(inputNumber)
	Dim remainder
	Dim numberToFormat
	
	numberToFormat = inputNumber

	If inputNumber > 20 then
		numberToFormat = inputNumber Mod 10
	End if

	If numberToFormat > 4 And numberToFormat < 20 Then
		FormatOrdinal = inputNumber & "th"
		Exit Function
	End if
	
	remainder =  numberToFormat Mod 4

					
	Select Case remainder
		Case 1
			FormatOrdinal = inputNumber & "st"
			'FormatOrdinal = inputNumber & "¢"
		Case 2 
			FormatOrdinal = inputNumber & "nd"
			'FormatOrdinal = inputNumber & "§"
		Case 3 
			FormatOrdinal = inputNumber & "rd"
			'FormatOrdinal = inputNumber & "£"
		Case Else
			FormatOrdinal = inputNumber & "th"
			'FormatOrdinal = inputNumber & "¶"
	End Select
End Function


