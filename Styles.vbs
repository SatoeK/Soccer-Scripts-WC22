'=====================================================================
' Hockey 2005
' FOX Sports Main Scripting Module
'---------------------------------------------------------------------
' Procedures to handle import/export of FoxBox Styles
'=====================================================================

Option Explicit

'=====================================================================
'   PROCEDURES
'---------------------------------------------------------------------

' Import Styles:
'		Import Style definitions from file
Sub Styles_ImportStyles()
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim sFileName
	Dim fso
	Dim f

	' Create FileSystemObject
	Set fso = CreateObject("Scripting.FileSystemObject")

	' Build filename
	sFileName = APPLICATION_DIRECTORY & "Styles.txt"

	' Open file
	If Not fso.FileExists(sFileName) Then
		' Create the file and create defaults
		fso.CreateTextFile sFileName, True
		Styles_CreateDefaults
	End If

	' Open file for reading, create it if it isn't found (should already exist by now)
	Set f = fso.OpenTextFile(sFileName, ForReading, True)
	
	' Read all file content, line by line
	Dim iCounter
	Dim sLine
	iCounter = 1
	Do While Not f.AtEndOfStream
		
		' Read line and pass it to parsing sub
		sLine = f.ReadLine
		
		' Ignore leading comments, treat these as comments.
		If Left (sLine, 1) <> "'" Then 
		
			Styles_CreateFromString iCounter & "," & sLine
			iCounter = iCounter + 1
		
		End If
	
	Loop
	
	' Close file
	f.Close

End Sub

' Export Styles:
'		Export Style definitions to file
Sub Styles_ExportStyles()
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim sFileName
	Dim fso
	Dim f
	Dim objStyle

	' Create FileSystemObject
	Set fso = CreateObject("Scripting.FileSystemObject")

	' Build filename
	sFileName = APPLICATION_DIRECTORY & "Styles.txt"

	' Open file for writing, create it if it isn't found (should already exist by now)
	Set f = fso.OpenTextFile(sFileName, ForWriting, True)

	' Dump all Styles to file
	For Each objStyle in Styles
	
		' Build string and write it to file
		f.WriteLine Styles_BuildString(objStyle)
	
	Next 
	
	' Close file
	f.Close


End Sub

' Build String:
'		Build a string that represents the Style.
'		Used to store styles back to file
Function Styles_BuildString(p_objStyle)
	Dim sStyleString

	If p_objStyle Is Nothing Then Exit Function
	
	' Initialize string
	sStyleString = ""

	With p_objStyle
		sStyleString = sStyleString & .ID & ","
		sStyleString = sStyleString & .Name & ","
		sStyleString = sStyleString & Replace(.Description, ",", "~") & ","
		sStyleString = sStyleString & .Color.Red & ","
		sStyleString = sStyleString & .Color.Green & ","
		sStyleString = sStyleString & .Color.Blue
	End With
	
	' Return Style string
	Styles_BuildString = sStyleString

End Function

' Create From String:
'		Creates a new style from a string
Sub Styles_CreateFromString(p_sStyleString)
	Dim objStyle
	Dim arrValues
	
	' Ignore leading comments, treat these as comments.
	If Left (p_sStyleString, 1) = "'" Then Exit Sub

	' Style String is a line read from the Styles.txt file and should look like this
	' 1,FSN,FOX Sports Net standard style,128,128,128

	' Break string down into array
	arrValues = Split(p_sStyleString, ",")
	
	' Check that enough values are present
	If UBound(arrValues) < 5 Then
		MsgBox "String '" & p_sStyleString & "' doesn't contain enough fields."
		Exit Sub
	End If

	' Retrieve style with AutoCreate
	Set objStyle = Styles.Style(arrValues(0), True)
	With objStyle
		.Name = arrValues(1)
		.Description = arrValues(2)
		.Color.Red = arrValues(3)
		.Color.Green = arrValues(4)
		.Color.Blue = arrValues(5)
		
		' Optional
		If (UBound (arrValues) > 5) Then
			.LoadFlags (arrValues(6))
		End If
		
	End With

End Sub

' Create Default:
'		Creates a default Styles file if none has been found
Sub Styles_CreateDefaults()

	Styles_CreateFromString "1,FSN,FOX Sports Net standard style,128,128,128"


End Sub


' Styles_TextColor
'		Some white text does not read properly on some styles.
Function Styles_TextColor

	Select Case UCASE( Settings.Style )

		Case "FSN_NATIONAL_COLLEGE","FSN_ARZ_COLLEGE","FSN_BA_COLLEGE","FSN_CHI_COLLEGE","FSN_DET_COLLEGE","FSN_FLA_COLLEGE","FSN_MW_COLLEGE","FSN_NORTH_COLLEGE","FSN_NORTHWEST_COLLEGE","FSN_NE_COLLEGE","FSN_NY_COLLEGE","FSN_OHIO_COLLEGE","FSN_PIT_COLLEGE","FSN_RM_COLLEGE","FSN_SOUTH_COLLEGE","FSN_SOUTHWEST_COLLEGE","FSN_WEST_COLLEGE","FSN_WEST2_COLLEGE","SUNSHINE"
			Styles_TextColor = "[c " & ecalBlack & "]"

		Case Else
			Styles_TextColor = "[c " & ecalWhite & "]"
			
	End Select

End Function

' Styles_LookupFlag
'		Look up flag and returns a value, return -1 for not found.
Function Styles_LookupFlag (p_sFlag)

	Dim objStyle
	Set objStyle = Styles.GetStyleByName(Settings.Style)
	
	' Force upper to match .exe
	p_sFlag = UCase(p_sFlag)
	
	If objStyle Is Nothing Then 
	
		Styles_LookupFlag = ""
		
	Else
		If objStyle.Flags.Exists (p_sFlag) Then
			Styles_LookupFlag = objStyle.Flags(p_sFlag)
		Else
			Styles_LookupFlag = ""
		End If
		
	End If
	

End Function



