'=====================================================================
' Football on FSN 2003
' FOX Sports VBScript Custom Scripting Module
'---------------------------------------------------------------------
' Jason Colbert (JasonCol@Fox.com)
' August 2003
'---------------------------------------------------------------------
' Contains "Custom scripts" - procedures that have been added to
' extend the functionalities of the application. These procedures
' are normally linked to menu items in the application's Scripts menu
'=====================================================================

Option Explicit

'=====================================================================
'   DECLARATIONS
'---------------------------------------------------------------------

' None

'=====================================================================
'   GENERAL PROCEDURES
'---------------------------------------------------------------------

' Custom Script Add:
'       Sub that creates a new script object that holds
'       a link to a procedure in the scripting module
Sub CustomScriptAdd(p_sName, p_sDescription, p_sProcedureName)

    Dim objScript
    
    ' Add the Custom Script
    Set objScript = Scripts.Add
    objScript.Name = p_sName
    objScript.Description = p_sDescription
    objScript.ProcedureName = p_sProcedureName

End Sub


' Custom Scripts Initialize:
'       Create Custom Script objects and add them to the Scripts collection
' CustomScriptAdd "Menu Item Caption", "Description", "Procedure Name"
Sub CustomScriptsInitialize()

    ' Add Custom script for checking the headshots for the players
    CustomScriptAdd "Check Headshots", "Check Existence of Headshots", "CheckHeadshots" 

End Sub




'=====================================================================
'   CUSTOM PROCEDURES
'---------------------------------------------------------------------

'=====================================================================
'   CUSTOM PROCEDURES
'---------------------------------------------------------------------

' Check Headshots:
'		Checks for headshots for all players in current teams
Sub CheckHeadshots()

	Dim fso
	Dim sReport
	Dim sHeadshot
	Dim iCounter

	Dim iPlayerCounter
	Dim objPlayer

	Set fso = CreateObject("Scripting.FileSystemObject")
	
	' Initialize report header
	sReport = ""
	sReport = sReport & "Start from directory: " & APPLICATION_DIRECTORY & "\" & OPTION_LEAGUE & "\Elements\" & vbCrlf & vbCrLf

	' Home Team
	sReport = sReport & "Home Team: " & Game.Home.Players.Count & " players." & vbCrLf
	iCounter = 0
	
	For iPlayerCounter = 1 to Game.Home.Players.Count

		' Retrieve player
		Set objPlayer = Game.Home.Players.Item(iPlayerCounter)

		' Player headshot
		sHeadshot = Interface.GetHeadshotPath (Game.Home.Players.Item(iPlayerCounter), False)
		
		' Check if file exists 
		If sHeadshot <> "" Then
			' Nothing, file exists		
		Else
			iCounter = iCounter + 1
			sReport = sReport & "  " & objPlayer.FirstName & " " & objPlayer.LastName & vbCrLf
		End If
	
	Next
	
	' Report on headshots missing
	Select Case iCounter
	
		Case 0 ' None
			sReport = sReport & "  (no headshots are missing)." & vbCrLf & vbCrLf
			
		Case 1 ' One
			sReport = sReport & "  1 headshot missing." & vbCrLf & vbCrLf

		Case Else ' More
			sReport = sReport & "  " & iCounter & " headshots missing." & vbCrLf & vbCrLf

	End Select

	' Visitors Team
	sReport = sReport & "Visitors Team: " & Game.Visitors.Players.Count & " players." & vbCrLf
	iCounter = 0
	
	' Retrieve player
	For iPlayerCounter = 1 to Game.Visitors.Players.Count

		' Retrieve player
		Set objPlayer = Game.Visitors.Players.Item(iPlayerCounter)

		' Player headshot
		sHeadshot = Interface.GetHeadshotPath (Game.Visitors.Players.Item(iPlayerCounter), False)
		
		' Check if file exists 
		If sHeadshot <> "" Then
			' Nothing, file exists		
		Else
			iCounter = iCounter + 1
			sReport = sReport & "  " & objPlayer.FirstName & " " & objPlayer.LastName & vbCrLf
		End If
	
	Next
	
	' Report on headshots missing
	Select Case iCounter
	
		Case 0 ' None
			sReport = sReport & "  (no headshots are missing)." & vbCrLf & vbCrLf
			
		Case 1 ' One
			sReport = sReport & "  1 headshot missing." & vbCrLf & vbCrLf

		Case Else ' More
			sReport = sReport & "  " & iCounter & " headshots missing." & vbCrLf & vbCrLf

	End Select
	
	' Report
	MsgBox sReport, vbOKOnly, "Headshot Check"

End Sub

