'=====================================================================
' FSN 2008
' FOX Sports VBScript module
'---------------------------------------------------------------------
' Etienne Laneville (Etienne.Laneville@FOX.com)
' July 2008
'=====================================================================

'=====================================================================
'   DECLARATIONS
'---------------------------------------------------------------------
Const YELLOWCARD_TEXT = "YELLOW CARD"
Const YELLOWCARD_TEXT_SPANISH = "AMONESTADO"
Const REDCARD_TEXT = "RED CARD"
Const REDCARD_TEXT_SPANISH = "EXPULSADO"
Const SUBSTITUTION_TEXT = "Substitution"
Const SUBSTITUTION_TEXT_SPANISH = "Sustitución"    ' SUSTITUCIÓN
const ET_TEXT = "ET"
Const ET_SPANISH_TEXT = "TE"
Const AGG_TEXT = "AGG"
Const AGG_TEXT_SPANISH = "GBI"
Const GOAL_TEXT = "GOAL"
Const GOAL_TEXT_SPANISH = "GOL"

Const STOPPAGETIME = 0
Const EXTRATIME = 1

Dim m_iFoxBoxPositionX
Dim m_iFoxBoxPositionY

Dim m_sFoxBoxState

Dim m_iGameNoteCounter
Dim m_sGameNoteTop
Dim m_sGameNoteBottom
Dim m_iTeamNoteCounter

Dim m_sSponsor
Dim m_iSponsorType
Dim m_fSponsorExists

Dim ThisIsTest

Dim SubstitutionTeamExist

Dim CurrentState
'Dim DropZoneType
Dim CurrentHomeScore
Dim CurrentVisitorsScore

'Dim TeamScored
Dim CardPlayerExist
Dim GoalPlayerExist

'Dim CurrentGameNoteHeader
Dim goalMinFormatted

Dim ogText

Dim FootnoteInsertionState

Const FOXBOX_FOXBOX = 2
Const FOXBOX_OUTOFTOWNSCORE = 22

'=====================================================================
'	FOX BOX
'---------------------------------------------------------------------

' FoxBox Go to State:
'       Controls FoxBox states
Sub FoxBox_GoToState(p_sStateName)
	Dim sInvokeMethod
	Dim iCounter
	Dim iLineCounter
	Dim sText
	Dim fCrawl
	Dim objTeam
	Dim objPlayer
	Dim objSplit
	Dim sCategory, sVisitors, sHome
	Dim titleSummary

	' Update Data Dictionary
	Game_Update

    	Log.LogEvent "FoxBox: go to " & p_sStateName, "Debug", 0, "Renderer"

	If p_sStateName = "OutOfTownScore" Then
		If FoxBox.State <> FOXBOX_FOXBOX Then
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=OFF")
			Delay .3
		End If
	Else
		If FoxBox.State = FOXBOX_OUTOFTOWNSCORE Then 
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=OFF")
			Delay .3
		End If
	End If

	Select Case p_sStateName
	
		Case "Retracted"	
			
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET BASE_DISPATCH=FALSE") 
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=OFF")
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET BASE_WIDTH=0") 
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET POPUP_NOTELINE_DISPATCH=OFF") 
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SIDEPOD_DISPATCH_AGG_LOAD=SIDEPOD_AGG_OFF") 

			CurrentState =  "Retracted"
                        
		Case "FoxBox"
			Set_SubstitutionChip
			
			FoxBox_GoToFoxBox

			CurrentState =  "FOXBOX"

		Case "GameNote"

			GameNote_Update
			
			CurrentState =  "GameNote"

		Case "Card"

			Insert_Card

			CurrentState =  "Card"	

		Case "TeamComparison"

			
			'TeamComparison_Update
			'Insert_TeamComparison

			Insert_MultilineTeamStats

			CurrentState =  "TeamComparison"
		
		Case "Substitution"
			Insert_Substitution

			CurrentState =  "Substitution"	


		Case "Goal"
			'Which team goaled is set in Game_Goal of GameAndClockUpdate.vbs
			'FoxBox_GoalInsert


			CurrentState =  "Goal"	
		
		Case "Shootout"
			
			If Game.Shootout Then
				TurnOffClock
				Initialize_ShootoutScore
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PENALTY_KICKS")
			End IF
			
			CurrentState =  "Shootout"

		Case "GoalMinutes"

			goalMinFormatted = ""

			Insert_GoalMinutes

			CurrentState =  "GoalMinutes"

		Case "TeamNote"

			Insert_TeamNote
		
			CurrentState =  "TeamNote"
		
		Case "PossessionChart"
			
			Update_PossessionChart
			Insert_PossessionChart

			CurrentState =  "PossessionChart"

		Case "GameFlow"

			GameFlow_Set

			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOCUS_LOAD=2")
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=GAMEFLOW")

			CurrentState =  "GameFlow"

		Case "JerseyColors"

			Update_JerseyColors
			Insert_JerseyColors
			
			CurrentState =  "JerseyColors"

		Case "Commentators"
			
			UpdateCommentators
			
			CurrentState =  "Commentators"

		Case "Locator"
			
			UpdateLocator
			Insert_Locator

			CurrentState =  "Locator"

		Case "BillboardSponsor"

			Insert_BillboardSponsor

			CurrentState =  "BillboardSponsor"

		Case "Weather"
			'Insert_Weather

		Case "Reporting"

			Insert_Reporting

			CurrentState = "Reporting"
	
		Case "Standings"
			
			Insert_Standings
			'Basics are laied out
			

			CurrentState = "Standings"

		Case "PlayerNote"

			Insert_PlayerNote

			CurrentState = "PlayerNote"

		Case "VoiceOf"
			
			UpdateVoiceOf

			CurrentState = "VoiceOf"

		'Case "VoiceOfHead"

			'UpdateVoiceOffWithHead
				
			'CurrentState = "VoiceOfHead"
		
		Case "TalentsHead"

			UpdateCommentatorsWithHead
			
			CurrentState = "TalentsHead"

		Case "Summary"
		 	
			If Plugin.SummarySelected = 0 Then	
				titleSummary = "SCORING SUMMARY"
				Call UpdateSummary(Plugin.ScoringSummary, titleSummary)
				'UpdateScoreSummary
			Else
				titleSummary = "MATCH SUMMARY"
				Call UpdateSummary(Plugin.MatchSummary, titleSummary)
				'UpdateMatchSummary
			End If

			CurrentState = "Summary"

		Case "OutOfTownScore"
			OTS_Update
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GLOW=ON")
			Insert_OTS

			CurrentState = "OutOfTownScore"

		Case Else

			CurrentState = "Else"			
		    Log.LogEvent "FoxBox: State Not Found", "Debug", 0, "Renderer"

	End Select

	If Plugin.FootterChecked Then
		If p_sStateName = "Retracted" or p_sStateName = "FoxBox"  or p_sStateName = "Shootout"Then
			'do not insert footer
			Retract_Footnote
		Else
			Footer_Update	
			'If FoxBox.State = 2 Then
			If Not FootnoteInsertionState Then
				Plugin.Start_FooterNoteTimer
			End If
		End If
	End If
	
	'keeping tack of preveiw state
	With Plugin
	 	If .SendToRenderer Then
       	 		.outputState = CurrentState
   	 	End If
    
    		If .SendToPreview Then
			'Goal is a special calse becaseu it animates only n output
			If p_sStateName <> "Goal" Then
	        		.previewState = CurrentState
			End If
    		End If
	End With

End Sub

Sub FoxBox_GoToFoxBox()
	'this is used for automated sponsor
	Plugin.PreviousState = CurrentState

	Select Case CurrentState
		Case "GameNote", "GoalMinutes", "TeamNote", "Card", "Substitution", "TeamComparison", "Shootout", "PossessionChart", "Commentators", "Locator", "JerseyColors", "BillboardSponsor", "Reporting", "VoiceOf", "Standings", "PlayerNote", "Summary", "TalentsHead", "OutOfTownScore"

			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=OFF")

			If CurrentState = "Summary" Or CurrentState = "OutOfTownScore" Then  
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GLOW=OFF")
			End If

		Case "Goal"
			'Goal_Retract	

		Case "GameFlow"
			'Do Notthing

		Case Else

			UpdateHeader
			'Insert FoxBox
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET BASE_DISPATCH=TRUE") 
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET BASE_WIDTH=1") 
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=OFF") 

			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SIDEPOD_REQUEST=OFF") 
	End Select
End Sub



Sub Reset_ToFoxBox(foxboxState)

	Select Case foxboxState
		Case "GameNote", "GoalMinutes", "TeamNote", "Card", "Substitution", "TeamComparison", "Shootout", "PossessionChart", "Commentators", "Locator", "JerseyColors", "BillboardSponsor", "Reporting", "VoiceOf", "Standings", "PlayerNote", "Summary", "TalentsHead", "OutOfTownScore"
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=OFF")

			If CurrentState = "Summary" Or CurrentState = "OutOfTownScore" Then  
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GLOW=OFF")
			End If
		Case "Goal"
			'Goal_Retract	
		Case "GameFlow"
			'Do Notthing

		Case Else
			'Insert FoxBox
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET BASE_DISPATCH=TRUE") 
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET BASE_WIDTH=1") 
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=OFF") 
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SIDEPOD_REQUEST=OFF") 
	End Select
End Sub



Sub FoxBox_GoalInsert()

	If FoxBox.State = 1 Then 'If FoxBox is not inserted do not play goal anim
		Exit Sub
	End If
	
	Set_GoalText

	If Plugin.GoalType = 1 Then
		AudiGoal_Insert
	Else
		RegularGoal_Insert
	End If
End Sub

Sub Test()

	'Call Set_SponsorElement(True)

	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SponsorBlock=False")
	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SPONSOR=ON") 

	'Delay .5

	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SponsorBlock=True")	
	'
	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET BASE_DISPATCH=FALSE") 
	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=OFF") 
	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET BASE_WIDTH=0") 
	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET POPUP_NOTELINE_DISPATCH=OFF") 
	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SIDEPOD_DISPATCH_AGG_LOAD=SIDEPOD_AGG_OFF") 	
	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SPONSOR=OFF") 
	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SPECIALTY_EVENT=GOAL_ON") 

	'Delay 6

'	'Goal_Retract
	'RegularGoal_Retract
	'
End Sub

'common commads used in both Regular and Audi Goal
Sub Invoke_GoalCommonCommands()
	Viz_Send_OutputDirect (VizLayer & "*FUNCTION*DataPool*Data SET BASE_DISPATCH=FALSE") 
	Viz_Send_OutputDirect (VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=OFF") 
	Viz_Send_OutputDirect (VizLayer & "*FUNCTION*DataPool*Data SET BASE_WIDTH=0") 
	Viz_Send_OutputDirect (VizLayer & "*FUNCTION*DataPool*Data SET POPUP_NOTELINE_DISPATCH=OFF") 
	Viz_Send_OutputDirect (VizLayer & "*FUNCTION*DataPool*Data SET SIDEPOD_DISPATCH_AGG_LOAD=SIDEPOD_AGG_OFF") 
	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SPONSOR=OFF") 
	Viz_Update_OutputDirect "SPONSOR", "OFF"
	
	Viz_Send_OutputDirect (VizLayer & "*FUNCTION*DataPool*Data SET SponsorBlock=True")
	
End Sub

Sub RegularGoal_Insert()
	Invoke_GoalCommonCommands
	Viz_Send_OutputDirect (VizLayer & "*FUNCTION*DataPool*Data SET SPECIALTY_EVENT=GOAL_ON") 
	Viz_Send_OutputDirect (VizLayer & "*FUNCTION*DataPool*Data SET SponsorBlock=True")
End Sub

Sub RegularGoal_Retract()
	Viz_Send_OutputDirect(VizLayer & "*FUNCTION*DataPool*Data SET SPECIALTY_EVENT=OFF") 
End Sub

Sub AudiGoal_Insert()
	Invoke_GoalCommonCommands
	Viz_Send_OutputDirect(VizLayer & "*FUNCTION*DataPool*Data SET SPECIALTY_EVENT=AUDI_GOAL_ON")
	Viz_Send_OutputDirect(VizLayer & "*FUNCTION*DataPool*Data SET SponsorBlock=True")	
End Sub

Sub AudiGoal_Retract()
	Viz_Send_OutputDirect(VizLayer & "*FUNCTION*DataPool*Data SET SPECIALTY_EVENT=AUDI_GOAL_OFF")
End Sub

Sub Goal_Retract()

	If Plugin.GoalType = 1 Then
		AudiGoal_Retract
	Else
		RegularGoal_Retract
	End If

End Sub


Sub Set_GoalText() 'in progress

	Dim GoalTextLangType
	Dim sText

	Select Case PluginSettings.GoalTextType 
		Case 0
			GoalTextLangType = GOAL_TEXT 
		Case 1
			GoalTextLangType = GOAL_TEXT_SPANISH
	End Select

	Viz_Send_OutputDirect(VizLayer & "*FUNCTION*DataPool*Data SET GOAL TEXT=" & GoalTextLangType) 

End Sub


'=====================================================================
'	Goal Minutes	
'---------------------------------------------------------------------
Sub Insert_GoalMinutes()

	Update_GoalMinutes

	If GoalPlayerExist Then
		'if formatted goal mins exist
		If Len(goalMinFormatted) > 0 then

			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERFLAG_LOAD=")
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERHEAD_LOAD=")

			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PLAYER_GOAL_MINUTES")
		End If
	End If
End Sub

Sub Set_GoalMinutes(goalMinute)
	Plugin.GoalMinutes = goalMinute
End Sub

Function FormatGoalMinutes(objEvent)
	Dim goalMin

	If objEvent.StoppageTime Then
		FormatGoalMinutes = objEvent.Minute & chr(39) & "+" &  objEvent.StoppageMinute
	Else
		FormatGoalMinutes = objEvent.Minute & chr(39) 
	End If

End Function

Sub Update_GoalMinutes()
	Dim objTeam 
	Dim objEvent
	Dim teamName
	Dim messageBody
	Dim objGoaler
	Dim goalMin

	GoalPlayerExist = False

	Set objEvent = Game.GameEvents.SelectedEvent

	If NOT( (objEvent.EventType = 0) Or (objEvent.EventType = 4) Or (objEvent.EventType = 5) ) then
		Exit Sub
	End If

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/CARD_SELECT_LOAD=0")

	Select Case objEvent.TeamCode

		
		Case 0
			Set objTeam = Game.Home
			
			TeamName = UCase(Game.Home.ShortAbbreviation)

			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/AwayLogo=" & UCase(Game.Home.Abbreviation))
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOCUS_LOAD=0")
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/TRICODE_TEXTCOLOR_LOAD=" & PluginSettings.HomeFontColor)
		Case 1
			Set objTeam = Game.Visitors
			
			TeamName = UCase(Game.Visitors.ShortAbbreviation)

			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/HomeLogo=" & UCase(Game.Visitors.Abbreviation))
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOCUS_LOAD=1")
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/TRICODE_TEXTCOLOR_LOAD=" & PluginSettings.VisitorsFontColor)
	End Select
	
	
	'If objEvent.EventType = 4 Then
	'	'Own Goal
	'	ogText = "  (OG)"
	'Else
	'	ogText = ""
	'End If

	Select Case objEvent.EventType
		Case 4
			ogText = "  (OG)"
		Case 5
			ogText = "  (P)"
		Case Else
			ogText = ""

	End Select
	
	'Can add "P" for Penalty Kick goal
	'ogText = GoalTextByEventType(objEvent.EventType)

	goalMinFormatted = FormatGoalMinutes(objEvent) & ogText

	Set objGoaler = Game.GameEvents.SelectedEvent.Player1

	If Not objGoaler Is Nothing Then

		GoalPlayerExist = True

		With objGoaler
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/PLAYER_NAME_LOAD=" & .DisplayName)
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/JERSEY1_LOAD=" & "#" & .Jersey)
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/TIME_LOAD=" & goalMinFormatted)
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/TRICODE_DISPLAY_LOAD=" & TeamName)

			Call Set_PlayerTemplate(.DisplayName, .DisplayJersey, goalMinFormatted, TeamName)
		End With
	End If
	
	'making sure notes below are empty
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LOAD=")
End Sub

Sub Set_PlayerTemplate(playerName, playerJersey, playerMintuesText, playerTeamName)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/PLAYER_NAME_LOAD=" & HEADER_DOUBLEQUOTES & playerName &  HEADER_DOUBLEQUOTES) 
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/JERSEY1_LOAD=" & HEADER_DOUBLEQUOTES & playerJersey & HEADER_DOUBLEQUOTES)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/TIME_LOAD=" & HEADER_DOUBLEQUOTES & playerMintuesText & HEADER_DOUBLEQUOTES)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/TRICODE_DISPLAY_LOAD=" & HEADER_DOUBLEQUOTES & playerTeamName & HEADER_DOUBLEQUOTES)
End Sub

Sub Set_PlayerTemplate_LIVE(playerName, playerJersey, playerMintuesText, playerTeamName)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/PLAYER_NAME_LIVE=" & HEADER_DOUBLEQUOTES & playerName &  HEADER_DOUBLEQUOTES) 
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/JERSEY1_LIVE=" & HEADER_DOUBLEQUOTES & playerJersey & HEADER_DOUBLEQUOTES)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/TIME_LIVE=" & HEADER_DOUBLEQUOTES & playerMintuesText & HEADER_DOUBLEQUOTES)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/TRICODE_DISPLAY_LIVE=" & HEADER_DOUBLEQUOTES & playerTeamName & HEADER_DOUBLEQUOTES)
End Sub



'Function FormatgoalMinutes_ORG(goalMin)
'	Dim tmpMins
'
'	If (goalMin >= 45) And (goalMin < 90) Then
'
'		If (goalMin - 45) = 0 Then
'			FormatgoalMinutes =  "45+" & Chr(39) 
'		Else
'			
'			tmpMins =  "45'+" & Cstr(goalMin - 45) '& Chr(39) 
'
'			If (goalMin - 45) > 9 Then '9 is arbitary
'				
'				result = MsgBox (tmpMins & "?", vbYesNo, "Yes No Example")	
'
'				Select Case result
'					Case vbYes
'						FormatgoalMinutes = tmpMins 
'					Case vbNo
'						FormatgoalMinutes = "" 
'				End Select
'			Else
'				FormatgoalMinutes = tmpMins 
'			End If
'
'		End If
'
'	End If 
'	
'	If (goalMin >= 90) And (goalMin < 105) Then
'
'		If (goalMin - 90) = 0 Then
'			FormatgoalMinutes =  "90+" & Chr(39) 
'		Else
'			tmpMins =  "90'+" & Cstr(goalMin - 90) '& Chr(39) 
'
'			If (goalMin - 90) > 9 Then '9 is arbitary
'				
'				result = MsgBox (tmpMins & "?", vbYesNo, "Yes No Example")	
'
'				Select Case result
'					Case vbYes
'						FormatgoalMinutes = tmpMins 
'					Case vbNo
'						FormatgoalMinutes = "" 
'				End Select
'			Else
'				FormatgoalMinutes = tmpMins 
'			End If
'			
'		End If
'	End If
'	
'	If (goalMin >= 105) And (goalMin < 120) Then
'
'		If (goalMin - 105) = 0 Then
'			FormatgoalMinutes =  "105+" & Chr(39) 
'		Else
'			tmpMins =  "105'+" & Cstr(goalMin - 105) '& Chr(39) 
'
'			If (goalMin - 105) > 9 Then '9 is arbitary
'				
'				result = MsgBox (tmpMins & "?", vbYesNo, "Yes No Example")	
'
'				Select Case result
'					Case vbYes
'						FormatgoalMinutes = tmpMins 
'					Case vbNo
'						FormatgoalMinutes = "" 
'				End Select
'			Else
'				FormatgoalMinutes = tmpMins 
'			End If
'
'
'		End If
'	End If
'	
'	If (goalMin >= 120) Then
'		If (goalMin - 120) = 0 Then
'			FormatgoalMinutes =  "120+" & Chr(39) 
'		Else
'			
'			tmpMins =  "120'+" & Cstr(goalMin - 120) '& Chr(39) 
'
'			If (goalMin - 120) > 9 Then '9 is arbitary
'				
'				result = MsgBox (tmpMins & "?", vbYesNo, "Yes No Example")	
'
'				Select Case result
'					Case vbYes
'						FormatgoalMinutes = tmpMins 
'					Case vbNo
'						FormatgoalMinutes = "" 
'				End Select
'			Else
'				FormatgoalMinutes = tmpMins 
'			End If
'		end If
'	End If
'End Function

'=====================================================================
'	Yellow/Red Card	
'---------------------------------------------------------------------

Sub Update_Card()
	Dim objTeam 
	Dim objEvent
	Dim teamName
	Dim objPlayer
	Dim cardType
	Dim pName
	Dim numOfYellowCard
	Dim objEventTmp


	CardPlayerExist = False
	numOfYellowCard = 0 
	
	Set objEvent = Game.GameEvents.SelectedEvent

	If objEvent is Nothing Then
		Exit Sub
	End If

	Select Case Game.GameEvents.SelectedEvent.TeamCode
		Case 0
			Set objTeam = Game.Home
			TeamName = UCase(Game.Home.Shortabbreviation)
		Case 1
			Set objTeam = Game.Visitors
			TeamName = UCase(Game.Visitors.Shortabbreviation)
	End Select

	Set objPlayer = Game.GameEvents.SelectedEvent.Player1

	If Not objPlayer Is Nothing Then
		If objEvent.EventType = 1 Then 'Yellow
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/CARD_SELECT_LOAD=1")

			'Checking if this is 2nd yellow
			If Game.GameEvents.Count > 1 Then
				For i = 1 to Game.GameEvents.Count	
					Set objEventTmp =  Game.GameEvents(i)
					If objEventTmp.EventType = 1 Then 'yellow
						'checking if team is the same as a selected player
						If objPlayer.Team.Abbreviation = objEventTmp.Team.Abbreviation Then
							If objPlayer.Jersey = objEventTmp.Player1.Jersey Then
								numOfYellowCard = numOfYellowCard + 1
							End If
						End If
					End If
				Next
			End If

			'2nd Yellow
			If numOfYellowCard = 2 Then
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/CARD_SELECT_LOAD=3")	
			Else
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/CARD_SELECT_LOAD=1")
			End If
			
			CardPlayerExist = True
		End If

		If objEvent.EventType = 2 Then  'Red
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/CARD_SELECT_LOAD=2")
			CardPlayerExist = True
		End If

		'command for 2 yellow
		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/CARD_SELECT_LOAD=3")

		With objPlayer	

		'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/PLAYER_NAME_LOAD=" & .DisplayName)
		'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/JERSEY1_LOAD=" & "#" & .Jersey)
		'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/TIME_LOAD=")
		'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/TRICODE_DISPLAY_LOAD=" & TeamName)

			Call Set_PlayerTemplate(.DisplayName, .DisplayJersey, "", TeamName)

		End With
	End If
	
	'making sure notes below is empty
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LOAD=")

End Sub

Sub Insert_Card()
	
	Update_Card

	If CardPlayerExist Then
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PLAYER_GOAL_MINUTES")
	End If

End Sub

Sub Insert_GameEvent()
	Dim objEvent
	Set objEvent = Game.GameEvents.SelectedEvent

	Select Case  objEvent.EventType

		Case 0, 4
			Insert_GoalMinutes
		Case 1, 2
			Insert_Card
		Case 3
			Insert_Substitution
	End Select

End Sub


'=====================================================================
'	TEAM NOTE
'---------------------------------------------------------------------
Sub Insert_TeamNote()

	Update_TeamNote

	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=NOTELINE_TEAM")
End Sub

'Sub Update_TeamNote_Original()
'	Dim teamName
'	Dim teamNoteText
'
'	If PluginSettings.TeamNoteSelectedTeam = 1 Then
'		'home selected
'		teamName = Game.Home.ShortAbbreviation	
'		teamNoteText = Game.Home.Notes.Selected.Body
'		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOCUS_LOAD=0")
'		
'	Else
'		'away selected
'		teamName = Game.Visitors.ShortAbbreviation
'		teamNoteText = Game.Visitors.Notes.Selected.Body
'		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOCUS_LOAD=1")
'	End If
'	
'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET NOTELINE_TEAM/TRICODE_DISPLAY_LOAD=" & teamName)
'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET NOTELINE_TEAM/TEXT1_LOAD=" & HEADER_DOUBLEQUOTES & teamNoteText & HEADER_DOUBLEQUOTES)
'End Sub

Sub Update_TeamNote()
	Dim teamName
	Dim teamNoteText
	Dim tmp
	Dim objTeamNote

	If Plugin.TeamNoteSelectedTeam = 1 Then
		'home selected
		teamName = Game.Home.ShortAbbreviation	
		Set objTeamNote = Game.Home.Notes.GetNoteByTitle(Plugin.TeamNoteSelected)
		'teamNoteText = Game.Home.Notes.GetNoteByTitle(Plugin.TeamNoteSelected)
	Else
		'away selected
		teamName = Game.Visitors.ShortAbbreviation
		Set objTeamNote = Game.Visitors.Notes.GetNoteByTitle(Plugin.TeamNoteSelected)
		'teamNoteText = Game.Visitors.Notes.GetNoteByTitle(Plugin.TeamNoteSelected)
	End If

	If Not objTeamNote Is Nothing Then
		Log.LogEvent "Team Note " & Plugin.TeamNoteSelected & " FOUND ",  "Debug", 0, "Renderer"
		teamNoteText = objTeamNote.Body
	Else
		Log.LogEvent "Team Note " & Plugin.TeamNoteSelected & " NOT FOUND ", "Debug", 0, "Renderer"
		Exit Sub
	End If

	tmp = Split(teamNoteText, "\")

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/JERSEY1_LOAD=")
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/TIME_LOAD=")
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/TRICODE_DISPLAY_LOAD=" & HEADER_DOUBLEQUOTES & teamName & HEADER_DOUBLEQUOTES)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERFLAG_LOAD=")
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERHEAD_LOAD=")

	Select Case Ubound(tmp)
		Case 0
			'Single line note
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/PLAYER_NAME_LOAD=" & HEADER_DOUBLEQUOTES & tmp(0) & HEADER_DOUBLEQUOTES) 
	'		making sure notes below are empty
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LOAD=")
			
			'dropdown type
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PLAYER_GOAL_MINUTES")
		Case 1
			'2 line note
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/PLAYER_NAME_LOAD=" & HEADER_DOUBLEQUOTES & tmp(0) & HEADER_DOUBLEQUOTES) 
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LOAD=" & HEADER_DOUBLEQUOTES & tmp(1) & HEADER_DOUBLEQUOTES)

			'dropdown type
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PLAYER_GOAL_MINUTES_NOTE")
		Case 2
			'3 line note
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/PLAYER_NAME_LOAD=" & HEADER_DOUBLEQUOTES & tmp(0) & HEADER_DOUBLEQUOTES) 
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LOAD=" & HEADER_DOUBLEQUOTES & tmp(1) & HEADER_DOUBLEQUOTES)
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE2_LOAD=" & HEADER_DOUBLEQUOTES & tmp(2) & HEADER_DOUBLEQUOTES)

			'dropdown type
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PLAYER_GOAL_MINUTES_NOTE2")
		Case else
			' do nothing
	End Select
End Sub

Sub Update_TeamNote_Original()
	Dim teamName
	Dim teamNoteText
	Dim tmp

	If Plugin.TeamNoteSelectedTeam = 1 Then
		'home selected
		teamName = Game.Home.ShortAbbreviation	
		teamNoteText = Game.Home.Notes.Selected.Body
	Else
		'away selected
		teamName = Game.Visitors.ShortAbbreviation
		teamNoteText = Game.Visitors.Notes.Selected.Body
	End If

	tmp = Split(teamNoteText, "\")

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/JERSEY1_LOAD=")
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/TIME_LOAD=")
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/TRICODE_DISPLAY_LOAD=" & HEADER_DOUBLEQUOTES & teamName & HEADER_DOUBLEQUOTES)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERFLAG_LOAD=")
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERHEAD_LOAD=")

	Select Case Ubound(tmp)
		Case 0
			'Single line note
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/PLAYER_NAME_LOAD=" & HEADER_DOUBLEQUOTES & tmp(0) & HEADER_DOUBLEQUOTES) 
	'		making sure notes below are empty
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LOAD=")
			
			'dropdown type
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PLAYER_GOAL_MINUTES")
		Case 1
			'2 line note
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/PLAYER_NAME_LOAD=" & HEADER_DOUBLEQUOTES & tmp(0) & HEADER_DOUBLEQUOTES) 
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LOAD=" & HEADER_DOUBLEQUOTES & tmp(1) & HEADER_DOUBLEQUOTES)

			'dropdown type
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PLAYER_GOAL_MINUTES_NOTE")
		Case 2
			'3 line note
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/PLAYER_NAME_LOAD=" & HEADER_DOUBLEQUOTES & tmp(0) & HEADER_DOUBLEQUOTES) 
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LOAD=" & HEADER_DOUBLEQUOTES & tmp(1) & HEADER_DOUBLEQUOTES)
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE2_LOAD=" & HEADER_DOUBLEQUOTES & tmp(2) & HEADER_DOUBLEQUOTES)

			'dropdown type
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PLAYER_GOAL_MINUTES_NOTE2")
		Case else
			' do nothing
	End Select
End Sub

'This is called from Plugin
Sub TeamNote_Change()
	'Plugin.ReverseTeamOrder = Settings.ReverseTeamOrder
	Plugin.TeamNotesRefresh
End Sub

'=====================================================================
'	BILLBOARD SPONSOR	
'---------------------------------------------------------------------
Sub Insert_BillboardSponsor()

	Update_BillboardSponsor

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=BILLBOARD")
End Sub

Sub Update_BillboardSponsor()

	Dim bbSponsorPath

	bbSponsorPath = PluginSettings.SponsorDirectory & Plugin.BillboardSponsorSelected 

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOCUS_LOAD=2")
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET BILLBOARD_IMAGE=" & bbSponsorPath )
End Sub

'=====================================================================
'	STANDINGS
'---------------------------------------------------------------------
Sub Insert_Standings()
	Plugin.RefreshStandingTitle
	'Update_Standings
	Update_Standings_ByOrder
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=STANDINGS")
End Sub

'Sub Update_Standings()
'
'	Dim objStanding
'	Dim standingTitle
'	Dim standingTitleHeader
'	Dim GDSighn
'	Dim goalsFor
'
'	GDSighn = ""
'
'	groupTitle = HEADER_DOUBLEQUOTES & Plugin.StandingGroupTitle & HEADER_DOUBLEQUOTES
'
'	Set objStanding = Soccer.Standings.GetStandingsByName(Plugin.StandingGroupSelected)
'
'	Dim objTeam 
'	Dim i
'	Dim title, gamesPlayed, teamPoints, goalDifferential
'
'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/HEADER_LOAD=" & groupTitle)
'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/CAT2_LOAD=GD")
'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/CAT1_LOAD=PTS")
'
'	If Plugin.StandingsThressStats Then
'		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/CAT3_LOAD=GF")
'	Else
'		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/CAT3_LOAD=")
'	End If
'
'	With Soccer.Standings.Teams
'		For i = 1 to 4
'			'title = .Team(objStanding.Team(i)).DisplayAbbreviation
'			title = .Team(objStanding.Team(i)).Name
'			'gamesPlayed = .Team(objStanding.Team(i)).GamesPlayed 
'			teamPoints = .Team(objStanding.Team(i)).Points
'			goalDifferential = .Team(objStanding.Team(i)).GoalDifferential
'			
'			If Plugin.StandingsThressStats Then
'				goalsFor = .Team(objStanding.Team(i)).Goals
'			Else
'				goalsFor = ""
'			End If
''
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/R" & i & "T_LOAD=" & title)
'			
'
'			If goalDifferential > 0 Then
'				GDSighn = "+"
'			Else
'				GDSighn = ""
'			End If
'
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/R" & i & "S2_LOAD=" & GDSighn & goalDifferential )
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/R" & i & "S1_LOAD=" & teamPoints )
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/R" & i & "S3_LOAD=" & goalsFor )
'		Next
'	End With
'
'End Sub

Sub Update_StandingStatsCount(ThreeStatsOn)
	Plugin.StandingsThressStats = ThreeStatsOn
End Sub

Sub Clear_StandingHighlight()
	Dim i, j
	For i = 1 to 4
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/R" & (i) & "TY_LOAD=0")
		For j = 1 to 3
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/R" & (i) & "S" & (j) & "Y_LOAD=0")
		Next
	Next	
End Sub

Sub Update_Standings_ByOrder()
	Dim objStanding
	Dim standingTitle
	Dim standingTitleHeader
	Dim GDSighn
	Dim goalsFor
	Dim standingCatText
	Dim standingCatStat
	Dim objTeam 
	Dim i, j
	Dim titleTeamName, gamesPlayed, teamPoints, goalDifferential

	Clear_StandingHighlight
	
	GDSighn = ""

	groupTitle = HEADER_DOUBLEQUOTES & Plugin.StandingGroupTitle & HEADER_DOUBLEQUOTES

	Set objStanding = Soccer.Standings.GetStandingsByName(Plugin.StandingGroupSelected)
	
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/HEADER_LOAD=" & groupTitle)

	'Setting Col Headers
	With PluginSettings
		For i = 1 to 3
			Select Case .StandingStatsSelected(i)
				Case 1 'PTS
					standingCatText = "PTS"
				Case 2 'GD
					standingCatText = "GD"
				Case 3 'GF
					standingCatText = "GF"
				Case 4 'GP
					standingCatText = "GP"
			End Select

			If i = 3 Then 'clearing 3rd col
				If Not Plugin.StandingsThressStats Then 'clearing 3rd stats
					standingCatText=""
				End If
			End If
		
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/CAT" & (i) & "_LOAD=" & standingCatText)
		Next
	End With

	With Soccer.Standings.Teams
		For i = 1 to 4
		
			'Hilighting Team Name
			If InStr(objStanding.Highlights(i), "TEAM") > 0 Then
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/R" & (i) & "TY_LOAD=1")
			End If
				

			titleTeamName = .Team(objStanding.Team(i)).Name
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/R" & i & "T_LOAD=" & titleTeamName)

			teamPoints = .Team(objStanding.Team(i)).Points
			
			goalDifferential = .Team(objStanding.Team(i)).GoalDifferential(Plugin.CalculatedGD)

			goalsFor = .Team(objStanding.Team(i)).Goals

			gamesPlayed = .Team(objStanding.Team(i)).GamesPlayed

			With PluginSettings
				For j = 1 to 3
					Select Case .StandingStatsSelected(j)
						Case 1 'PTS
							standingCatStat = teamPoints

							If InStr(objStanding.Highlights(i), "PTS") > 0 Then
								Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/R" & (i) & "S" & (j) & "Y_LOAD=1")
							End If
						
						Case 2 'GD
							If goalDifferential > 0 Then
								GDSighn = "+"
							Else
								GDSighn = ""
							End If

							standingCatStat = GDSighn & goalDifferential

							If InStr(objStanding.Highlights(i), "GD") > 0 Then
								Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/R" & (i) & "S" & (j) & "Y_LOAD=1")
							End If
						
						Case 3 'GF
							standingCatStat = goalsFor

							If InStr(objStanding.Highlights(i), "GF") > 0 Then
								Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/R" & (i) & "S" & (j) & "Y_LOAD=1")
							End If
						Case 4 'GP
							standingCatStat = gamesPlayed

							If InStr(objStanding.Highlights(i), "GP") > 0 Then
								Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/R" & (i) & "S" & (j) & "Y_LOAD=1")
							End If
					End Select

					If j = 3 Then
						If Not Plugin.StandingsThressStats Then 'clearing 3rd stats
							standingCatStat=""
						End If
					End if

					Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STANDINGS/R" & (i) & "S" & (j) & "_LOAD=" & standingCatStat)

				Next

			End With
		Next
	End With
End Sub

'=====================================================================
'	PLAYER NOTE
'---------------------------------------------------------------------
Sub Insert_PlayerNote()
	Update_PlayerNote
End Sub

Sub PlayerNote_Change()

	If Not Interface.Player Is Nothing Then
		Plugin.PlayerNotesRefresh(Interface.Player.Notes)
	End If  

End Sub

'Sub Update_PlayerNote_Original()
'	Dim playerNote
'	Dim TeamName
'	Dim teamAbbreviation
'	Dim tmp
'	Dim headshotFileName
'
'	With Interface.Player
'		playerNote = .Notes.Selected.Body
'		TeamName = .Team.DisplayAbbreviation
'		teamAbbreviation = .Team.Abbreviation
'
'		msgbox playerNote
'		
'		'TRI_LASTNAME_FIRSTNAME_960
'		headshotFileName = teamAbbreviation & "_" & .LastName & "_" & .FirstName & "_960" 
'
'		tmp = Split(playerNote, "\")
'
'		'Player Info
'		Call Set_PlayerTemplate(.DisplayName, .Jersey, "", TeamName)
'		
'		If Plugin.PlayerNoteWithHeadshot Then
'			'logo 7 head
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERFLAG_LOAD=" & teamAbbreviation)
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERHEAD_LOAD=" & HEADSHOT_DIRECTORY_GH & teamAbbreviation & "/" & headshotFileName)
'		Else
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERFLAG_LOAD=")
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERHEAD_LOAD=")
'		End If
'	
'		If Ubound(tmp) > 0 Then
'			'2 line notes
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LOAD=" & HEADER_DOUBLEQUOTES & tmp(0) & HEADER_DOUBLEQUOTES)
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE2_LOAD=" & HEADER_DOUBLEQUOTES & tmp(1) & HEADER_DOUBLEQUOTES)
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PLAYER_GOAL_MINUTES_NOTE2")
'		Else
'			'l line note
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LOAD=" & HEADER_DOUBLEQUOTES & tmp(0) & HEADER_DOUBLEQUOTES)
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE2_LOAD=")
'
'			If Plugin.PlayerNoteWithHeadshot Then
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PLAYER_GOAL_MINUTES_NOTE2")
'			Else
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PLAYER_GOAL_MINUTES_NOTE")
'			End If
'		End If
'
'	End With 
'End Sub


Sub Update_PlayerNote()
	Dim objplayerNote
	Dim TeamName
	Dim teamAbbreviation
	Dim tmp
	Dim headshotFileName
	Dim selectedTeam
	Dim headshotExists
	
	headshotExists = false

	If Interface.Player Is Nothing Then
		Exit Sub
	End If

	With Interface.Player
		TeamName = .Team.DisplayAbbreviation
		teamAbbreviation = .Team.Abbreviation

		If  .Team.Abbreviation = Game.Home.Abbreviation Then
			selectedTeam = 1
		Else
			selectedTeam = 0
		End If

		'TRI_LASTNAME_FIRSTNAME_960
		'headshotFileName = teamAbbreviation & "-WOMEN" & "_" & .LastName & "_" & .FirstName & "_960" 
		headshotFileName = teamAbbreviation & "_" & .LastName & "_" & .FirstName & "_960" 

		Set objplayerNote = .Notes.GetNoteByTitle(Plugin.PlayerNoteSelected)
		
		If objplayerNote Is Nothing  Then
			Exit Sub	
		End If

		tmp = Split(objplayerNote.Body, "\")
		
		'Player Info
		'Call Set_PlayerTemplate(.DisplayName, .Jersey, "", TeamName)
		Call Set_PlayerTemplate(.DisplayName, .DisplayJersey, "", TeamName)

		headshotExists = Plugin.CheckHeadShotsExists(headshotFileName, selectedTeam)
		
		If Plugin.PlayerNoteWithHeadshot Then
			'setting headshots
			If headshotExists Then
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERFLAG_LOAD=" & teamAbbreviation)
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERHEAD_LOAD=" & HEADSHOT_DIRECTORY_GH & teamAbbreviation & "/" & headshotFileName)
			Else
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERFLAG_LOAD=")
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERHEAD_LOAD=")
			End If
		Else
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERFLAG_LOAD=")
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERHEAD_LOAD=")
		End If

		If Ubound(tmp) > 0 Then
			'2 line notes
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LOAD=" & HEADER_DOUBLEQUOTES & tmp(0) & HEADER_DOUBLEQUOTES)
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE2_LOAD=" & HEADER_DOUBLEQUOTES & tmp(1) & HEADER_DOUBLEQUOTES)
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PLAYER_GOAL_MINUTES_NOTE2")
		Else
			'l line note
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LOAD=" & HEADER_DOUBLEQUOTES & tmp(0) & HEADER_DOUBLEQUOTES)
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE2_LOAD=")

			If Plugin.PlayerNoteWithHeadshot Then
				If headshotExists Then
					Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PLAYER_GOAL_MINUTES_NOTE2")
				Else
					Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PLAYER_GOAL_MINUTES_NOTE")
				End If
			Else
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PLAYER_GOAL_MINUTES_NOTE")
			End If
		End If

	End With 
End Sub

'Sub UpdatePlayerNote_Next()
'	Dim objplayerNote
'	Dim TeamName
'	Dim teamAbbreviation
'	Dim tmp
'	Dim headshotFileName
'
'	If Interface.Player Is Nothing Then
'		Exit Sub
'	End If
'
'	With Interface.Player
'		TeamName = .Team.DisplayAbbreviation
'		teamAbbreviation = .Team.Abbreviation
'		
'		'TRI_LASTNAME_FIRSTNAME_960
'		headshotFileName = teamAbbreviation & "_" & .LastName & "_" & .FirstName & "_960" 
'
'		Set objplayerNote = .Notes.GetNoteByTitle(Plugin.PlayerNoteSelected)
'		
'		If objplayerNote Is Nothing  Then
'			Exit Sub	
'		End If
'
'		tmp = Split(objplayerNote.Body, "\")
'
'		Call Set_PlayerTemplate_LIVE(.DisplayName, .Jersey, "", TeamName)
'		
'		If Plugin.PlayerNoteWithHeadshot Then
'			'logo 7 head
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERFLAG_LIVE=" & teamAbbreviation)
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERHEAD_LIVE=" & HEADSHOT_DIRECTORY_GH & teamAbbreviation & "/" & headshotFileName)
'		Else
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERFLAG_LIVE=")
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYERHEAD_LIVE=")
'		End If
'	
'		If Ubound(tmp) > 0 Then
'			'2 line notes
'			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LIVE=" & HEADER_DOUBLEQUOTES & tmp(0) & HEADER_DOUBLEQUOTES)
'			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE2_LIVE=" & HEADER_DOUBLEQUOTES & tmp(1) & HEADER_DOUBLEQUOTES)
'			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PLAYER_GOAL_MINUTES_NOTE2")
'			
'			'wiht head it is always 2 line
'			If Plugin.PlayerNoteWithHeadshot Then
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LIVE=" & HEADER_DOUBLEQUOTES & tmp(0) & HEADER_DOUBLEQUOTES)
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE2_LIVE=" & HEADER_DOUBLEQUOTES & tmp(1) & HEADER_DOUBLEQUOTES)
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LIVE=PLAYER_GOAL_MINUTES_NOTE2")
'			Else
'				'msgbox "one line no head"
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LIVE=" & HEADER_DOUBLEQUOTES & tmp(0) & HEADER_DOUBLEQUOTES)
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE2_LIVE=" & HEADER_DOUBLEQUOTES & tmp(1) & HEADER_DOUBLEQUOTES)
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LIVE=PLAYER_GOAL_MINUTES_NOTE2")
'			End If
'
'			
'		Else
'			'l line note
'			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LIVE=" & HEADER_DOUBLEQUOTES & tmp(0) & HEADER_DOUBLEQUOTES) 'this makes it animate
'			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE2_LIVE=") '
'			
'			If Plugin.PlayerNoteWithHeadshot Then
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LIVE=" & HEADER_DOUBLEQUOTES & tmp(0) & HEADER_DOUBLEQUOTES) 'this makes it animate
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE2_LIVE=") '
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LIVE=PLAYER_GOAL_MINUTES_NOTE2")
'			Else
'				'msgbox "one line no head"
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE_LIVE=" & HEADER_DOUBLEQUOTES & tmp(0) & HEADER_DOUBLEQUOTES) 'this makes it animate
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PLAYER_GOAL_MINUTES/NOTE2_LIVE=") '
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LIVE=PLAYER_GOAL_MINUTES_NOTE")
'			End If
'		End If
'
'	End With 
'End Sub
'
'=====================================================================
'	DROPDOWN EXTENSIONS
'---------------------------------------------------------------------
Sub Dropdown_SetExtension(p_sExtensionName, p_fState)

	Select Case p_sExtensionName
		Case "Footer"
			If p_fState Then
				Select Case FoxBox.State
					Case 1, 2
					Case Else
						Footer_Update
						Insert_Footnote	
				End Select				
			Else
				Retract_Footnote
			End If

		Case "Sponsor"
			'Dim sponsorName 
			'sponsorName = Plugin.BillboardSponsorSelected
			
			'If p_fState Then

				'If FoxBox.State <> 1 Then
				'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SPONSOR_IMAGE=" &  PluginSettings.SponsorDirectory & sponsorName)	
				'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SPONSOR=ON")
				'	Plugin.DropDownSponsorState = True
				'End If
			'Else
				'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SPONSOR=OFF")
				'Plugin.DropDownSponsorState = False
			'End If

'		Case "Shootout"
'
'			If p_fState Then
'				
'				If Game.Shootout Then
'					Initialize_ShootoutScore
'					Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=PENALTY_KICKS")
'				End If
'			Else
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=OFF")
'			End If
'
	End Select
End Sub

Sub Insert_Footnote()
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOOTNOTE=ON")
	FootnoteInsertionState = True
	'Viz_Update "FOOTNOTE", "ON"
End Sub

Sub Retract_Footnote()
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOOTNOTE=OFF")
	FootnoteInsertionState = False
	'Viz_Update "FOOTNOTE", "OFF"
End Sub

Sub Dropdown_InsertExtension(p_sExtensionName)
	Log.LogEvent "Dropdown Extension: Inserting " & p_sExtensionName & " Dropdown Extension", "Debug", 0, "Renderer"
	Dropdown_SetExtension p_sExtensionName, True
End Sub

' FoxBox Remove Extension:
'       Controls FoxBox states
Sub Dropdown_RemoveExtension(p_sExtensionName)
	Log.LogEvent "Dropdown Extension: Removing " & p_sExtensionName & " Dropdown Extension", "Debug", 0, "Renderer"
	Dropdown_SetExtension p_sExtensionName, False
End Sub

Sub Insert_DropdownSponsor()
	Dim sponsorName 
	sponsorName = Plugin.BillboardSponsorSelected

	If FoxBox.State <> 1 Then
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SPONSOR_IMAGE=" &  PluginSettings.SponsorDirectory & sponsorName)	
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SPONSOR=ON")
		Plugin.DropDownSponsorState = True
	End If
End Sub

Sub Retract_DropdownSponsor()
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SPONSOR=OFF")
	Plugin.DropDownSponsorState = False
End Sub


'=====================================================================
'	10 Players
'---------------------------------------------------------------------
Sub Insert_TenPlayers_Home()
	Viz_Update "TENPLAYER/T1_LOAD", Plugin.NumberOfPlayersHome & " PLAYERS"
	Viz_Update "TENPLAYERVEIL", "ON"
	Plugin.TenPlayersInsertedHome = True
	'Font color is set in Set_FontColor In GameAndClockUpdate.vbs
End Sub

Sub Remove_TenPlayers_Home()
	Viz_Update "TENPLAYER/T1_LOAD", ""
	Viz_Update "TENPLAYERVEIL", "OFF"
	Plugin.TenPlayersInsertedHome = False
End Sub

Sub Insert_TenPlayers_Visitors()
	Viz_Update "TENPLAYER/T2_LOAD", Plugin.NumberOfPlayersVisitors & " PLAYERS"
	Viz_Update "TENPLAYERVEIL2", "ON"
	Plugin.TenPlayersInsertedVisitors = True
	'Font color is set in Set_FontColor In GameAndClockUpdate.vbs
End Sub

Sub Remove_TenPlayers_Visitors()
	Viz_Update "TENPLAYER/T2_LOAD", ""
	Viz_Update "TENPLAYERVEIL2", "OFF"
	Plugin.TenPlayersInsertedVisitors = False
End Sub


Sub Update_PlayerCountIndicator_Home(p_fState)
	If p_fState Then
		Viz_Update "TEAMLREDFLAG", "ON"
		Plugin.TenPlayersFlagHome = True
	Else
		Viz_Update "TEAMLREDFLAG", "OFF"
		Plugin.TenPlayersFlagHome = False
	End If
End Sub

Sub Update_PlayerCountIndicator_Visitors(p_fState)
	If p_fState Then
		Viz_Update "TEAMRREDFLAG", "ON"
		Plugin.TenPlayersFlagVisitors = True
	Else
		Viz_Update "TEAMRREDFLAG", "OFF"
		Plugin.TenPlayersFlagVisitors = False
	End If
End Sub

'=====================================================================
'	EXTENSIONS
'---------------------------------------------------------------------
' FoxBox Set Extension:
'       Controls FoxBox states
Sub FoxBox_SetExtension(p_sExtensionName, p_fState)
   	
	If p_fState Then
		Log.LogEvent "FoxBox: Inserting " & p_sExtensionName & " " & p_fState, "Debug", 0, "Renderer"
	Else
		Log.LogEvent "FoxBox: Retracting " & p_sExtensionName & " " & p_fState, "Debug", 0, "Renderer"
	End If

	Select Case p_sExtensionName
	
		Case "StoppageTime"	

			If p_fState Then
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET STOPPAGE_TIME=+" & Game.StoppageMinutes )
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SIDEPOD_REQUEST=ST")
			Else
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SIDEPOD_REQUEST=STOFF")
			End If

		Case "ExtraTime"
			Dim etText

			If p_fState Then
				
				If PluginSettings.ETTextType = 0 Then
					etText = ET_TEXT
				Else
					etText = ET_SPANISH_TEXT
				End If

				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET ET_TEXT=" & etText)	
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SIDEPOD_REQUEST=ET")

			Else
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SIDEPOD_REQUEST=ETOFF")
				
			End If

		Case "Aggregate"
			Dim aggText
			Dim aggScores
			
			If p_fState Then
			
				aggScores = Game.Home.AggregateScore & "-" & Game.Visitors.AggregateScore

		       		If PluginSettings.AGGTextType = 0 Then
					aggText = AGG_TEXT
				Else
					aggText = AGG_TEXT_SPANISH
				End If	

				Viz_Update "AGG_HEADER", aggText
				Viz_Update "AGG_VALUE", aggScores

				Viz_Update "SIDEPOD_DISPATCH_AGG_LOAD", "SIDEPOD_AGG_ON" 

			Else
				Viz_Update "SIDEPOD_DISPATCH_AGG_LOAD", "SIDEPOD_AGG_OFF" 
			End If

		Case "Sponsor"

			If p_fState Then
				'Sponsor_Set Settings.SponsorFoxBox
				If FoxBox.State <> 1 Then
					Insert_Sponsor
				End If
			Else
				Retract_Sponsor
			End If

		Case "SubstitutionCounts"

			Call Show_SubstitutionCount(p_fState)
		
		Case "HomeTenMen"
			If p_fState Then
				Viz_Update "TEAMLREDFLAG", "ON"
			Else
				Viz_Update "TEAMLREDFLAG", "OFF"
			End If

		Case "VisitorsTenMen"
			If p_fState Then
				Viz_Update "TEAMRREDFLAG", "ON"
			Else
				Viz_Update "TEAMRREDFLAG", "OFF"
			End If

		Case "JerseyColor"
			If p_fState Then
				Viz_Update "JERSEY", "ON"
			Else
				Viz_Update "JERSEY", "OFF"
			End If

		'Case "GameFlow"

		'	If p_fState Then
		'		'GameFlow_Set
		'		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAMEFLOW=ON")
		'	Else
		'		'Ventuz_Invoke ".Extensions.GameFlow.Retract"
		'	End If
		
		'Case "ScoreSponsor"

		'	If FoxBox.State <> 1 Then
		'		Exit Sub
		'	End If

		'	ScoreSponsor_Set

		'	If p_fState Then
		'		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SCORESPON=ON")
		'	Else
		'		'Ventuz_Update ".Extensions.Header.Show", False
		'	End If

		Case "Header"

			If p_fState Then
				UpdateHeader
			Else
			End If

		Case Else
		    Log.LogEvent "FoxBox: Extension Not Found", "Debug", 0, "Renderer"

	End Select

End Sub

Sub UpdateHeader()
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET POPUP_NOTELINE_LIVE=" & HEADER_DOUBLEQUOTES & PluginSettings.FoxBoxHeader & HEADER_DOUBLEQUOTES)
End Sub

' FoxBox Update Extension:
'       Forces all extension to their correct state
Sub FoxBox_UpdateExtensions()

	FoxBox_SetExtension "ExtraTime", FoxBox.Extensions(0)
	FoxBox_SetExtension "StoppageTime", FoxBox.Extensions(1)
	FoxBox_SetExtension "Aggregate", FoxBox.Extensions(2)
	FoxBox_SetExtension "Sponsor", FoxBox.Extensions(3)
	'FoxBox_SetExtension "GameFlow", FoxBox.Extensions(4)
	FoxBox_SetExtension "Header", FoxBox.Extensions(6)


	'FoxBox_SetExtension "TenMenHome", FoxBox.Extensions(6)
	'FoxBox_SetExtension "TenMenVisitors", FoxBox.Extensions(7)
End Sub

' FoxBox Insert Extension:
'       Controls FoxBox states
Sub FoxBox_InsertExtension(p_sExtensionName)
	Log.LogEvent "FoxBox: Inserting " & p_sExtensionName & " Extension", "Debug", 0, "Renderer"
	FoxBox_SetExtension p_sExtensionName, True
End Sub

' FoxBox Remove Extension:
'       Controls FoxBox states
Sub FoxBox_RemoveExtension(p_sExtensionName)
	Log.LogEvent "FoxBox: Removing " & p_sExtensionName & " Extension", "Debug", 0, "Renderer"
	FoxBox_SetExtension p_sExtensionName, False
End Sub


'=====================================================================
'	POSSESSION CHART
'---------------------------------------------------------------------
Sub Insert_PossessionChart()
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=TIME_OF_POSSESSION")
End Sub

Sub Update_PossessionChart()
	Dim objHomeSplit
	Dim objVisitorsSplit
	Dim sHomeStat
	Dim sVisitorsStat
	
	Set objHomeSplit = Game.Home.Stats.Game.Total
	Set objVisitorsSplit = Game.Visitors.Stats.Game.Total
	
	sHomeStat = objHomeSplit.Possession 
	sVisitorsStat = objVisitorsSplit.Possession

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TIME_OF_POSSESSION/HEADER_LOAD=POSSESSION")

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TIME_OF_POSSESSION/T1_ABB_LOAD=" & Game.Home.DisplayAbbreviation)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TIME_OF_POSSESSION/T2_ABB_LOAD=" & Game.Visitors.DisplayAbbreviation)
	
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TIME_OF_POSSESSION/T1_BAR_LOAD=" & sHomeStat)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TIME_OF_POSSESSION/T2_BAR_LOAD=" & sVisitorsStat)
End Sub

'=====================================================================
'	GAME NOTE
'---------------------------------------------------------------------
'This gets invoked by Main body of FoxBox
Sub GameNote_Change()
	'updating UI of Game Notes in Plugin 
	'Plugin.GameNotesRefresh
	
	'commented out for now
	Plugin.NotesRefresh
End Sub

Sub GameNote_Update()
	Dim objNote
	Dim sBody

	Log.LogEvent "GameNote_Update is called", "Debug", 0, "Renderer"
	Log.LogEvent "Note Title Selected: "  & Plugin.GameNoteSelected , "Debug", 0, "Renderer"

	Set objNote = Soccer.Game.Notes.GetNoteByTitle(Plugin.GameNoteSelected)
	
	If Not objNote Is Nothing Then
	
		Log.LogEvent "Note " & Plugin.GameNoteSelected & " FOUND ",  "Debug", 0, "Renderer"
		' Update Note
		sNote = objNote.Body

		Call Set_GameNote(sNote)

	Else
		Log.LogEvent "Note " & Plugin.GameNoteSelected & " NOT FOUND ", "Debug", 0, "Renderer"
	End If
End Sub

Sub Set_GameNote(sNote)

	Dim tmp
	Dim IsLine1Bold

	IsLine1Bold = 0

	tmp = Split(sNote, "\")

	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET LINE1BOLD_LOAD=0")

	'setting team focus to be both
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOCUS_LOAD=2")

	Select Case Ubound(tmp)
		Case 0 'single line
		
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SINGLE/TEXT1_LOAD=" & HEADER_DOUBLEQUOTES & tmp(0))
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=SINGLE")

			Call Set_SingleLineNote(tmp(0), IsLine1Bold)
			Insert_SingleLineNote

			

		Case 1 'double line

			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DOUBLE/TEXT1_LOAD=" & HEADER_DOUBLEQUOTES & tmp(0))
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DOUBLE/TEXT2_LOAD=" & HEADER_DOUBLEQUOTES & tmp(1))
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=DOUBLE")

			Call Set_DoubleLineNote(tmp(0), tmp(1), IsLine1Bold)
			Insert_DoubleLineNote	

		Case 2 'triple line
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT1_LOAD=" & HEADER_DOUBLEQUOTES & tmp(0))
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT2_LOAD=" & HEADER_DOUBLEQUOTES & tmp(1))
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT3_LOAD=" & HEADER_DOUBLEQUOTES & tmp(2))
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=TRIPLE")

			Call Set_TripleLineNote(tmp(0), tmp(1), tmp(2), IsLine1Bold)
			Insert_TripleLineNote
			

		Case Else

	End Select

	If Ubound(tmp) >= 0 AND Ubound(tmp) < 2 Then
		m_iGameNoteCounter = Ubound(tmp) + 1
	Else
		m_iGameNoteCounter = -1
	End If
End Sub

Sub Highlight_GameNote()
	If Plugin.HilightNote Then
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DZ_COLOR=0")
	Else
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DZ_COLOR=1")
	End If
End Sub


Sub Set_SingleLineNote(note1, IsLine1Bold)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET LINE1BOLD_LOAD=" & IsLine1Bold)

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SINGLE/TEXT1_LOAD=" & HEADER_DOUBLEQUOTES & note1 & HEADER_DOUBLEQUOTES)
End Sub

Sub Set_DoubleLineNote(note1, note2, IsLine1Bold)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET LINE1BOLD_LOAD=" & IsLine1Bold)

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DOUBLE/TEXT1_LOAD=" & HEADER_DOUBLEQUOTES & note1 & HEADER_DOUBLEQUOTES)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DOUBLE/TEXT2_LOAD=" & HEADER_DOUBLEQUOTES & note2 & HEADER_DOUBLEQUOTES)
End Sub

Sub Set_TripleLineNote(note1, note2, note3, IsLine1Bold)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET LINE1BOLD_LOAD=" & IsLine1Bold)

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT1_LOAD=" & HEADER_DOUBLEQUOTES & note1 & HEADER_DOUBLEQUOTES)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT2_LOAD=" & HEADER_DOUBLEQUOTES & note2 & HEADER_DOUBLEQUOTES)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT3_LOAD=" & HEADER_DOUBLEQUOTES & note3 & HEADER_DOUBLEQUOTES)
End Sub

Sub Insert_SingleLineNote()
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=SINGLE")
End Sub

Sub Insert_DoubleLineNote()
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=DOUBLE")
End Sub

Sub Insert_TripleLineNote()
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=TRIPLE")
End Sub


'=====================================================================
'	FOOTER
'---------------------------------------------------------------------
Sub Footer_Update()
	Dim objNote
	Dim sBody

	Log.LogEvent "Footer_Update is called", "Debug", 0, "Renderer"
	Log.LogEvent "Footer Title Selected: "  & Plugin.FooterSelected , "Debug", 0, "Renderer"

	Set objNote = Soccer.Game.Notes.GetNoteByTitle(Plugin.FooterSelected)
	
	If Not objNote Is Nothing Then
	
		Log.LogEvent "Footer " & Plugin.FooterSelected & " FOUND ",  "Debug", 0, "Renderer"
		' Update Note
		sNote = objNote.Body

		Call Set_Footer(sNote)

	Else
		Log.LogEvent "Note " & Plugin.FooterSelected & " NOT FOUND ", "Debug", 0, "Renderer"
	End If
End Sub

Sub Set_Footer(sNote)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOOTNOTE_LOAD=" & HEADER_DOUBLEQUOTES & sNote & HEADER_DOUBLEQUOTES)
End Sub

'=====================================================================
'	TEAM STATS
'---------------------------------------------------------------------
Sub TeamStats_Change()

	If FoxBox.State = 4 Then
		TeamComparison_Update
		
	End If

	If CurrentState = "TeamComparison" Then
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET OTS_TRANS=NEXT")
	End If
	
End Sub

Sub TeamComparison_Update()
	Dim objHomeSplit
	Dim objVisitorsSplit
	Dim sHomeStat
	Dim sVisitorsStat
	Dim sTitle
	
	Set objHomeSplit = Game.Home.Stats.Game.Total
	Set objVisitorsSplit = Game.Visitors.Stats.Game.Total
	
	sHomeStat = TeamStats_GetValue(objHomeSplit, Plugin.TeamStat)
	sVisitorsStat = TeamStats_GetValue(objVisitorsSplit, Plugin.TeamStat)

	'If Left(Plugin.TeamStat, 6) = "Custom" Then
	'	Dim iCustomIndex
	'	iCustomIndex = Mid(Plugin.TeamStat, 9, 1)
	'	sTitle = UCase(Settings.CustomStats(iCustomIndex))
	'Else
		sTitle = TeamStats_GetTitle(Plugin.TeamStat)
	'End If

	
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TEAM_STATS_COMPARISON/HEADER_LOAD=" & HEADER_DOUBLEQUOTES & sTitle & HEADER_DOUBLEQUOTES)

	'right stat
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TEAM_STATS_COMPARISON/TEXT1_LOAD=" & sHomeStat)
	' left stat
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TEAM_STATS_COMPARISON/TEXT2_LOAD=" & sVisitorsStat)
	
	
End Sub



Sub Insert_TeamComparison()
	
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=TEAM_STATS_COMPARISON_SINGLE")
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOCUS_LOAD=2")
End Sub

' Team Stats - Get Value:
'		Get specified stat from given split
Function TeamStats_GetValue(p_objSplit, p_sStat)
    Dim sStat

    Select Case UCase(p_sStat)
        Case "POSSESSION"
            sStat = p_objSplit.Possession & "%" '& Chr(60) & "scale value" & Chr(61) & Chr(34) & "0.73" & Chr(34) & Chr(62) & "%" & chr(60) & "/scale" & Chr(62)
        Case "YELLOW CARDS"
            sStat = p_objSplit.YellowCards
        Case "RED CARDS"
            sStat = p_objSplit.RedCards
        Case "PASS COMP"
            sStat = p_objSplit.PassesCompleted
        Case "PASS ATT"
            sStat = p_objSplit.PassesAttempted
        Case "PASS %"
            'sStat = FormatNumber(100 * p_objSplit.PassingPercentage, 1) & "%"  '& Chr(60) & "scale value" & Chr(61) & Chr(34) & "0.73" & Chr(34) & Chr(62) & "%"  & chr(60) & "/scale" & Chr(62)
            sStat = FormatNumber(100 * p_objSplit.PassingPercentage, 0) & "%"  '& Chr(60) & "scale value" & Chr(61) & Chr(34) & "0.73" & Chr(34) & Chr(62) & "%"  & chr(60) & "/scale" & Chr(62)
        Case "SHOTS"
            sStat = p_objSplit.Shots
        Case "SHOTS ON GOAL"
            sStat = p_objSplit.ShotsOnGoal
        Case "FOULED"
            sStat = p_objSplit.FoulsSuffered
        Case "OFFSIDES"
            sStat = p_objSplit.Offsides
        Case "CORNERS"
            sStat = p_objSplit.CornerKicks
        Case "FREE KICKS"
            sStat = p_objSplit.FreeKicks
        Case "SAVES"
            sStat = p_objSplit.Saves
        Case "TACKLES"
            sStat = p_objSplit.Tackles
        Case "CLEARANCES"
            sStat = p_objSplit.Clearances
        Case "INTERCEPTIONS"
            sStat = p_objSplit.Interceptions
        Case "BLOCKS"
            sStat = p_objSplit.ShotsBlocked
        Case "FOULS"
            sStat = p_objSplit.FoulsCommitted

	Case Ucase(Settings.CustomStats(1))
            sStat = p_objSplit.Custom(1)
	Case UCase(Settings.CustomStats(2))
            sStat = p_objSplit.Custom(2)
        'Case "CUSTOM #3"
	Case Ucase(Settings.CustomStats(3))
            sStat = p_objSplit.Custom(3)
        'Case "CUSTOM #4"
	Case Ucase(Settings.CustomStats(4))
            sStat = p_objSplit.Custom(4)
    End Select

	TeamStats_GetValue = sStat
	
End Function

' Team Stats - Get Title:
'		Get specified title for given stat
Function TeamStats_GetTitle(p_sStat)
    Dim sTitle
    
    Select Case UCase(p_sStat)
        Case "POSSESSION"
            sTitle = "POSSESSION"
        Case "YELLOW CARDS"
            sTitle = "YELLOW CARDS"
        Case "RED CARDS"
            sTitle = "RED CARDS"
        Case "PASS COMP"
            sTitle = "PASSES COMPLETED"
        Case "PASS ATT"
            sTitle = "PASSES ATTEMPTED"
        Case "PASS %"
            sTitle = "PASS COMPLETION"
        Case "SHOTS"
            sTitle = "SHOTS"
        Case "SHOTS ON GOAL"
            sTitle = "SHOTS ON GOAL"
        Case "FOULED"
            sTitle = "FOULED"
        Case "OFFSIDES"
            sTitle = "OFFSIDES"
        Case "CORNERS"
            sTitle = "CORNERS"
        Case "FREE KICKS"
            sTitle = "FREE KICKS"
        Case "SAVES"
            sTitle = "SAVES"
        Case "TACKLES"
            sTitle = "TACKLES"
        Case "CLEARANCES"
            sTitle = "CLEARANCES"
        Case "INTERCEPTIONS"
            sTitle = "INTERCEPTIONS"
        Case "BLOCKS"
            sTitle = "BLOCKS"
        Case "FOULS"
            sTitle = "FOULS"
	Case Else
	    sTitle = UCase(p_sStat)
    End Select

	TeamStats_GetTitle = sTitle
	
End Function

Function Get_SelectedTeamStat(selectedIndex)
	Dim sSelectedStat
	
	Select Case selectedIndex
		Case 0 
			sSelectedStat =  "NONE"
		Case 1
			sSelectedStat =  "POSSESSION"
		Case 2
			sSelectedStat = "YELLOW CARDS"
		Case 3
			sSelectedStat = "RED CARDS"
		Case 4
			sSelectedStat = "PASS COMP"
		Case 5
			sSelectedStat = "PASS ATT"
		Case 6
			sSelectedStat = "PASS %"
		Case 7
			sSelectedStat = "SHOTS"
		Case 8
			sSelectedStat = "SHOTS ON GOAL"
		Case 9
			sSelectedStat = "OFFSIDES"
		Case 10
			sSelectedStat =  "CORNERS"
		Case 11
			sSelectedStat = "FREE KICKS"
		Case 12
			sSelectedStat = "SAVES"
		Case 13
			sSelectedStat = "FOULS"
		Case 14
			sSelectedStat = Settings.CustomStats(1)
		Case 15
			sSelectedStat = Settings.CustomStats(2)
		Case 16
			sSelectedStat = Settings.CustomStats(3)
		Case 17
			sSelectedStat = Settings.CustomStats(4)

		'Case 9
			'sSelectedStat = "FOULED"
		'Case 14
		'	sSelectedStat =  "TACKLES"
		'Case 15
		'	sSelectedStat = "CLEARANCES"
		'Case 16
		'	sSelectedStat = "INTERCEPTIONS"
		'Case 17
			'sSelectedStat = "BLOCKS"
		

	End Select

	Get_SelectedTeamStat = sSelectedStat

End Function 

Function Get_TeamStat(objSplit, statType)
	Dim stat
	stat = TeamStats_GetValue(objSplit, statType)

	Get_TeamStat = stat
End Function


Sub Insert_MultilineTeamStats()
	Dim numStats
	Dim sSelectedStat1, sSelectedStat2, sSelectedStat3
	Dim objHomeSplit, objVisitorsSplit
	Dim row1stat1, row1stat2, row2stat1, row2stat2, row3stat1, row3stat2

	numStats = Get_NumberOfTeamStats

	Set objHomeSplit = Game.Home.Stats.Game.Total
	Set objVisitorsSplit = Game.Visitors.Stats.Game.Total
	
	dim statTitle
	dim team1stat, team2stat

	With PluginSettings

		For i = 0 to numStats - 1
			statTitle = Get_SelectedTeamStat(.TeamStatsSelected(i))
			team1stat = Get_TeamStat(objHomeSplit, statTitle)
			team2stat = Get_TeamStat(objVisitorsSplit, statTitle)
			
			Call Set_TeamComparionsStats(statTitle, team1stat, team2stat, i+1)
		Next

		'sSelectedStat1 = Get_SelectedTeamStat(.TeamStatsSelected(0))
		'row1stat1 = Get_TeamStat(objHomeSplit, sSelectedStat1)
		'row1stat2 = Get_TeamStat(objVisitorsSplit, sSelectedStat1)
		'Call Set_TeamComparionsStats(sSelectedStat1, row1stat1, row1stat2, 1)

		'sSelectedStat2 = Get_SelectedTeamStat(.TeamStatsSelected(1))
		'row2stat1 = Get_TeamStat(objHomeSplit, sSelectedStat2)
		'row2stat2 = Get_TeamStat(objVisitorsSplit, sSelectedStat2)
		'Call Set_TeamComparionsStats(sSelectedStat2, row2stat1, row2stat2, 2)

		'sSelectedStat3 = Get_SelectedTeamStat(.TeamStatsSelected(2))
		'row3stat1 = Get_TeamStat(objHomeSplit, sSelectedStat3)
		'row3stat2 = Get_TeamStat(objVisitorsSplit, sSelectedStat3)
		'Call Set_TeamComparionsStats(sSelectedStat3, row3stat1, row3stat2, 3)
		
		Select Case numStats
			Case 1
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=TEAM_STATS_COMPARISON_SINGLE")
			Case 2
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=TEAM_STATS_COMPARISON_DOUBLE")
			Case 3
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=TEAM_STATS_COMPARISON_TRIPLE")
		End Select

	End With
End Sub

Function Get_NumberOfTeamStats()
	
	Dim sSelectedStat
	Dim numStatLines
	Dim i

	numStatLines = 0

	With PluginSettings
		'Loop from 0 to 2
		For i = 0 to 2

			sSelectedStat = Get_SelectedTeamStat(.TeamStatsSelected(i))

			if sSelectedStat = "NONE" Then
				Exit For
			Else
				numStatLines = numStatLines + 1
			End If
		Next

	End With

	Get_NumberOfTeamStats = numStatLines


End Function

Sub Set_TeamComparionsStats(title, team1stat, team2stat, statRow)
	
	If statRow = 1 Then
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TEAM_STATS_COMPARISON/HEADER_LOAD=" & title)
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TEAM_STATS_COMPARISON/TEXT1_LOAD=" & team1stat)
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TEAM_STATS_COMPARISON/TEXT2_LOAD=" & team2stat)
	Else
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TEAM_STATS_COMPARISON/HEADER_" & statRow & "_LOAD=" & title)
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TEAM_STATS_COMPARISON/TEXT1_" & statRow & "_LOAD=" & team1stat)
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TEAM_STATS_COMPARISON/TEXT2_" & statRow & "_LOAD=" & team2stat)
	End If


End Sub
'=====================================================================
'	JERSEY COLORS
'---------------------------------------------------------------------
Sub Insert_JerseyColors()
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=JERSEY")	
End Sub

Sub Update_JerseyColors
	
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET JERSEY/HEADER_LOAD=TEAM COLORS")
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET JERSEY/T1_LOAD=" & Game.Home.Name)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET JERSEY/T2_LOAD=" & Game.Visitors.Name)

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET Team2Color=" &  Ucase(Game.Visitors.DisplayColor))
	
	With PluginSettings
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TEAM1_TEXTCOLOR=" & .HomeFontColor)
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TEAM2_TEXTCOLOR=" & .VisitorsFontColor)
	End With
End Sub 


'=====================================================================
'	COMMENTATORS	
'---------------------------------------------------------------------
Sub Insert_Commentators()
	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=COMMENTATORS")	
	UpdateCommentators
End Sub

'=====================================================================
'	Commentators	
'---------------------------------------------------------------------
'Sub UpdateCommentators_RegularFoxBox()
'	
'	Dim i
'
'	ClearCommentators
'
'	Viz_Update "COMMENTATORS/HEADER_LOAD", "Commentators"
'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOCUS_LOAD=2")
'
'	Dim com1, com2, com3
'	Dim result
'	
'	com1 = Trim(Settings.Commentators(1).FirstName) '& chr(32) & Trim(Settings.Commentators(1).LastName)
'	com2 = Trim(Settings.Commentators(2).FirstName) '& chr(32) & Trim(Settings.Commentators(2).LastName)
'	com3 = Trim(Settings.Commentators(3).FirstName) '& chr(32) & Trim(Settings.Commentators(3).LastName)
'	com4 = Trim(Settings.Commentators(4).FirstName) '& chr(32) & Trim(Settings.Commentators(4).LastName) 
'	
'	
'	'Checking First Name exists
'	If Len(com1) > 0 Then
'		Viz_Update "COMMENTATORS/NAME1_LOAD", HEADER_DOUBLEQUOTES & com1 & chr(32) & Trim(Settings.Commentators(1).LastName)
'		
'		'There is only 1 commentator
'		If Len(com2) < 1 Then
'			Viz_Update "COMMENTATORS/NAME1_LOAD", ""
'			Viz_Update "COMMENTATORS/NAME2_LOAD", HEADER_DOUBLEQUOTES & com1 & chr(32) & Trim(Settings.Commentators(1).LastName)
'		End If
'	end if
'
'	If Len(com2) > 0  Then
'		Viz_Update "COMMENTATORS/NAME2_LOAD", HEADER_DOUBLEQUOTES & com2 & chr(32) & Trim(Settings.Commentators(2).LastName) 	
'	end If
'	
'	If Len(com3) > 0 Then
'		Viz_Update "COMMENTATORS/NAME3_LOAD", HEADER_DOUBLEQUOTES & com3 & chr(32) & Trim(Settings.Commentators(3).LastName) 
'	End If
'		
'	Insert_Commentators
'
'	If Len(com4) > 0 Then
'		result = MsgBox("There is One More Commentator.  Move Next?", vbYesNo)
'
'		If result = vbYes Then
'			AdvanceCommentators
'		Else
'			Exit Sub
'		End If
'	End If
'	
'
'End Sub

'this needs to clean up
Sub UpdateCommentators()

	
	Dim com1, com2, com3
	Dim result
	Dim commentator_header
	Dim IsLine1Bold


	IsLine1Bold = 1
	commentator_header = "COMMENTATORS"
	
	com1 = Trim(Settings.Commentators(1).FirstName) '& chr(32) & Trim(Settings.Commentators(1).LastName)
	com2 = Trim(Settings.Commentators(2).FirstName) '& chr(32) & Trim(Settings.Commentators(2).LastName)
	com3 = Trim(Settings.Commentators(3).FirstName) '& chr(32) & Trim(Settings.Commentators(3).LastName)
	com4 = Trim(Settings.Commentators(4).FirstName) '& chr(32) & Trim(Settings.Commentators(4).LastName


	'Header Bold
	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET LINE1BOLD_LOAD=1") 'bolding line 1
	
	If Len(com1) > 0 Then
		If Len(com2) > 0 Then
			'at least 2 commentators
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT1_LOAD=" & commentator_header)
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT2_LOAD=" & HEADER_DOUBLEQUOTES & com1 & chr(32) & Trim(Settings.Commentators(1).LastName))
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT3_LOAD=" & HEADER_DOUBLEQUOTES & com2 & chr(32) & Trim(Settings.Commentators(2).LastName))
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=TRIPLE")

			Call Set_TripleLineNote(commentator_header, com1 & chr(32) & Trim(Settings.Commentators(1).LastName), com2 & chr(32) & Trim(Settings.Commentators(2).LastName), IsLine1Bold)
			Insert_TripleLineNote

		Else
			'One commentator
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DOUBLE/TEXT1_LOAD=" & commentator_header)
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DOUBLE/TEXT2_LOAD=" & HEADER_DOUBLEQUOTES & com1 & chr(32) & Trim(Settings.Commentators(1).LastName))
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=DOUBLE")
			commentator_header = "COMMENTATOR"
			
			Call Set_DoubleLineNote(commentator_header, com1 & chr(32) & Trim(Settings.Commentators(1).LastName), IsLine1Bold)
			Insert_DoubleLineNote
		End If
	End If


	If Len(com3) > 0 Then

		result = MsgBox("More Commentators.  Move Next?", vbYesNo)
		
		If result = vbYes Then
			If Len(com4) > 0 Then
			
				'at least 2 commentators
				'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT1_LOAD=" & commentator_header)
				'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT2_LOAD=" & HEADER_DOUBLEQUOTES & com3 & chr(32) & Trim(Settings.Commentators(3).LastName))
				'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT3_LOAD=" & HEADER_DOUBLEQUOTES & com4 & chr(32) & Trim(Settings.Commentators(4).LastName))
				'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=TRIPLE")

				Call Set_TripleLineNote(commentator_header, com3 & chr(32) & Trim(Settings.Commentators(3).LastName), com4 & chr(32) & Trim(Settings.Commentators(4).LastName), IsLine1Bold)
				Insert_TripleLineNote
			Else
				'One commentator
				'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DOUBLE/TEXT1_LOAD=" & commentator_header)
				'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DOUBLE/TEXT2_LOAD=" & HEADER_DOUBLEQUOTES & com3 & chr(32) & Trim(Settings.Commentators(3).LastName))
				'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=DOUBLE")
				
				Call Set_DoubleLineNote(commentator_header, com3 & chr(32) & Trim(Settings.Commentators(3).LastName), IsLine1Bold)
				Insert_DoubleLineNote
		End If
		
		Else 'answer is NO
			Exit Sub
		End If
	End If
End Sub


Sub AdvanceCommentators()
	Dim com4

	ClearCommentators
	
	Viz_Update "COMMENTATORS/HEADER_LOAD", "Commentators"

	com4 = Trim(Settings.Commentators(4).FirstName) & chr(32) & Trim(Settings.Commentators(4).LastName)
	Viz_Update "COMMENTATORS/NAME2_LOAD", HEADER_DOUBLEQUOTES & com4 & HEADER_DOUBLEQUOTES
End Sub

Sub ClearCommentators()
	Dim i
	
	Viz_Update "COMMENTATORS/HEADER_LOAD", ""
	
	For i = 1 to 3
		Viz_Update "COMMENTATORS/NAME" & i & "_LOAD", ""
	Next
End Sub


Sub UpdateCommentatorsWithHead()
	Dim i 
	dim result

	With PluginSettings

	For i = 1 to 4

		If Len(Trim(.Commentators(i).FirstName)) > 0 Then
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TALENT/TEXT1_LOAD=" & .CommentatorsTitles(i))
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TALENT/TEXT2_LOAD=" & .Commentators(i).FirstName & " " & .Commentators(i).LastName)
			
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TALENT/PLAYERHEAD_LOAD=" & .CommentatorsHeadshots(i))

			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=TALENT")

			If Len(Trim(.Commentators(i+1).FirstName)) > 0 And (i < 4)Then
				result = MsgBox("More Talents.  Move Next?", vbYesNo)
				
				If result = vbYes Then
					'contintue
				Else
					Exit Sub 
				End if
			Else
				Exit for

			End If


		Else
			Exit For
		End If
	Next
	

	End With
End Sub

Sub UpdateCommentatorsWithHead_BackupInProgress()
	Dim i 
	dim result

	With PluginSettings
	
		If Len(Trim(.Commentators(1).FirstName)) > 0 Then
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TALENT/TEXT1_LOAD=" & .Commentators(1).FirstName & " " & .Commentators(1).LastName)
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TALENT/TEXT2_LOAD=")
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TALENT/PLAYERHEAD_LOAD=" & .CommentatorsHeadshots(1))

			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=TALENT")
		End If

		If Len(Trim(.Commentators(2).FirstName)) > 0 Then
			result = MsgBox("More Talents.  Move Next?", vbYesNo)
				
			If result = vbYes Then
				'contintue
			Else
				Exit Sub 
			End if
			
		End If

		For i = 2 to 4

			If Len(Trim(.Commentators(i).FirstName)) > 0 Then
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TALENT/TEXT1_LOAD=" & .Commentators(i).FirstName & " " & .Commentators(i).LastName)
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TALENT/TEXT2_LOAD=")
				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TALENT/PLAYERHEAD_LOAD=" & .CommentatorsHeadshots(i))

			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=TALENT")

			If Len(Trim(.Commentators(i+1).FirstName)) > 0 And (i < 4)Then
				result = MsgBox("More Talents.  Move Next?", vbYesNo)
				
				If result = vbYes Then
					'contintue
				Else
					Exit Sub 
				End if
			Else
				Exit for

			End If


		Else
			Exit For
		End If
	Next
	

	End With
End Sub


'Sub UpdateCommentatorsWithHeadshots()
'	Dim i
'	Dim countCommentators
'	Dim objFileSystem
'	Dim sFileName
'
'	countCommentators = 0
'
'	' Create FileSystem Object
'	Set objFileSystem = CreateObject("Scripting.FileSystemObject")
'
'	' Get Headshot Folder
'	sFileName = COMMON_DIRECTORY & "Elements\Headshots\"
'
'	Set objFolder = objFileSystem.GetFolder(sFileName)
'                  
'	If objFolder Is Nothing Then
'		Log.LogEvent "Headshot folder '" & sFileName & "' doesn't exist.", "Debug", 0, "Renderer"    
'	End If
'
'
'	For i = 1 to 4
'		If  Len(Settings.Commentators(i).FirstName) > 0 Then
'			Ventuz_Update ".Dropdown.Talent.Item" & i & ".Name1", Trim(Settings.Commentators(i).FirstName)
'			Ventuz_Update ".Dropdown.Talent.Item" & i & ".Name2", Trim(Settings.Commentators(i).LastName)
'			countCommentators = countCommentators + 1
'		End If
'		
'
'		'Settignb headshots
'                If Not objFolder Is Nothing Then
'			sFileName = objFolder.Path & "\" & Settings.Commentators(i).FirstName & " " & Settings.Commentators(i).LastName & ".png"
'
'			If objFileSystem.FileExists(sFileName) Then
'				Ventuz_Update ".Dropdown.Talent.Item" & i & ".ImagePath", sFileName
'			End If
'		End If
'	Next
'
'
'	If countCommentators > 3 then
'		countCommentators = 3
'	End If
'
'	Ventuz_Update ".Dropdown.Talent.Count", countCommentators
'
'End Sub
'=====================================================================
'	Reporting	
'---------------------------------------------------------------------
Sub Insert_Reporting()
	UpdateReporting

End Sub

'Sub UpdateReporting_Regular()
'
'	
'	Dim Reporting1, Reporting2
'
'	
'	ClearCommentators
'	
'	
'	Viz_Update "COMMENTATORS/HEADER_LOAD", Settings.ReportingTitle
'	
'
'	If  Len(Settings.Reporting(1).FirstName) > 0 Then
'		
'		Reporting1 = Trim(Settings.Reporting(1).FirstName) & " " & Trim(Settings.Reporting(1).LastName)
'		Viz_Update "COMMENTATORS/NAME1_LOAD", HEADER_DOUBLEQUOTES & Reporting1
'
'		'there is only one reporter
'		If Len(Settings.Reporting(2).FirstName) < 1 Then
'			Viz_Update "COMMENTATORS/NAME1_LOAD", ""
'			Viz_Update "COMMENTATORS/NAME2_LOAD", HEADER_DOUBLEQUOTES & Reporting1
'		End If
'	End If
'
'		
'	If Len(Settings.Reporting(2).FirstName) > 0 Then			
'		Reporting2 = Trim(Settings.Reporting(2).FirstName) & " " & Trim(Settings.Reporting(2).LastName)
'		Viz_Update "COMMENTATORS/NAME2_LOAD", HEADER_DOUBLEQUOTES & Reporting2
'	End If
'
'End Sub

Sub UpdateReporting()

	Dim Reporting1, Reporting2
	Dim reportingTitle
	Dim IsLine1Bold

	IsLine1Bold = 1

	If  Len(Settings.Reporting(1).FirstName) > 0 Then
		
		Reporting1 = Trim(Settings.Reporting(1).FirstName) & " " & Trim(Settings.Reporting(1).LastName)
	
	
		'there are 2 reporters
		If Len(Settings.Reporting(2).FirstName) > 0 Then

			Reporting2 = Trim(Settings.Reporting(2).FirstName) & " " & Trim(Settings.Reporting(2).LastName)

			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT1_LOAD=" & Settings.ReportingTitle)
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT2_LOAD=" & HEADER_DOUBLEQUOTES & Reporting1)
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT3_LOAD=" & HEADER_DOUBLEQUOTES & Reporting2)
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=TRIPLE")	

			Call Set_TripleLineNote(Settings.ReportingTitle, Reporting1, Reporting2, IsLine1Bold)
			Insert_TripleLineNote
		Else
		'there is one reorter
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DOUBLE/TEXT1_LOAD=" & Settings.ReportingTitle)
			'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DOUBLE/TEXT2_LOAD=" & HEADER_DOUBLEQUOTES & Reporting1)
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=DOUBLE")	

			Call Set_DoubleLineNote(Settings.ReportingTitle, Reporting1, IsLine1Bold)
			Insert_DoubleLineNote
		End If

	End If
	
	
End Sub

'=====================================================================
'	Locator
'---------------------------------------------------------------------
Sub Insert_Locator()

	UpdateLocator
	
End sub

'Sub UpdateLocator_RegularSeason()
'
'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOCUS_LOAD=2")
'
'	If Trim(Settings.LocatorTitle) <> "" Then
'		Viz_Update "LOCATION/LOCATION_LOAD", Chr(34)& Chr(34) & Settings.LocatorTitle
'	Else
'		Viz_Update "LOCATION/LOCATION_LOAD", Chr(34)& Chr(34) & Ucase(Game.Home.Stadium)
'	End If
'
'
'	If Trim(Settings.LocatorLocation) <> "" Then
'		Viz_Update "LOCATION/CITY_LOAD", Chr(34)& Chr(34) & Settings.LocatorLocation
'	Else
'		Viz_Update "LOCATION/CITY_LOAD", Chr(34)& Chr(34) & UCase(Game.Home.CityAndState)
'	End If
'End Sub

Sub UpdateLocator()
	Dim locatorTile
	Dim locatorLocation
	Dim IsLine1Bold

	IsLine1Bold = 0

	If Trim(Settings.LocatorTitle) <> "" Then
		locatorTile =  Settings.LocatorTitle
	Else
		locatorTile =   Ucase(Game.Home.Stadium)
	End If
	
	
	If Trim(Settings.LocatorLocation) <> "" Then
		locatorLocation = Settings.LocatorLocation
	Else
		locatorLocation = UCase(Game.Home.CityAndState)
	End If
	
	'taking off bold
	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET LINE1BOLD_LOAD=0")
	
	If Len(locatorTile) > 0 Then
		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DOUBLE/TEXT1_LOAD=" & HEADER_DOUBLEQUOTES & locatorTile)
		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DOUBLE/TEXT2_LOAD=" & HEADER_DOUBLEQUOTES & locatorLocation)
		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=DOUBLE")

		Call Set_DoubleLineNote(locatorTile, locatorLocation, IsLine1Bold)
		Insert_DoubleLineNote
	End If

End Sub

Sub UpdateVoiceOf()
	Dim voiceOfNote
	Dim voiceOfTitle
	Dim voiceOfFullName
	Dim IsLine1Bold

	IsLine1Bold = 1

	voiceOfTitle = "VOICE OF"

	voiceOfFullName = Settings.VoiceOf.Firstname & " " & Settings.VoiceOf.Lastname 
	
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET LINE1BOLD_LOAD=1")

	If Trim(Settings.VoiceOfNote) <> "" Then
		
		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT1_LOAD=" & voiceOfTitle)
		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT2_LOAD=" & HEADER_DOUBLEQUOTES & voiceOfFullName)
		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TRIPLE/TEXT3_LOAD=" & HEADER_DOUBLEQUOTES & Settings.VoiceOfNote)
		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=TRIPLE")	

		Call Set_TripleLineNote(voiceOfTitle, voiceOfFullName, Settings.VoiceOfNote, IsLine1Bold)
		Insert_TripleLineNote
	Else
		

		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DOUBLE/TEXT1_LOAD=" & HEADER_DOUBLEQUOTES & voiceOfTitle)
		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DOUBLE/TEXT2_LOAD=" & HEADER_DOUBLEQUOTES & voiceOfFullName)
		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=DOUBLE")

		Call Set_DoubleLineNote(voiceOfTitle, voiceOfFullName, IsLine1Bold)
		Insert_DoubleLineNote
	End If
	
End Sub

Sub UpdateVoiceOffWithHead()
	Dim voiceOfFname
	Dim voiceOffLname
	Dim headPath

	With Settings
		voiceOfFname = .VoiceOf.Firstname
		voiceOffLname = .VoiceOf.Lastname 
		
		
		'If Plugin.TalentHeadIsManual Then
			headPath = Plugin.TalentMeadManualPath
		'Else
		'	headPath = "WC_" & Ucase(voiceOffLname) & "_" & UCase(voiceOfFname) & "_960"
			
		'End If
		
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TALENT/PLAYERHEAD_LOAD=" & headPath)


		'If Plugin.NumberOfRowsVoiceOfWithHeadSelected = 0 Then
			'one line ne head
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TALENT/TEXT1_LOAD=" & voiceOfFname & " " & voiceOffLname)
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TALENT/TEXT2_LOAD=") 
		'Else	
			'2 line head
		'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TALENT/TEXT1_LOAD=" & voiceOfFname)
		'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TALENT/TEXT2_LOAD=" & voiceOffLname)
		'End If
	
		If Len(Trim(.VoiceOf.Firstname)) > 0 Then
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=TALENT")
		End If

	End With
End Sub

'=====================================================================
'	Weather
'---------------------------------------------------------------------
Sub Insert_Weather()

	'Dim showSponsor
	'showSponsor = False

	'If Plugin.ShowWeatherSponsor Then
	'	Sponsor_Set Settings.SponsorFoxBox
	'	showSponsor = True 
	'End If

	'Ventuz_Update ".Dropdown.Weather.ShowSponsor", showSponsor

	'Ventuz_Update ".Dropdown.Type", 7
	'Ventuz_Update ".Dropdown.Weather.Labels", "TEMPERATURE,WIND"
	'Ventuz_Update ".Dropdown.Weather.Values", Settings.WeatherTemperature & "º F," & Settings.WeatherWind & "%"
	'Ventuz_Update ".Dropdown.Weather.Forecast", Settings.WeatherForecast
	'Ventuz_Invoke ".Dropdown.Insert"
End sub

'=====================================================================
'	Substitution
'---------------------------------------------------------------------
Sub Update_Substitution()
	Dim objTeam 
	Dim objIncoming 
	Dim objOutgoing
	Dim objEvent
	Dim teamName
	Dim subText
	Dim numSubs

	SubstitutionTeamExist = False

	If PluginSettings.SubstitutionTextType = 0 Then
		subText = Ucase(SUBSTITUTION_TEXT)
	Else
		subText = Ucase(SUBSTITUTION_TEXT_SPANISH)
	End If
	
	Set objEvent = Game.GameEvents.SelectedEvent

	If Not objEvent.EventType = 3 then
		Exit Sub
	End If
	
	Select Case objEvent.TeamCode
		Case 0
			TeamName = UCase(Game.Home.Shortabbreviation)
		Case 1
			TeamName = UCase(Game.Visitors.Shortabbreviation)
	End Select
	
	Select Case Plugin.SubCountsSelected

		Case 0
			numSubs = ""
		Case 1
			numSubs = "1st" & Chr(32)
		Case 2
			numSubs = "2nd" & Chr(32)
		Case 3
			numSubs = "3rd" & Chr(32)
		Case 4
			numSubs = "4th" & Chr(32)
		Case 5
			numSubs = "5th" & Chr(32)
	End Select

	Set objIncoming = Game.GameEvents.SelectedEvent.Player2
	Set objOutgoing = Game.GameEvents.SelectedEvent.Player1

	If Not objIncoming Is Nothing Then
		If Not objOutgoing Is Nothing Then

			SubstitutionTeamExist = True

			'subText = TeamName & " " & subText
			 'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBSTITUTION/HEADER_LOAD=" & HEADER_DOUBLEQUOTES & subText & HEADER_DOUBLEQUOTES)
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBSTITUTION/HEADER_LOAD=" & HEADER_DOUBLEQUOTES & numSubs & subText & HEADER_DOUBLEQUOTES)
			
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBSTITUTION/PLAYER_NAME_LOAD=" & HEADER_DOUBLEQUOTES & objOutgoing.Displayname & HEADER_DOUBLEQUOTES )
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBSTITUTION/PLAYER_NAME2_LOAD=" & HEADER_DOUBLEQUOTES & objIncoming.Displayname & HEADER_DOUBLEQUOTES )

			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBSTITUTION/JERSEY1_LOAD=" & HEADER_DOUBLEQUOTES & objOutgoing.DisplayJersey & HEADER_DOUBLEQUOTES )
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBSTITUTION/JERSEY2_LOAD=" & HEADER_DOUBLEQUOTES & objIncoming.DisplayJersey & HEADER_DOUBLEQUOTES )

			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBSTITUTION/TRICODE_DISPLAY_LOAD=" & HEADER_DOUBLEQUOTES & TeamName & HEADER_DOUBLEQUOTES)

			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBSTITUTION/CARD_LOAD=0")

			
		End If
	End If
End Sub


Sub Insert_Substitution
	
	Update_Substitution
	
	If SubstitutionTeamExist Then
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOCUS_LOAD=2")
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=SUBSTITUTION")
	End If

End Sub

Sub GameEvents_SelectedEventChange(objEvent)
	
	Select Case  objEvent.EventType 'Substitution
		Case 3
			If FoxBox.State = 7 Then 'FOXBOX_SUBSTITUTION
				Update_Substitution
			End If
		Case 1, 2 'Cards
			If FoxBox.State = 6 Then 'FOXBOX_CARD
				Update_Card
			End If

		Case 0, 4 'Goal Min
			If FoxBox.State = 5 Then 'FOXBOX_GOALMINUTES
				Update_GoalMinutes
			End If

		Case Else

	End Select
End Sub



Sub Get_SubstitutionCount_Home()
	'Plugin.HomeSubstitutionCount = Get_SubstitutionCount(0)
End Sub

Sub Get_SubstitutionCount_Visitors()
	'Plugin.VisitorsSubstitutionCount = Get_SubstitutionCount(1)
End Sub

Sub Show_SubstitutionCount(visible)

	If visible Then
		Viz_Update "SUBSTITUTIONS", "ON"
	Else
		Viz_Update "SUBSTITUTIONS", "OFF"
	End If
End Sub

'This gets called in SoccerPlugin body
Sub Initialize_SubstitutionCount()
	
	For i = 1 to PluginSettings.MaxSubstitutionCounts
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBLEFT" & i & "=1")
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBRIGHT" & i & "=1")
	Next
	
End Sub

Sub Update_SubstitutionCount_Home()

	Dim i
	Dim subRemaining

	With Plugin
		
		If .HomeSubstitutionCount > PluginSettings.MaxSubstitutionCounts Then
			Exit Sub
		End If

		subRemaining = .HomeSubstitutionCount
		
		If subRemaining = PluginSettings.MaxSubstitutionCounts Then
			For i = 1 to PluginSettings.MaxSubstitutionCounts
				If Settings.ReverseTeamOrder Then
					Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBRIGHT" & i & "=1")
				Else
					Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBLEFT" & i & "=1")
				End If
			Next
		Else
			If subRemaining  > 0 Then  
				For i = 1 to subRemaining  
					If Settings.ReverseTeamOrder Then
						Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBRIGHT" & i & "=1")
					Else
						Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBLEFT" & i & "=1" )
					End If
				Next
				
				For i = subRemaining + 1 to PluginSettings.MaxSubstitutionCounts

					If Settings.ReverseTeamOrder Then
						Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBRIGHT" & i & "=0")
					Else
						Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBLEFT" & i & "=0" )
					End If
				Next
			Else
				For i = 1 to PluginSettings.MaxSubstitutionCounts
					If Settings.ReverseTeamOrder Then
						Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBRIGHT" & i & "=0")
					Else
						Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBLEFT" & i & "=0")
					End If
				Next
	
			End If
		End If

	End With
End Sub

Sub Update_SubstitutionCount_Visitors()

	Dim i
	Dim subRemaining
	
	With Plugin
		If .VisitorsSubstitutionCount > PluginSettings.MaxSubstitutionCounts Then
			Exit Sub
		End If

		subRemaining = .VisitorsSubstitutionCount
		
		If subRemaining = PluginSettings.MaxSubstitutionCounts Then
			For i = 1 to PluginSettings.MaxSubstitutionCounts

				If Settings.ReverseTeamOrder Then
					Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBLEFT" & i & "=1")
				Else	
					Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBRIGHT" & i & "=1")
				End if
			Next
		Else
			If subRemaining  > 0 Then
				For i = 1 to subRemaining 

					If Settings.ReverseTeamOrder Then
						Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBLEFT" & i & "=1")
					Else
						Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBRIGHT" & i & "=1")
					End If

				Next

				For i = subRemaining + 1 to PluginSettings.MaxSubstitutionCounts

					If Settings.ReverseTeamOrder Then
						Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBLEFT" & i & "=0")
					Else
						Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBRIGHT" & i & "=0")
					End If
				Next

			Else
				For i = 1 to PluginSettings.MaxSubstitutionCounts

					If Settings.ReverseTeamOrder Then
						Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBLEFT" & i & "=0")
					Else
						Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBRIGHT" & i & "=0")
					End If
				Next
			End If

		End If

	End With
End Sub

'Sub Update_SubstitutionCount_Visitors()
'
'	Dim i
'	Dim subRemaining
'	
'
'	'Get_SubstitutionCount_Visitors
'	
'
'	With Plugin
'		If .VisitorsSubstitutionCount > 3 Then
'			Exit Sub
'		End If
'
'		subRemaining = .VisitorsSubstitutionCount
'		
'		If subRemaining = 3 Then
'			For i = 1 to 3
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET HOME_SUB" & i & "=1")
'			Next
'		End If
'
'		If subRemaining = 2 Then
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET HOME_SUB1=1")
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET HOME_SUB2=1")
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET HOME_SUB3=0")
'		End If
'
'		If subRemaining = 1 Then
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET HOME_SUB1=1")
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET HOME_SUB2=0")
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET HOME_SUB3=0")
'
'		End If
'
'		If subRemaining = 0 Then
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET HOME_SUB1=0")
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET HOME_SUB2=0")
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET HOME_SUB3=0")
'		End If
'	End With
'End Sub
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'Not In Use.  No longer using Game Events Sub counts
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'Function Get_SubstitutionCount(team) 
'	Dim i 
'    	Dim subCounts
'	Dim subExists
'	
'	subCounts = 0
'	subExists = false
'
'	For i = 1 To Game.GameEvents.Count
'        	
'		With Game.GameEvents(i)
'			If .EventType = 3 Then 'substitution
'				If .TeamCode = team Then '0:Home  1:Vis
'					subCounts = subCounts + 1
'					subExists = True
'				End If
'			End If
'        	End With
'    	Next
'
'	Get_SubstitutionCount = subCounts
'
'End Function

Sub Set_SubstitutionChip()
	If Plugin.SubChipsState Then
		Insert_SubstitutionChip
	Else
		Retract_SubstitutionChip
	End If
End Sub
Sub Insert_SubstitutionChip()
	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBS=ON")
	Viz_Update "SUBS", "ON"
End Sub

Sub Retract_SubstitutionChip()
	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SUBS=OFF")
	Viz_Update "SUBS", "OFF"
End Sub

'=====================================================================
'	SPONSORS	
'---------------------------------------------------------------------
Sub Set_SponsorElement(automatic)

	Dim index
	Dim sponsorName

	If automatic Then
		sponsorName = Plugin.AutomatedSponsorName
	Else
		If Plugin.DropDownSponor Then
			sponsorName = Plugin.ManualSponsorNameDrop
		Else
			sponsorName = Plugin.ManualSponsorNameSide
		End If
	End If

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SPONSOR_IMAGE=" &  PluginSettings.SponsorDirectory & sponsorName)
End Sub

Sub Insert_Sponsor()

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SPONSOR=ON")
	'Viz_Update "SPONSOR", "ON"

End Sub

Sub Retract_Sponsor()
	'Viz_Update "SPONSOR", "OFF"
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SPONSOR=OFF")
End Sub

Sub UpdateSponsorInterval()
	Plugin.SetSponsorIntervals
End Sub

Sub UpdateSponsorDuration()
	Plugin.SetSponsorDruation
End Sub

Sub BlockSponsor(blockState)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SponsorBlock=" & blockState)
End Sub



'=====================================================================
'	SHOOTOUT
'---------------------------------------------------------------------
Dim roundMove  'to identify when to move to next 10
Dim ShootoutCurrentRound

'Not Good way to do this but this is available
Sub Viz_Reset_PenaltyKicks()
	ResetShootout
	Plugin.ResetShootoutGraphic
End Sub


Sub Shootout_Update() 'This gets called from frmMain

	Update_PenaltyKicks 'updating hit/miss
	Update_ShootoutScore 'updating penalty scores
	
	'Shootout is Reset
	If Game.shootoutorder = -1 AND Game.ShootoutRound = 1 Then
		Initialize_ShootoutScore
	End If
End Sub


Sub Update_ShootoutScore()
	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PENALTY_KICKS/T1/TALLY_LOAD=" & Game.Home.ShootoutScore)
	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PENALTY_KICKS/T2/TALLY_LOAD=" & Game.Visitors.ShootoutScore)

	'Viz_Update_OutputDirect "PENALTY_KICKS/T1/TALLY_LOAD", Game.Home.ShootoutScore
	'Viz_Update_OutputDirect "PENALTY_KICKS/T2/TALLY_LOAD", Game.Visitors.ShootoutScore

	Viz_Update "PENALTY_KICKS/T1/TALLY_LOAD", Game.Home.ShootoutScore
	Viz_Update "PENALTY_KICKS/T2/TALLY_LOAD", Game.Visitors.ShootoutScore

End Sub

Sub Initialize_ShootoutScore()
	Dim i	
	
	'this variable is used to move next
	ShootoutCurrentRound = Game.ShootoutRound
End Sub

Sub UpdateShootoutOrder()

	If Game.shootoutorder <> -1 Then
		If Game.ShootoutOrder = 1 Then
			Viz_Update "PENALTY_KICKS/WHO_SHOT_FIRST", "AWAY"
		End If
		
		If Game.ShootoutOrder = 0 Then
			Viz_Update "PENALTY_KICKS/WHO_SHOT_FIRST", "HOME"
		End If
	End If
End Sub


Dim numCircle

Sub Update_PenaltyKicks_WIP()
	Dim i
	Dim rounds

	If Game.ShootoutRound <= 5 Then
		rounds = Game.ShootoutRound
	Else
	
		rounds = Game.ShootoutRound - (5*(Game.ShootoutRound\5))
	End If

	If  Game.ShootoutRound > 5 Then
		'If Game.ShootoutRound Mod 5 = 1 Then 
			If ShootoutCurrentRound <> Game.ShootoutRound Then
				RefreshPenalty	
			End if
		'End If
	End If

	For i = 1 to rounds
		If  Game.ShootoutRound <= 5 Then
			Viz_Update "PENALTY_KICKS/T1/ROUND" & i, ConvertKickResultForDP(Game.Home.ShootoutScores(i))
			Viz_Update "PENALTY_KICKS/T2/ROUND" & i, ConvertKickResultForDP(Game.Visitors.ShootoutScores(i))
		Else
			Viz_Update "PENALTY_KICKS/T1/ROUND" & i, ConvertKickResultForDP(Game.Home.ShootoutScores(i+5*(Game.ShootoutRound\5)))
			Viz_Update "PENALTY_KICKS/T2/ROUND" & i, ConvertKickResultForDP(Game.Visitors.ShootoutScores(i+5*(Game.ShootoutRound\5)))
		End If
	Next

	If ShootoutCurrentRound <> Game.ShootoutRound Then
		ShootoutCurrentRound =  Game.ShootoutRound
	End If

End Sub

Sub Update_PenaltyKicks()
	Dim homeConverted
	Dim awayConverted
	Dim result

	If Game.ShootoutRound <= 5 Then

		Dim roundIndex 
		
		If Game.ShootoutRound = 0 Then
			roundIndex = 1
		Else
			roundIndex = Game.ShootoutRound
		End If

		homeConverted = ConvertKickResultForDP(Game.Home.ShootoutScores(roundIndex))
		awayConverted = ConvertKickResultForDP(Game.Visitors.ShootoutScores(roundIndex))
		
		'Viz_Update_OutputDirect "PENALTY_KICKS/T1/ROUND" & roundIndex, homeConverted
		'Viz_Update_OutputDirect "PENALTY_KICKS/T2/ROUND" & roundIndex, awayConverted
		
		Viz_Update "PENALTY_KICKS/T1/ROUND" & roundIndex, homeConverted
		Viz_Update "PENALTY_KICKS/T2/ROUND" & roundIndex, awayConverted

		numCircle = 5

	Else

		roundIndex = Game.ShootoutRound Mod 5
		'msgbox roundIndex
		
		If Game.ShootoutRound Mod 5 = 0 Then
			roundIndex = 5
		End If

		homeConverted = ConvertKickResultForDP(Game.Home.ShootoutScores(Game.ShootoutRound))
		awayConverted = ConvertKickResultForDP(Game.Visitors.ShootoutScores(Game.ShootoutRound))
		
		'Viz_Update_OutputDirect "PENALTY_KICKS/T1/ROUND" & roundIndex, homeConverted
		'Viz_Update_OutputDirect "PENALTY_KICKS/T2/ROUND" & roundIndex, awayConverted
		
		Viz_Update "PENALTY_KICKS/T1/ROUND" & roundIndex, homeConverted
		Viz_Update "PENALTY_KICKS/T2/ROUND" & roundIndex, awayConverted
		
		If ShootoutCurrentRound <> Game.ShootoutRound Then	

			If Game.ShootoutRound < ShootoutCurrentRound Then
				ShootoutCurrentRound = Game.ShootoutRound 
			End If

			If ShootoutCurrentRound + 1 = Game.ShootoutRound Then
				If numCircle = roundIndex Then

				Else
					RefreshPenalty
				End If
			End If
			
		End If

	End If
	
	If ShootoutCurrentRound <> Game.ShootoutRound Then
		ShootoutCurrentRound =  Game.ShootoutRound
	End If


End Sub

Sub RefreshPenalty() 
	numCircle = Game.ShootoutRound Mod 5

	If Game.ShootoutRound Mod 5 = 0 Then
		numCircle = 5
	End If
	

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET PENALTY_KICKS/REFRESH=TRUE")
End Sub

Function ConvertKickResultForDP(kickResult)

	Select case kickResult
		case -1  'defalut
			converted = 0
		case 0 'miss
			converted = 1
		case 1 'score
			converted = 2
		Case else
			converted = 0
	End Select

	ConvertKickResultForDP = converted

End Function 

Sub Update_CustomStatCategory()
	Dim i

	For i = 1 to 4 

		PluginSettings.CustomStats(i) = Settings.CustomStats(i)
	Next 
	
End Sub

Sub HitRoundOne()
	Viz_Update_OutputDirect "PENALTY_KICKS/T1/ROUND1" , "2"
End Sub

Sub MIssRoundOne()
	Viz_Update_OutputDirect "PENALTY_KICKS/T1/ROUND1" , "1"
End Sub

Sub ResetRoundOne()
	Dim i
	For i = 1 to 5
		Viz_Update_OutputDirect "PENALTY_KICKS/T" & i & "/ROUND1" , "2"
	Next
End Sub

Sub TurnOnClock()
	Viz_Update "PENALTYCLOCK", "ON"
End Sub


Sub TurnOffClock()
	Viz_Update "PENALTYCLOCK", "OFF"
End Sub

Sub ResetShootout()
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET RESET=ON")
End Sub


Function GetGameEvenObjectByID(eventID)
	Dim i
	Dim gameEventObject 'As SoccerObjects.clsGameEvent

	For i = 1 to Soccer.Game.GameEvents.Count
		If  eventID = Soccer.Game.GameEvents(i).ID Then
			Set GetGameEvenObjectByID = Soccer.Game.GameEvents(i) 

			Log.LogEvent "GetGameEvenObjectByID: Object with ID:" & Soccer.Game.GameEvents(i).ID & " Found" , "Debug", 0, "Renderer"
			Exit Function
		End if
	Next

End Function 


Sub UpdateSummary(summaryArray(), title)
	Dim i
	Dim iteration
	Dim row1ID, row2ID
	Dim result
	Dim objGEHelper
	Dim goalText
	Dim minuteText

	Dim row1flag, row2Flag
	
	On Error Resume Next

	If  summaryArray.Count < 1 Then
		Exit Sub
	End If


	iteration =  summaryArray.Count\2 
	
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAMESUMMARY/TEXT1_LOAD=" & title)

	If summaryArray.Count = 1 Then 'this case iteration is 0
		row1ID = summaryArray.Item(1)
		
		With GetGameEvenObjectByID(row1ID)
			goalText =  GetGoalText(row1ID) 
			row1flag = .Team.Abbreviation
			
			If .EventType = 4 Then
				'opponent
				row1flag = GetOpponentAbbreviation(row1flag)
			End If
			'
			minuteText = FormatGoalMinutes(GetGameEvenObjectByID(row1ID))

			Call SetSummaryRowOne(row1flag, .Player1.DisplayName & goalText, minuteText , Plugin.GameEventHeperObjectByEventID(row1ID).EventAdditonalInfo)
		End With

		Call SetSummaryRowTwo("BLANKALPHA", "", "", "")
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GLOW=ON")
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=GAMESUMMARYSINGLE")
	End If
	
	For i = 1 to iteration
		
		row1ID = summaryArray.Item(2*i-1)
		row2ID = summaryArray.Item(2*i)

		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAMESUMMARY/TEXT1_LOAD=" & title)

		'row1		
		With GetGameEvenObjectByID(row1ID)
			goalText =  GetGoalText(row1ID) 
			row1flag = .Team.Abbreviation
			
			If .EventType = 4 Then
				'opponent flag
				row1flag = GetOpponentAbbreviation(row1flag)
			End If

			Call SetSummaryRowOne(row1flag, .Player1.DisplayName & goalText, FormatGoalMinutes(GetGameEvenObjectByID(row1ID)) , Plugin.GameEventHeperObjectByEventID(row1ID).EventAdditonalInfo)
		
		End With
		'row2
		With GetGameEvenObjectByID(row2ID)
			goalText =  GetGoalText(row2ID) 
			row2flag = .Team.Abbreviation
			
			If .EventType = 4 Then
				'opponent flag
				row2flag = GetOpponentAbbreviation(row2flag)
			End If
			'
			Call SetSummaryRowTwo(row2flag, .Player1.DisplayName & goalText, FormatGoalMinutes(GetGameEvenObjectByID(row2ID)), Plugin.GameEventHeperObjectByEventID(row2ID).EventAdditonalInfo)
		End With
		
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GLOW=ON")
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=GAMESUMMARYDOUBLE")
		
		if (2*i+1) <= summaryArray.Count Then

			result = MsgBox("More " & title &  " Move Next?", vbYesNo)
		
			If result = vbYes Then
				
				If (2*i+1) = summaryArray.Count Then
					'Msgbox "index: " & 2*i+1 & "ID: " & row1ID
					
					row1ID = summaryArray.Item(2*i+1)

					'row1		
					With GetGameEvenObjectByID(row1ID)
						
						goalText =  GetGoalText(row1ID) 
						row1flag = .Team.Abbreviation

						If .EventType = 4 Then
							'opponent flag
							row1flag = GetOpponentAbbreviation(row1flag)
						End If


						Call SetSummaryRowOne(row1flag, .Player1.DisplayName & goalText , FormatGoalMinutes(GetGameEvenObjectByID(row1ID)), Plugin.GameEventHeperObjectByEventID(row1ID).EventAdditonalInfo)
					End With

					Call SetSummaryRowTwo("BLANKALPHA", "", "", "")
					
					'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GLOW=ON")
					Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=GAMESUMMARYSINGLE")
				End If
			Else
				Exit Sub
			End If
		End If
	Next
End Sub

Function GetOpponentAbbreviation(teamAbb)

	If Game.Home.Abbreviation = teamAbb Then
		GetOpponentAbbreviation = Game.Visitors.Abbreviation
	Else
		GetOpponentAbbreviation = Game.Home.Abbreviation
	End If
End Function



Sub SetSummaryRowOne(teamName, pName, scoringMin, description)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAMESUMMARY/T1FLAG_LOAD=" & teamName)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAMESUMMARY/TEXT2_LOAD=" & pName)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAMESUMMARY/TEXT2A_LOAD=" & scoringMin)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAMESUMMARY/TEXT2B_LOAD=" & description)
End Sub

Sub SetSummaryRowTwo(teamName, pName, scoringMin, description)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAMESUMMARY/T2FLAG_LOAD=" & teamName)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAMESUMMARY/TEXT3_LOAD=" & pName )
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAMESUMMARY/TEXT3A_LOAD=" & scoringMin)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAMESUMMARY/TEXT3B_LOAD=" & description)
End Sub

Function GetGoalText(eventID)
	Dim objGEHelper

	Set objGEHelper =  Plugin.GameEventHeperObjectByEventID(eventID)
	
	GetGoalText = GoalTextByEventType(objGEHelper.EventType)

	'Select Case objGEHelper.EventType
	'	Case 4 'OG
	'		GetGoalText = " (OG)"
	'	Case 5 'Penalty Kick
	'		GetGoalText = " (P)"
	'	Case else
	'		GetGoalText = ""

	'End Select

End Function

Function GoalTextByEventType(eventType)
	
	Select Case eventType
		Case 4 'OG
			GoalTextByEventType = " (OG)"
		Case 5 'Penalty Kick
			GoalTextByEventType = " (P)"
		Case else
			GoalTextByEventType = ""

	End Select

End Function


Function GetGoalMinutesText(MinuteText)

	If InStr(MinuteText, "+") > 0 Then
		Dim tmp
		tmp = Split(MinuteText, "+")
		GetGoalMinutesText = FormatgoalMinutes(Cint(tmp(0))) 
	Else
		GetGoalMinutesText = MinuteText & chr(39)
	End If
End Function

'=====================================================================
'	NATIONAL ANTHEM
'---------------------------------------------------------------------
Sub Insert_NationalAnthem()

	UpdateNationalAnThemFlag

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET L3=ON")
End Sub

Sub UpdateNationalAnThemFlag()
	Dim teamAbb
	
	If PluginSettings.NationalAnthemSelectedTeam = 0 Then
		teamAbb = Game.Visitors.Abbreviation
	Else
		teamAbb = Game.Home.Abbreviation
	End If

	'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET L3_LOAD=" & teamAbb)	
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET L3_LIVE=" & teamAbb)	
End Sub

Sub Retract_NationalAnthem()
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET L3=OFF")
End Sub

'=====================================================================
'	Out of Town Score	
'---------------------------------------------------------------------
Sub OTS_Update()
	Dim objNote
	Dim sBody
	Dim tmp
	Dim team1Name, team2Name
	Dim team1Score, team2Score
	Dim matchStatus

	Log.LogEvent "OTS_Update is called", "Debug", 0, "Renderer"

	Set objNote = Soccer.Game.Notes.GetNoteByTitle(Plugin.OTSNoteSelected)
	
	If Not objNote Is Nothing Then
	
		Log.LogEvent "Note " & Plugin.OTSNoteSelected & " FOUND ",  "Debug", 0, "Renderer"
		' Update Note
		sNote = objNote.Body

		tmp = Split(sNote, ":")

		'msgbox "Ubound(tmp): " & Ubound(tmp)
		
		If Ubound(tmp) <= 5   Then
			team1Name = tmp(1)
			team2Name = tmp(2)
			team1Score = tmp(3)
			team2Score = tmp(4)
			matchStatus = ""

			'Viz_Update "GAME_SCORE/TEAM1_LOAD", team1Name
			'Viz_Update "GAME_SCORE/TEAM2_LOAD", team2Name
			'Viz_Update "GAME_SCORE/SCORE1_LOAD", team1Score
			'Viz_Update "GAME_SCORE/SCORE2_LOAD", team2Score
			
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAME_SCORE/TEAM1_LOAD=" & team1Name)
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAME_SCORE/TEAM2_LOAD=" & team2Name)
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAME_SCORE/SCORE1_LOAD=" & team1Score)
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAME_SCORE/SCORE2_LOAD=" & team2Score)

			If  Ubound(tmp) = 5 Then
				'line text exist
				matchStatus = tmp(5)
			End If
			
			'Viz_Update "GAME_SCORE/TEXT1_LOAD", matchStatus  
			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAME_SCORE/TEXT1_LOAD=" & matchStatus )

		Else
			Log.LogEvent "OTS Note " & sNote & " INCORRECTLY FORMATTED", "Debug", 0, "Renderer"
		End If

		'msgbox team1Name & " " & team2name

	Else
		Log.LogEvent "OTS Note " & Plugin.OTSNoteSelected & " NOT FOUND", "Debug", 0, "Renderer"
	End If
End Sub

Sub Insert_OTS()
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET DROPDOWN_DISPATCH_LOAD=GAME_SCORE")
End Sub

'GAME_SCORE/TEAM1_LOAD  (Left team)
'GAME_SCORE/TEAM2_LOAD (Right team)
'GAME_SCORE/SCORE1_LOAD (left team score)
'GAME_SCORE/SCORE2_LOAD (right team score)
'DROPDOWN_DISPATCH_LOAD=GAME_SCORE

'=====================================================================
'	WRAPPER Misc FUNCTIONS
'---------------------------------------------------------------------
Sub RefreshGraphicPluginUIElements()

	If Game.Home Is Nothing Then
		Plugin.HomeShortAbbreviation = "HOME"
	Else
		Plugin.HomeShortAbbreviation = UCase(Game.Home.ShortAbbreviation)
	End If
	

	If Game.Visitors Is Nothing Then
		Plugin.VisitorsShortAbbreviation = "VISITORS"
	Else
		Plugin.VisitorsShortAbbreviation = UCase(Game.Visitors.ShortAbbreviation)
	End If

	Plugin.RefreshUIElements
End Sub

Sub TurnOffBuildInterface()
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET CONTROLS=0")
End Sub


'=====================================================================
'	WRAPPERS
'---------------------------------------------------------------------




'Function ReplaceHighlights(p_sText)
'	Dim sText
'	
'	sText = p_sText
'
'	sText = Replace(sText, "<", "{")
'	sText = Replace(sText, ">", "}")
'	sText = Replace(sText, "{", "<HIGHLIGHT>")
'	sText = Replace(sText, "}", "</HIGHLIGHT>")
'
'	' Return text
'	ReplaceHighlights = sText 
'End Function
'
'Function ReplaceCrLf(p_sText)
'	Dim sText
'	
'	sText = p_sText
'
'	sText = Replace(sText, vbCrLf, "&#13;&#10;")
'	sText = Replace(sText, "<P>", "&#13;&#10;")
'	sText = Replace(sText, "<CR>", "&#13;&#10;")
'	sText = Replace(sText, "<BR>", "&#13;&#10;")
'
'	' Return text
'	ReplaceCrLf = sText
'
'End Function
'
'Function ReplaceOrdinal(p_sText)
'
'	sText = p_sText
'	
'	sText = Replace(sText, "2nd", "2" & "<sm>ND</sm>")
'	sText = Replace(sText, "1st", "1" & "<sm>ST</sm>")
'
'	ReplaceOrdinal = sText
'End Function


