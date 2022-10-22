'=====================================================================
' FSN Soccer 2015
' FOX Sports VBScript Scripting Module
'---------------------------------------------------------------------
' FSN Game / Clock Updates.
'=====================================================================

Option Explicit

'=====================================================================
'   PROCEDURES
'---------------------------------------------------------------------

'	Game Update
'---------------------------------------------------------------------

' Game_Update:
'       Update the Game
Sub Game_Update()

	If Not SceneIsLoaded Then
		Exit Sub
	End If

	On Error Resume Next

	' League Options
	With League
		
	End With
	
	' Update Game data
	With Game

		' Clock and Period
		Viz_Update_OutputDirect "GAMECLOCK", Game.Clock.Value


		'Ventuz_Update ".Game.Period.Period", Game.Period
		'Ventuz_Update ".Game.Period.PeriodText1", Game.PeriodText(3)
		'Ventuz_Update ".Game.Period.PeriodText2", Game.PeriodText(4)
		'Ventuz_Update ".Game.Period.PeriodEndText", ReplaceOrdinal(Game.PeriodEndText(3))
		'Ventuz_Update ".Game.Period.PeriodStart", Game.PeriodStart
		'Ventuz_Update ".Game.Period.PeriodEnd", Game.PeriodEnd
		'Ventuz_Update ".Game.Halftime", Game.Halftime
		'Ventuz_Update ".Game.Final", Game.Final

		
	End With

	' Update Home data
	With Game.Home


		Viz_Update_OutputDirect "TEAM1_LIVE", Ucase(.DisplayAbbreviation)
		Viz_Update_OutputDirect "AwayLogo", Ucase(.Abbreviation)
		Viz_Update_OutputDirect "SCORE1_LOAD", .Score

		'If .DisplayColor = "DEFAULT" Then
		'	Viz_Update_OutputDirect "AwayColor", PluginSettings.TeamColorDirectory & Ucase(.Abbreviation)
		'Else
			Viz_Update_OutputDirect "Team1Color", Ucase(.DisplayColor)
		'End If


'		Viz_Update "HOME_SCORE_S", .Score
'		Viz_Update "HOME_ABB", Ucase(.DisplayAbbreviation)
'		Viz_Update "HOME_TEAM", PluginSettings.TeamColorDirectory & Ucase(.Abbreviation)
'
		
'		
'		If .DisplayColorSecondary = "DEFAULT" Then
'			Viz_Update "HOME_COLOR_2", PluginSettings.VizTeamColorDirectory & Ucase(.Abbreviation)
'		Else
'			Viz_Update "HOME_COLOR_2", PluginSettings.TeamColorSecondaryDirectory & Ucase(.DisplayColorSecondary)
'		End If
'
'		' If Shootout Final winning teams gets +1 goal:
'		If Game.Shootout And Game.Final Then
'			If Game.Visitors.ShootoutScore > Game.Home.ShootoutScore Then
'				'Ventuz_Update ".Home.Score", .Score
'			Else
'				'Ventuz_Update ".Home.Score", .Score + 1
'			End If
'		Else
'			'Ventuz_Update ".Home.Score", .Score
'		End If
'
'		' Aggregate score
'		Viz_Update "HOME_SCORE_AGGREGATE", .AggregateScore
'
'
'		If League.Abbreviation = "BUNDESLIGA" Then
'			If (Ucase(Game.Home.DisplayAbbreviation) <> BundHomeTeam) Then 
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET HOME_ABB=" & UCase(Game.Home.DisplayAbbreviation))
'				Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOXBOX=INITIALIZE")
'				BundHomeTeam = Ucase(Game.Home.DisplayAbbreviation)
'			End If
'		End If
'

	
		
	End With

	' Update Visitors data
	With Game.Visitors

		Viz_Update_OutputDirect "TEAM2_LIVE", Ucase(.DisplayAbbreviation)
		Viz_Update_OutputDirect "HomeLogo", Ucase(.Abbreviation)
		Viz_Update_OutputDirect "SCORE2_LOAD", .Score
		
		'If .DisplayColor = "DEFAULT" Then
		'	Viz_Update_OutputDirect "HomeColor", PluginSettings.TeamColorDirectory & Ucase(.Abbreviation)
		'Else
			Viz_Update_OutputDirect "Team2Color", Ucase(.DisplayColor)
		'End If


		'Viz_Update "AWAY_SCORE_S", .Score
		'Viz_Update "AWAY_ABB", Ucase(.DisplayAbbreviation)
		'Viz_Update "AWAY_TEAM", PluginSettings.TeamColorDirectory & Ucase(.Abbreviation)
		'
		
		'
		'If .DisplayColorSecondary = "DEFAULT" Then
		'	Viz_Update "AWAY_COLOR_2", PluginSettings.TeamColorDirectory & Ucase(.Abbreviation)
		'Else
		'	Viz_Update "AWAY_COLOR_2", PluginSettings.TeamColorSecondaryDirectory & Ucase(.DisplayColorSecondary)
		'End If

		'' If Shootout Final winning teams gets +1 goal:
		'If Game.Shootout And Game.Final Then
		'	If Game.Home.ShootoutScore > Game.Visitors.ShootoutScore Then
		'		'Ventuz_Update ".Visitors.Score", .Score
		'	Else
		'		'Ventuz_Update ".Visitors.Score", .Score + 1
		'	End If
		'Else
		'	'Ventuz_Update ".Visitors.Score", .Score
		'End If
		'
		'' Aggregate score
		''Ventuz_Update ".Visitors.AggregateScore", .AggregateScore
		'Viz_Update "AWAY_SCORE_AGGREGATE", .AggregateScore

		'If League.Abbreviation = "BUNDESLIGA" Then
		'	If (Ucase(Game.Visitors.DisplayAbbreviation) <> BundVisitorsTeam) Then
		'		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET AWAY_ABB=" & UCase(Game.Visitors.DisplayAbbreviation))
		'		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOXBOX=INITIALIZE")
		'		BundVisitorsTeam = Ucase(Game.Visitors.DisplayAbbreviation)

		'	End If
		'End If

		

	End With
	
	If Game.Shootout Then

		Plugin.EnableShootoutUI
		TurnOffClock
		
		If CurrentState <> "Shootout" Then
			If FoxBox.State <> 1 Then 'Retracted
				Plugin.GoTo_FoxBox_State
			End If
		End If

		Shootout_Update

		'do not need special handlings for final
	Else
		Plugin.DisableShootoutUI
		TurnOnClock

		If Game.Final Then
			Plugin.GoTo_FoxBox_State

			Viz_Update_OutputDirect "FINAL", "ON"
		Else
			Viz_Update_OutputDirect "FINAL", "OFF"
		End If

	End If

'	If Game.Final Then
'		Plugin.GoTo_FoxBox_State
'
'		Viz_Update_OutputDirect "FINAL", "ON"
'	Else
'		Viz_Update_OutputDirect "FINAL", "OFF"
'	End If


	GameServer_GameUpdate
	

	UpdateHomeSelected
	UpdateVisitorsSelected

	' Push to DataHive
	'DataHive_Update	DataHive_GetGameString
	
End Sub

Sub UpdateHomeSelected()
		If Game.Home Is Nothing Then
		HomeSelected = "NoTeam"
	Else
		If HomeSelected <> Game.Home.Abbreviation Then
			'msgbox "home team changed"
			Get_VizHomeHeadshots
			HomeSelected = Game.Home.Abbreviation
		End If
	End If
End Sub

Sub UpdateVisitorsSelected()
	If Game.Visitors Is Nothing Then
		HomeSelected = "NoTeam"
	Else
		If VisitorsSelcted <> Game.Visitors.Abbreviation Then
			'msgbox "visitor team changed"
			Get_VizVisitorsHeadshots
			VisitorsSelcted = Game.Visitors.Abbreviation
		End If
	End if
End Sub

Sub Game_Reset()
	Plugin.GameReset
End Sub


'=====================================================================
'---------------------------------------------------------------------
Sub Set_FontColor()

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TEAM1_TEXTCOLOR=" & PluginSettings.HomeFontColor)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TEAM2_TEXTCOLOR=" & PluginSettings.VisitorsFontColor)

	'10 Players Font color
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TENPLAYER/T1COLOR_LOAD =" & PluginSettings.HomeFontColor)
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET TENPLAYER/T2COLOR_LOAD =" & PluginSettings.VisitorsFontColor)

	'Viz_Update "TEAM1_TEXTCOLOR", PluginSettings.HomeFontColor
	'Viz_Update "TEAM2_TEXTCOLOR", PluginSettings.VisitorsFontColor
End Sub

' Game Goal:
'       Display Goal
' This gets invoked from frmMain
Sub Game_Goal(p_objTeam)

	'Setting which team goaled

	If p_objTeam.Abbreviation = Game.Home.Abbreviation Then
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GOAL_FOCUS=0")
	End If

	If  p_objTeam.Abbreviation = Game.Visitors.Abbreviation Then
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GOAL_FOCUS=1")
	End If
End Sub

'	Clock Update
'---------------------------------------------------------------------

' Clock_Update:
'       Update the clock
Sub Clock_Update()

	GameServer_GameUpdate

	' Push to DataHive
	'DataHive_Update	DataHive_GetGameString

End Sub

'	Stoppage Clock Update
'---------------------------------------------------------------------

Sub StoppageTimeClock_Update()

End Sub




Sub Initialize_Game()
	'this gets called at Initialize_SceneData in Viz.vbs

	On Error Resume Next

	' Update Game data
	With Game
	
	End With

	' Update Home data
	With Game.Home
		
	End With

	' Update Visitors data
	With Game.Visitors
		

	End With
	
	
End Sub

