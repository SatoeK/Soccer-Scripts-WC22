'=====================================================================
' FSN 2008
' FOX Sports VBScript Lowerthird Scoreboard module
'---------------------------------------------------------------------
' Etienne Laneville (Etienne.Laneville@FOX.com)
' November 2002
'=====================================================================

Dim m_objFileSystem

'=====================================================================
'   METHODS
'---------------------------------------------------------------------

' Scoreboard Goto Regular:
'       Insert Lowerthird Scoreboard
Sub Scoreboard_Goto_Regular()
	
	' Update scoreboard
	Scoreboard_Update 

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET INSERT_DISPATCH_LOAD=L3_SCOREBOARD") 
	
	Log.LogEvent "Scoreboard: go to Regular", "Debug", 0, "Renderer"
End Sub

' Scoreboard Goto Rollout:
'       Insert Rollout Scoreboard
Sub Scoreboard_Goto_Rollout()
	
	SlabScoreboard_Update

	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET INSERT_DISPATCH_LOAD=SLAB_SCOREBOARD") 

	Log.LogEvent "Scoreboard: go to Rollout", "Debug", 0, "Renderer"

End Sub

Sub SlabScoreboard_Update()
	Dim fShowSponsor
	
	Viz_Update "SLAB_SCOREBOARD/GAME_STATUS_LOAD", Game.PeriodEndText(3) 

	' Set sponsor 
	'What is the Sponsor Element for Slab?
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SLAB_SCOREBOARD/SPONSOR_IMAGE=" & PluginSettings.SponsorDirectory & PluginSettings.ScoreboardSponsor) 

	fShowSponsor = Show_ScoreboardSponsor
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SLAB_SCOREBOARD/SPONSOR_ACTIVE=" & fShowSponsor) 


'SLAB_SCOREBOARD/AWAY_RECORD_LOAD=AWAYRECORD;
'SLAB_SCOREBOARD/HOME_RECORD_LOAD=HOMERECORD;
	
End Sub



Sub Scoreboard_Goto_Chip()
	' Update scoreboard
	Scoreboard_Update
	Ventuz_Update ".Scoreboards.Chip.Text", PluginSettings.ScoreboardHeader 
	Ventuz_Invoke ".Scoreboards.Chip.Insert"
End Sub

' Scoreboard Update:
'       Update scoreboard
Sub Scoreboard_Update()
	Dim objSponsor
	Dim objCategory
	Dim objVersion

	'Dim fShowSponsor
	Dim fShowClock
	
	'fShowClock = Not Game.Final

	' Initialize
	fShowSponsor = False
	
	'Footer 
	Viz_Update "L3_SCOREBOARD/FOOTNOTE_RIGHT_LOAD", ""
	'Game Status
	Viz_Update "L3_SCOREBOARD/GAME_STATUS_LOAD", Game.PeriodEndText(3) 

	' Set sponsor 
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET L3_SCOREBOARD/SPONSOR_IMAGE=" & PluginSettings.SponsorDirectory & PluginSettings.ScoreboardSponsor) 

	' Sponsor switch
	'If PluginSettings.SponsorShowOnLowerThirds Then
	'	fShowSponsor = 1
	'Else
	'	fShowSponsor = 0
	'End If

	fShowSponsor =  Show_ScoreboardSponsor
	
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET L3_SCOREBOARD/SPONSOR_ACTIVE=" & fShowSponsor) 
	
End Sub

Function Show_ScoreboardSponsor()
	' Sponsor switch
	If PluginSettings.SponsorShowOnLowerThirds Then
		Show_ScoreboardSponsor = 1
	Else
		Show_ScoreboardSponsor = 0
	End If
End Function

' Scoreboard Goto Out:
'       Retract the Lower Third Scoreboard
Sub Scoreboard_Goto_Out()
	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET INSERT_DISPATCH_LOAD=L3_OFF") 
End Sub


