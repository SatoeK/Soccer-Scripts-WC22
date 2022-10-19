'=====================================================================
' FOX NE&O
' Soccer FoxBox Main module
' FOX Sports VBScript Main Scripting Module
'---------------------------------------------------------------------
' Etienne Laneville (Etienne.Laneville@FOX.com)
' July 2015
'---------------------------------------------------------------------
' Main procedures called by Soccer FoxBox application
'=====================================================================

Option Explicit

'=====================================================================
'   CONSTANTS
'---------------------------------------------------------------------

' All constants should be placed in Constants.vbs module.

'=====================================================================
'   DECLARATIONS
'---------------------------------------------------------------------
Dim BundHomeTeam
Dim BundVisitorsTeam

Dim HEADER_DOUBLEQUOTES 
Dim SceneIsLoaded 

Dim HomeSelected
Dim VisitorsSelcted

' None

'=====================================================================
'   PROCEDURES
'---------------------------------------------------------------------

' Application Start:
'       Runs when application is loaded
Sub Application_Startup()
	Dim objScript
	
	SceneIsLoaded = False

	' Import Data
	Splash.Status "Importing Data..."
	Database_Set
	Database_Import

	' Initialize Game
	Splash.Status "Initializing Game..."
	GameInitialize

	' Custom scripts for Add on 
    	CustomScriptsInitialize

    	' Load Sponsors
    	Splash.Status "Importing Sponsors..."
    	Sponsors_Load

	League_Refresh

	Identities_CreateIdentities

	Settings.ReverseTeamOrder = False
	
	'Initialize_HeaderText

   	' Finish launching application
    	Splash.Status "Application starting..."
    	Log.LogEvent "Application Start event: Scripting works fine.", "Debug", 0, "Init"
	
	HEADER_DOUBLEQUOTES = Chr(34) & Chr(34)
	PluginSettings.SponsorShowOnLowerThirds = false
	FootnoteInsertionState = False
	
	'making sure all engines including preview has secne get loaded
	Plugin.SendToRenderer = True
	Plugin.SendToPreview = True

	UpdateHomeSelected
	UpdateVisitorsSelected

	Plugin.ReverseTeamOrder = Settings.ReverseTeamOrder

End Sub

' Application Shutdown:
'       Runs when application is closed
Sub Application_Shutdown()
    'On Error Resume Next
    
    Log.LogEvent "Shutting down...", "Debug", 0, "Init"

    '
    'Plugin.CloseListnerSocket
	' Shut Ventuz Presenter down
	'Ventuz_KillPresenter

	PreviewSocket.Disconnect
	RendererSecondarySocket.Disconnect

	Plugin.DisposeGameEventsArray

    Log.LogEvent "Shut down.", "Debug", 0, "Init"
	
End Sub

' Game Initialize:
'       Initialize Game object
Sub GameInitialize()
	On Error Resume Next

    	Game.Load
	Game.Restore

End Sub

' League Refresh:
'		This should get overridden by the League script's version
Sub League_Refresh
    Log.LogEvent "League scripts not found.", "Debug", 0, "Init"
End Sub

'=====================================================================
'   SHORTCUTS
'---------------------------------------------------------------------

' Delay:
'       Delays following statements for specified time
'       by yielding to system events
Sub Delay(p_sglSeconds)
    Interface.Delay p_sglSeconds
End Sub

'       Delays following statements for specified number of frames
'       by yielding to system events
Public Sub DelayFrames(p_iFrames)
    Delay CSng(p_iFrames / 30)
End Sub

'=====================================================================
'   OPTIONS
'---------------------------------------------------------------------

Function Option_Debug()

	Select Case UCase(CommandLine.Parameter("Debug"))
		
		Case "TRUE", "YES"
			Option_Debug = True
			
		Case Else
			Option_Debug = False

	End Select
	
End Function

Function Option_Database()

	Select Case UCase(CommandLine.Parameter("Database"))
		
		Case "" ' No database specified
			Option_Database = League.Abbreviation & ".mdb"
			
		Case Else
			Option_Database = CommandLine.Parameter("Database")

	End Select
	
End Function

Function Option_Style()

	Select Case UCase(CommandLine.Parameter("Style"))
		
		Case "SAVED"
			Option_Style = Settings.Style
			
		Case Else
			Option_Style = CommandLine.Parameter("Style")

	End Select
	
End Function

Function Option_League()

	If UCase(CommandLine.Parameter("League")) = "" Then
		Option_League = "MLS"
		League.Abbreviation = Option_League
		
	Else
		Option_League = UCase(CommandLine.Parameter("League"))
		League.Abbreviation = Option_League	
	End If
	

	Plugin.LeagueAbbreviation = Option_League

End Function

'Sub Set_BundesLigaTeams()
'	BundHomeTeam = UCase(Game.Home.DisplayAbbreviation)
'	BundVisitorsTeam = UCase(Game.Visitors.DisplayAbbreviation)
'End Sub








