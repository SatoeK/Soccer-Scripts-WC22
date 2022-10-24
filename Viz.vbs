OptiON Explicit

Dim VizLayer
Dim SceneExistInEngine


Dim m_objVizDictionary
Dim m_objPreviewDictionary

'Const COMMAND_HEADER = "foxbox_send "
Const COMMAND_HEADER = "-1 "


'=====================================================================
'---------------------------------------------------------------------
Sub Viz_Send(p_sCommand)

	On Error Resume Next

	If Plugin.SendToRenderer Then
		Log.LogEvent  p_sCommand, "Viz", 0, "Renderer"
		
		If InStr(1, p_sCommand, "™", 0) > 0 Then
			Interface.Viz_FoxBoxRendererSocket_Send(p_sCommand)
		Else
			RendererSocket.Send ConvertToUTF8(p_sCommand) & Chr(0)
		End If
		
		Call VizSecondary_Send(p_sCommand)

	End If

	If Plugin.SendToPreview Then
		Preview_Send(p_sCommand)
	End If
	
End Sub

Sub TEST2()
	Viz_SendTest(VizLayer & "*FUNCTION*DataPool*Data SET POPUP_NOTELINE_LIVE=" & HEADER_DOUBLEQUOTES & PluginSettings.FoxBoxHeader & HEADER_DOUBLEQUOTES)
End Sub

Sub Viz_SendTest(p_sCommand)

	On Error Resume Next

	Dim  stxt

	If Plugin.SendToRenderer Then

		'msgbox InStr(1, stxt, "™", 0)
		'If InStr(1, p_sCommand, "™", 0) > 0 Then
		'	Interface.TEST(p_sCommand)
		'End If

		RendererSocket.Send Interface.ConvertToUTF8(p_sCommand) & Chr(0)
		
		Log.LogEvent  p_sCommand, "Viz", 0, "Renderer"
		'Call VizSecondary_Send(p_sCommand)

	End If

	'If Plugin.SendToPreview Then
	'	Preview_Send(p_sCommand)
	'End If
	
End Sub


'=====================================================================
'---------------------------------------------------------------------
Sub Viz_Update(p_sPropertyName, p_sValue)
	Dim sValue
	
	If Plugin.SendToRenderer Then
		'Check Library
		Call Update_Output(p_sPropertyName, p_sValue)
	End If

	If Plugin.SendToPreview Then
		'Check Library
		Call  Update_Preview (p_sPropertyName, p_sValue)
	End If		

End Sub

Sub Update_Output(p_sPropertyName, p_sValue)
	Dim sValue

	'Check Library
	If Viz_UpdateDictionary(p_sPropertyName, p_sValue) Then

		sValue = p_sValue

		sValue =  VizLayer & "*FUNCTION*DataPool*Data SET " & p_sPropertyName & "=" & sValue
		
		Log.LogEvent "Updating '" & p_sPropertyName & "' to '" & p_sValue & "'", "Debug", 0, "Viz"
		
		Viz_Send sValue
	End If
End Sub

Sub VizSecondary_Send(p_sCommand)
	
	If  RendererSecondarySocket.Connected Then
		Log.LogEvent  p_sCommand, "Viz", 0, "VizSecondary_Send"
		'
		If InStr(1, p_sCommand, "™", 0) > 0 Then
			Plugin.Viz_FoxBoxRendererSecondarySocket_Send(p_sCommand)
		Else
			RendererSecondarySocket.Send ConvertToUTF8(p_sCommand) & Chr(0)
		End If
	End If

End Sub

'Straight to Output
'=====================================================================
'---------------------------------------------------------------------
'This gets sent to output always
Sub Viz_Send_OutputDirect(p_sCommand)
	Log.LogEvent  p_sCommand, "Viz", 0, "Renderer"

	If InStr(1, p_sCommand, "™", 0) > 0 Then
		Interface.Viz_FoxBoxRendererSocket_Send(p_sCommand)
	Else
		RendererSocket.Send ConvertToUTF8(p_sCommand) & Chr(0)
	End If
	
	'RendererSocket.Send ConvertToUTF8(p_sCommand) & Chr(0)


	Call VizSecondary_Send(p_sCommand)
End Sub

'=====================================================================
'---------------------------------------------------------------------
Sub Viz_Update_OutputDirect(p_sPropertyName, p_sValue)
	Dim sValue
	
	If Viz_UpdateDictionary(p_sPropertyName, p_sValue) Then
		sValue = p_sValue

		sValue =  VizLayer & "*FUNCTION*DataPool*Data SET " & p_sPropertyName & "=" & sValue
		
		Log.LogEvent "Updating '" & p_sPropertyName & "' to '" & p_sValue & "'", "Debug", 0, "Viz"

		'Output Direct
		Viz_Send_OutputDirect (sValue)
	End If

	If Plugin.SendToPreview Then
		Call Update_Preview (p_sPropertyName, p_sValue)
	End If
End Sub


'=====================================================================
'PREVEIW 
'---------------------------------------------------------------------
'it gets sent even the value is the ame as before1
Sub Preview_Send(p_sCommand)

	If PreviewSocket.Connected Then
		Log.LogEvent  p_sCommand, "Viz", 0, "Preview_Send"

		If InStr(1, p_sCommand, "™", 0) > 0 Then
			Plugin.Viz_FoxBoxPreviewSocket_Send(p_sCommand)
		Else
			PreviewSocket.Send  ConvertToUTF8(p_sCommand) & Chr(0)
		End If
	End If
End Sub

'value gets sent if it is different
Sub Update_Preview (p_sPropertyName, p_sValue)
	Dim sValue
	
	'Check Library
	If Preview_UpdateDictionary(p_sPropertyName, p_sValue) Then
		sValue = p_sValue

		sValue =  VizLayer & "*FUNCTION*DataPool*Data SET " & p_sPropertyName & "=" & sValue
		
		Log.LogEvent "Updating '" & p_sPropertyName & "' to '" & p_sValue & "'", "Debug", 0, "Viz"
		Preview_Send (sValue)
	End If
End Sub


''=====================================================================
'---------------------------------------------------------------------
Sub Get_VizSceneNames() 

	Dim cmd

	cmd = "R_sceneNames SCENE*" & PluginSettings.SceneDirectory & " GET"
	Viz_Send(cmd)

	'ex: send 1 SCENE*ONLINE/_NFL_ON_FOX_ GET
End Sub

'=====================================================================
'---------------------------------------------------------------------
Sub ProcessResponseFromTheEngine(p_sData)

    Dim retArray
    Dim retHeader

    
    retArray = Split(p_sData, Chr(32))
    
    retHeader = Trim(retArray(0))

    Select Case retHeader

        Case "R_sceneNames"
            If Trim(retArray(1)) = "ERROR" Then
                MsgBox "Error in getting Scene names from the Engine: " & vbCrLf & p_sData
                'LogEvent "Error in getting Scene names from the Engine: " & vbCrLf & p_sData
                Exit Sub
            End If
                  
            
            If UBound(retArray) = 1 And retArray(1) = Chr(0) Then
                MsgBox "No scenes exist in the folder: " & PluginSettings.SceneDirectory
                'LogEvent "No scenes exist in the folder: " & Settings.VizSceneDirectory
                Exit Sub
            End If

	    'PluginSettings.SelectedSceneName = "" 'for debug delete later
	    ProcessAvailableScenes(p_sData) 


	    If Len(PluginSettings.SelectedSceneName) > 0 Then
		'the scene has been selected previously
		If SceneExistInEngine Then
			 Log.LogEvent PluginSettings.SelectedSceneName & " exists in the returned list.  Loading next @ProcessResponseFromTheEngine", "Debug", 0, "Viz"	
			 Viz_UnLoadScene
			 Viz_LoadScene
		Else
			Log.LogEvent PluginSettings.SelectedSceneName & " does NOT exists in the returned list.  Invoking Plugin.DisplaySceneSelection next @ProcessResponseFromTheEngine", "Debug", 0, "Viz"	
			Plugin.DisplaySceneSelection
		End If
	    Else
		'never been selcted
		Log.LogEvent "Scene has never been selected before. Invoking Plugin.DisplaySceneSelection next @ProcessResponseFromTheEngine", "Debug", 0, "Viz"	
		Plugin.DisplaySceneSelection
	    End If

            
        Case "R_sponsorNames"

            If Trim(retArray(1)) = "ERROR" Then
                MsgBox "Error in getting Sponsor names from the Engine: " & vbCrLf & p_sData
                'LogEvent "Error in getting Sponsor names from the Engine: " & vbCrLf & p_sData
                Exit Sub
            End If
            
            If UBound(retArray) = 1 And retArray(1) = Chr(0) Then
                MsgBox "No sponsor images exist in the folder: " & Settings.VizSponsorDirectory
                'LogEvent "No sponsor images exist in the folder: " & Settings.VizSponsorDirectory
                Exit Sub
            End If
            
             Call ProcessAvailableSponsors(p_sData)

	Case "R_talentNames"
		Call ProcessAvailableTalents(p_sData)

     	Case "R_HomeHeadShots"
		Call ProcessHeadshots(p_sData, 1)
	
	Case "R_VisitorsHeadShots"
		Call ProcessHeadshots(p_sData, 0)

     	Case Else
            'Do nothing
	    'Log.LogEvent "It does not meet with return header @ProcessResponseFromTheEngine.  returned: " & p_sData , "Debug", 0, "Viz"	
    End Select
    

End Sub

'=====================================================================
'---------------------------------------------------------------------
Sub ProcessAvailableScenes(byVal sData)
	Dim tmpArray
	Dim sceneArray
	Dim i
	Dim lastIndex
	Dim sceneEntry
	Dim tmp

	Log.LogEvent "ProcessAvailableScenes is called", "Debug", 0, "Viz"

	sData = Trim(sData)

	tmpArray = Split(sData, chr(32))

	If ubound(tmpArray) < 0 Then
		Exit Sub
	End If

	ReDim sceneArray(ubound(tmpArray)-1)
	
	'skipping the header
	for i = 1 to ubound(tmpArray)-1
		sceneEntry = tmpArray(i) 'ex: IMAGE*SOCCER/IMAGES/SPONSOR/FOXBOX/VISA 
		tmp = Split(sceneEntry, chr(47)) ' split with "/"
		lastIndex = ubound(tmp) 'last index contains scene name
		sceneArray(i) = tmp(lastIndex)

		If Len(PluginSettings.SelectedSceneName) > 0 Then
			If StrComp(PluginSettings.SelectedSceneName,  tmp(lastIndex)) = 0 Then
				SceneExistInEngine = True
			End If
	 	End If
	next

	Plugin.Set_SceneList sceneArray

End Sub

'=====================================================================
'---------------------------------------------------------------------
Sub Get_VizSponsorNames()
	Dim cmd

	'msgbox "GetVizSponsorName is called"

	'sponsorDirectory = PluginSettings.SponsorDirectory 

	cmd = "R_sponsorNames IMAGE*" & PluginSettings.SponsorDirectory & " GET"

	Viz_Send(cmd)

End Sub

'=====================================================================
'---------------------------------------------------------------------
Sub Get_VizTalentNames()
	Dim cmd

	'msgbox "GetVizSponsorName is called"

	'sponsorDirectory = PluginSettings.SponsorDirectory 

	cmd = "R_talentNames IMAGE*" & TALENT_HEADSHOT_DIRECTORY_GH & " GET"

	Viz_Send(cmd)

End Sub

Sub Get_VizHomeHeadshots()

	Dim cmd
	
	'cmd = "R_HomeHeadShots IMAGE*" & HEADSHOT_DIRECTORY_GH & Game.Home.Abbreviation & "-WOMEN" & " GET"
	cmd = "R_HomeHeadShots IMAGE*" & HEADSHOT_DIRECTORY_GH & Game.Home.Abbreviation & " GET"

	Viz_Send(cmd)
End Sub


Sub Get_VizVisitorsHeadshots()

	Dim cmd
	
	'cmd = "R_VisitorsHeadShots IMAGE*" & HEADSHOT_DIRECTORY_GH & Game.Visitors.Abbreviation & "-WOMEN" & " GET"
	cmd = "R_VisitorsHeadShots IMAGE*" & HEADSHOT_DIRECTORY_GH & Game.Visitors.Abbreviation & " GET"

	Viz_Send(cmd)
End Sub

'=====================================================================
'---------------------------------------------------------------------
Sub ProcessAvailableSponsors(byVal sData)
	Dim tmpArray
	Dim sponsorArray
	Dim i
	Dim lastIndex
	Dim sponsorEntry
	Dim tmp

	sData = Trim(sData)

	tmpArray = Split(sData, chr(32))

	If ubound(tmpArray) < 0 Then
		Exit Sub
	End If

	ReDim sponsorArray(ubound(tmpArray)-1)

	'skipping the header
	for i = 1 to ubound(tmpArray)-1
		'msgbox tmpArray(i)
		sponsorEntry = tmpArray(i) 'ex: IMAGE*SOCCER/IMAGES/SPONSOR/FOXBOX/VISA 
		tmp = Split(sponsorEntry, chr(47)) ' split with "/"
		lastIndex = ubound(tmp) 'last index contains sponor name
		sponsorArray(i) = tmp(lastIndex)
	next

	Plugin.Set_SponsorList sponsorArray
End Sub
'=====================================================================
'---------------------------------------------------------------------
Sub ProcessAvailableTalents(byVal sData)
	Dim tmpArray
	Dim talentArray
	Dim i
	Dim lastIndex
	Dim talentEntry
	Dim tmp

	sData = Trim(sData)

	tmpArray = Split(sData, chr(32))

	If ubound(tmpArray) < 0 Then
		Exit Sub
	End If

	ReDim talentArray(ubound(tmpArray)-1)

	'skipping the header
	for i = 1 to ubound(tmpArray)-1
	'msgbox tmpArray(i)
		talentEntry = tmpArray(i) 'ex: IMAGE*SOCCER/IMAGES/SPONSOR/FOXBOX/VISA 
		tmp = Split(talentEntry, chr(47)) ' split with "/"
		lastIndex = ubound(tmp) 'last index contains sponor name
		talentArray(i) = tmp(lastIndex)
	next

	Plugin.Set_TalentList talentArray
End Sub

'Sub ProcessHeadshots(byVal sData, byVal selectedTeam)
'	Dim headArray
'	Dim tmpArray
'	Dim i
'	Dim lastIndex
'	Dim headEntry
'	Dim tmp
'
'	sData = Trim(sData)
'	tmpArray = Split(sData, chr(32))
'
'	If ubound(tmpArray) < 0 Then
'		Exit Sub
'	End If
'
'	ReDim headArray(ubound(tmpArray)-1)
'
'	for i = 1 to ubound(tmpArray)-1
'		headEntry = tmpArray(i) 'ex: IMAGE*SOCCER/IMAGES/SPONSOR/FOXBOX/VISA 
'		tmp = Split(headEntry, chr(47)) ' split with "/"
'		lastIndex = ubound(tmp) 'last index contains sponor name
'		headArray(i) = tmp(lastIndex)
'	next
'	
'	Call Plugin.Set_HeadsList(headArray, selectedTeam)
'End Sub

Sub ProcessHeadshots(byVal sData, byVal selectedTeam)
	Dim headArray
	Dim tmpArray
	Dim i
	Dim lastIndex
	Dim headEntry
	Dim tmp

	sData = Trim(sData)
	tmpArray = Split(sData, chr(32))

	If ubound(tmpArray) < 0 Then
		Exit Sub
	End If

	ReDim headArray(ubound(tmpArray)-1)

	If ubound(headArray) > 0 Then

		for i = 1 to ubound(tmpArray)-1
			headEntry = tmpArray(i) 'ex: IMAGE*SOCCER/IMAGES/SPONSOR/FOXBOX/VISA 
			tmp = Split(headEntry, chr(47)) ' split with "/"
			lastIndex = ubound(tmp) 'last index contains sponor name
			headArray(i) = tmp(lastIndex)
		next
	Else
		'no headshot names got returned  probably there are no headshots in the folder
		ReDim headArray(1)
		headArray(1) = "ERROR"
	End If

	Call Plugin.Set_HeadsList(headArray, selectedTeam)


End Sub

'=====================================================================
'---------------------------------------------------------------------
Sub RendererSocket_ConnectionEstablished()
	GetTcpConnectionStates
	Log.LogEvent "Renderer Connection Established.", "Debug", 0, "Viz"
End Sub

Sub  PreviewSocket_ConnectionEstablished()
	GetTcpConnectionStates
	Log.LogEvent "Preview Connection Established.", "Debug", 0, "Viz"
End Sub

Sub RendererSecondarySocket_ConnectionEstablished()
	GetTcpConnectionStates
	Log.LogEvent "Renderer Secondatry Connection Established.", "Debug", 0, "Viz"
End Sub


Sub Setup_Renderer() 'this gets called in Startup in GraphicPlugin
	
	Viz_SetupSceneLayer

	Viz_ClearDictionary
	Preview_ClearDictionary

	Delay 5
	
	Get_VizSceneNames

	Delay 5

	Get_VizSponsorNames	

	Delay 5

	Get_VizTalentNames

	Delay 5

	Get_VizHomeHeadshots

	Delay 5
	
	Get_VizVisitorsHeadshots
	
	Delay 5

	SceneIsLoaded = True

	Initialize_SceneData

End Sub

'=====================================================================
'---------------------------------------------------------------------

Sub Renderer_Init()

	'UpdateAggText

	'Settings.Save
End Sub

Sub Initialize_SceneData()

	'Set_SceneByLeague
	Set_FontColor
	'Set_FoxBoxHeader
	'Set_Positions

	'Initialize_Game

	Game_Update

	TurnOffBuildInterface

	Axis_PositionChange("FoxBox")	


End Sub

Sub ReloadScene()

	'forcing both to be true so that all machines get loaded
	Plugin.SendToRenderer = True
	Plugin.SendToPreview = True
	
	Viz_ClearDictionary
	
	Preview_ClearDictionary

	Viz_UnLoadScene

	Delay 5

	Viz_LoadScene

	Delay 10

	Plugin.ResetUserSelectedMode

	TurnOffBuildInterface

        Axis_PositionChange("FoxBox")	

End Sub


'=====================================================================
'---------------------------------------------------------------------
'Sub Set_SceneByLeague()
'	dim sceneType
'
'	If Len(Option_League) < 1 Then
'		sceneType = 1
'	End If
'
'	If Option_League = "UEFA" Then
'		sceneType = 2
'	Else
'		sceneType = 1	
'	End If
'
'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SKIN=" & sceneType)
'
'End Sub

'Sub SwitchToUEFAScene(skinType)
'
'	Select Case skinType
'		Case 0
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SKIN=1")
'		Case 1
'			Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SKIN=2")
'	End Select
'End Sub




'=====================================================================
'---------------------------------------------------------------------
Sub Viz_LoadScene()

	Viz_Send VizLayer & " SET_OBJECT SCENE*" & PluginSettings.SceneDirectory & PluginSettings.SelectedSceneName
	Log.LogEvent "Viz Loading the Scene: " & PluginSettings.SelectedSceneName, "Debug", 0, "ScriptRun"
	Log.LogEvent "RendererSocket.Connected: " &  RendererSocket.Connected, "Debug", 0, "ScriptRun"
	
	If RendererSocket.Connected Then
		Plugin.RefreshSceneName	
	End If

End Sub

'=====================================================================
'---------------------------------------------------------------------
Sub Viz_UnLoadScene()

	Viz_Send COMMAND_HEADER & "RENDERER SET_OBJECT"
	Viz_Send COMMAND_HEADER & "RENDERER*FRONT_LAYER SET_OBJECT"
	Viz_Send COMMAND_HEADER & "RENDERER*BACK_LAYER SET_OBJECT"
	Viz_Send COMMAND_HEADER & "SCENE CLEANUP"
	Viz_Send COMMAND_HEADER & "IMAGE CLEANUP"
	Viz_Send COMMAND_HEADER & "FONT CLEANUP"
	Viz_Send COMMAND_HEADER & "MATERIAL CLEANUP"
	Viz_Send COMMAND_HEADER & "GEOM CLEANUP"

	Log.LogEvent "Viz UnLoading the Scene ", "Debug", 0, "ScriptRun"
End Sub


'=====================================================================
'---------------------------------------------------------------------
Sub Viz_SetupSceneLayer()

	PluginSettings.SceneLayer = 0 'this is default for now

	Select Case PluginSettings.SceneLayer
		Case 0 'front
			VizLayer = COMMAND_HEADER & "RENDERER*FRONT_LAYER"
		Case 1 'middle
			VizLayer = COMMAND_HEADER & "RENDERER"
		Case 2 'back
			VizLayer = COMMAND_HEADER & "RENDERER*BACK_LAYER"
	End Select
End Sub


'=====================================================================
'---------------------------------------------------------------------
Function Viz_UpdateDictionary(p_sKey, p_sValue)
    Dim sValue
    Dim fUpdate
    
    On Error Resume Next
    
    If m_objVizDictionary Is Nothing Then
    	Set m_objVizDictionary = CreateObject("Scripting.Dictionary")
    End If
    
    With m_objVizDictionary 

        ' Check if Key exists
        If Not .Exists(p_sKey) Then
            ' Key doesn't exist, create it
            .Add p_sKey, " "
        End If
        
        ' Get Key value
        sValue = .Item(p_sKey)
        
        If sValue <> p_sValue Then
            
            ' Store new key value
            .Item(p_sKey) = p_sValue
            
            ' Return Update
            fUpdate = True
            
        Else
            fUpdate = False
        End If

    End With
        
    Viz_UpdateDictionary = fUpdate
    
End Function

'=====================================================================
'---------------------------------------------------------------------
Sub Viz_ClearDictionary()

	' Simply create a new instance of the Dictionary object 
	Set m_objVizDictionary = Nothing
	Set m_objVizDictionary = CreateObject("Scripting.Dictionary")

	Log.LogEvent "Viz Dictionary cleared.", "Debug", 0, "Viz"

End Sub

'=====================================================================
'---------------------------------------------------------------------
Function Preview_UpdateDictionary(p_sKey, p_sValue)
    Dim sValue
    Dim fUpdate
    
    On Error Resume Next
    
    If m_objPreviewDictionary Is Nothing Then
    	Set m_objPreviewDictionary = CreateObject("Scripting.Dictionary")
    End If
    
    With m_objPreviewDictionary 

        ' Check if Key exists
        If Not .Exists(p_sKey) Then
            ' Key doesn't exist, create it
            .Add p_sKey, " "
        End If
        
        ' Get Key value
        sValue = .Item(p_sKey)
        
        If sValue <> p_sValue Then
            
            ' Store new key value
            .Item(p_sKey) = p_sValue
            
            ' Return Update
            fUpdate = True
            
        Else
            fUpdate = False
        End If

    End With
        
    Preview_UpdateDictionary = fUpdate
    
End Function

Sub Preview_ClearDictionary()

	' Simply create a new instance of the Dictionary object 
	Set m_objPreviewDictionary = Nothing
	Set m_objPreviewDictionary = CreateObject("Scripting.Dictionary")

	Log.LogEvent "Preview Dictionary cleared.", "Debug", 0, "Viz"

End Sub




'=====================================================================
'---------------------------------------------------------------------
' Convert to UTF-8
'		Convert Unicode string to UTF-8 
Function ConvertToUTF8(p_sString)
	Dim lngAsc
	Dim lngIndex
	Dim sReturn, sCharacter
	Dim b1, b2, b3
 
	' Initialize return value
	sReturn = ""
 
	For lngIndex = 1 to Len(p_sString)
 
		' Get Ascii value
		lngAsc = AscW(Mid(p_sString, lngIndex, 1))
	 
		If lngAsc < 128 Then
			sCharacter = ChrW(lngAsc)

		ElseIf lngAsc < 2048 Then
			b1 = lngAsc Mod 64
			b2 = (lngAsc - b1) / 64
			sCharacter = ChrW(&hc0 + b2) & Chr(&h80 + b1)

		ElseIf lngAsc < 65536 Then
			b1 = lngAsc Mod 64
			b2 = ((lngAsc - b1) / 64) Mod 64
			b3 = (lngAsc - b1 - (64 * b2)) / 4096
			sCharacter = ChrW(&he0 + b3) & ChrW(&h80 + b2) & Chr(&h80 + b1)

		End If

		' Build return string
		sReturn = sReturn & sCharacter
		
	Next
	
	' Return string
	ConvertToUTF8 = sReturn
	
End Function

Sub RefreshSettings()
	Set_FontColor
	'Set_FoxBoxHeader
	'Set_AggText
End Sub


'=====================================================================
'---------------------------------------------------------------------
Sub Set_Positions()
	Axis_PositionChange("FoxBox")
	Axis_PositionChange("Bug")
End Sub

'=====================================================================
'---------------------------------------------------------------------
Sub Axis_PositionChange(p_sAxisCode)


	Log.LogEvent "Position Change for " & p_sAxisCode & " to " & Settings.FoxBoxOffsetAxis.PositionX & ";" & Settings.FoxBoxOffsetAxis.PositionY, "Debug", 0, "Viz"
	
	If p_sAxisCode = "FoxBox" Then

		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOXBOX_P_X=" & FormatNumber(Settings.FoxBoxOffsetAxis.PositionX,3) )
		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET FOXBOX_P_X=" & FormatNumber(Settings.FoxBoxOffsetAxis.PositionY,3) )


		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET OFFSET_X=" & FormatNumber(Settings.FoxBoxOffsetAxis.PositionX,3) )
		Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET OFFSET_Y=" & FormatNumber(Settings.FoxBoxOffsetAxis.PositionY,3) )
	Else
		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET BUG_P_X=" & FormatNumber(Settings.BugOffsetAxis.PositionX,3) )
		'Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET BUG_P_Y=" & FormatNumber(Settings.BugOffsetAxis.PositionY,3) )
	End If





End Sub

'=====================================================================
'   TOOLS
'---------------------------------------------------------------------
' Renderer - Set Background:
'		Add a background to the renderer scene
Sub Renderer_SetBackground(p_iBackground)

	Select Case p_iBackground
		
		Case 1 ' Clear
			Viz_Update "CONTROLS", 0
		Case 2 ' Image
			'Ventuz_Update ".Options.Background.Image", LEAGUE_ELEMENTS_DIRECTORY & "Background.png"
			Viz_Update "CONTROLS", 1
		Case 3 ' Video Input
			'Ventuz_Update ".Options.Background.Type", 3
		Case 4 ' Color
			'Ventuz_Update ".Options.Background.Type", 1
	End Select		

End Sub



'Sub Set_AggText()
'	Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET AGG_TEXT=" & PluginSettings.AggregateDisplayText)
'End Sub





