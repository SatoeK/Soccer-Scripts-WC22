'=====================================================================
' FSN 2008
' FOX Sports VBScript Sponsors Scripting Module
'---------------------------------------------------------------------
' Etienne Laneville (Etienne.Laneville@FOX.com)
' August 2008
'---------------------------------------------------------------------
' Procedures to handle FoxBox Sponsors
'=====================================================================

Option Explicit

'=====================================================================
'   DECLARATIONS
'---------------------------------------------------------------------

Dim m_objFileSystem

'=====================================================================
'   PROCEDURES
'---------------------------------------------------------------------

'   Initialization
'---------------------------------------------------------------------

' Load Sponsors:
'		Create Sponsors from available Elements
Sub Sponsors_Load()
	' Create FileSystem Object
	Set m_objFileSystem = CreateObject("Scripting.FileSystemObject")

	' Read sponsors from hard drive 
	CreateSponsorsFromFolders COMMON_DIRECTORY & "Elements\Sponsors\"

	' Load user's pre-selected sponsor list
	SponsorManager.SelectedListFile = APPLICATION_DIRECTORY & Option_League & "\SponsorList.txt"
	SponsorManager.LoadSelectedList

	Sponsors_CreateSponsoredItem "GameFlow", "GameFlow"
	'Sponsors_CreateSponsoredItem "ScoreSponsor", "ScoreSponsor"

End Sub


' Create Sponsored Item:
'		Create an item in the SponsoredItems collection
Sub Sponsors_CreateSponsoredItem(p_sName, p_sCategory)
	Dim objItem
	
	Set objItem = SponsorManager.SponsoredItems.GetSponsoredItemByName(p_sName)
	If objItem Is Nothing Then
		' Create it
		Set objItem = SponsorManager.SponsoredItems.Create
	End If
	
	With objItem
		.Name = p_sName
		.Category = p_sCategory
	End With

End Sub


' Create Sponsors from Folders:
'       Create Sponsors from available elements existing in folders
Sub CreateSponsorsFromFolders(p_sFolder)
    Dim objFile
    Dim objRootFolder
    Dim objFolder

    Dim objSponsor
    Dim objVersion
    Dim sSponsorName
    Dim sFileName

    ' Get Root Folder object
    Set objRootFolder = m_objFileSystem.GetFolder(p_sFolder)

    ' Hold events
    SponsorManager.MasterList.HoldEvents = True
    
    ' Loop through all subfolders
    ' Each subfolder represents an Sponsor
    For Each objFolder In objRootFolder.SubFolders
    
        ' Store Sponsor Name
        sSponsorName = objFolder.Name

        ' Retrieve Sponsor
        Set objSponsor = SponsorManager.MasterList.GetSponsorByName(sSponsorName)

        If objSponsor Is Nothing Then
            ' Sponsor didn't exist, create one
            Set objSponsor = SponsorManager.MasterList.Create()
        End If

        With objSponsor
            .Name = sSponsorName
        End With
    
        ' Check for any existing Versions
        CheckSponsorVersion objSponsor, objFolder.Path & "\{name}" & "_GameFlow", "GameFlow", "General", 1
        CheckSponsorVersion objSponsor, objFolder.Path & "\{name}" & "_ScoreSponsor", "ScoreSponsor", "General", 1

	'CheckSponsorVersion objSponsor, objFolder.Path & "\{name}" & "_FoxBox_Horizontal", "FoxBox", "General", 1

        'CheckSponsorVersion objSponsor, objFolder.Path & "\{name}" & "_Scoreboard", "Scoreboard", "Default", 1
        'CheckSponsorVersion objSponsor, objFolder.Path & "\{name}" & "_Scoreboard_Horizontal", "Scoreboard", "Default", 2
        'CheckSponsorVersion objSponsor, objFolder.Path & "\{name}" & "_Scoreboard_Square", "Scoreboard", "Default", 3
        'CheckSponsorVersion objSponsor, objFolder.Path & "\{name}" & "_Scoreboard_Vertical", "Scoreboard", "Default", 4

    Next

    ' Hold events
    SponsorManager.MasterList.HoldEvents = False

End Sub

' Sponsors Check Version
'       Checks if the specifed sponsor exists
Sub CheckSponsorVersion(p_objSponsor, p_sPath, p_sCategoryName, p_sVersionName, p_iVersionIndex)
    Dim objCategory
    Dim objVersion
    Dim sSponsorName
    Dim sMovie
    Dim sTextureTGA
    Dim sTexturePNG
    Dim sPath

    If p_objSponsor Is Nothing Then Exit Sub

    ' Get sponsor name
    sSponsorName = p_objSponsor.Name

    ' Create file names
    sPath = Replace(p_sPath, "{name}", sSponsorName)
    sTextureTGA = sPath & ".tga"
    sTexturePNG = sPath & ".png"
    sMovie = sPath & ".avi"

    If m_objFileSystem.FileExists(sTextureTGA) Or m_objFileSystem.FileExists(sTexturePNG) Or m_objFileSystem.FileExists(sMovie) Then

        ' Retrieve Category
        Set objCategory = p_objSponsor.GetCategoryByName(p_sCategoryName)

        If objCategory Is Nothing Then
            ' Category didn't exist, create one
            Set objCategory = p_objSponsor.Create()
        End If

        With objCategory
            .Name = p_sCategoryName
        End With

        ' Retrieve Version
        Set objVersion = objCategory.GetVersionByName(p_sVersionName)

        If objVersion Is Nothing Then
            ' Version didn't exist, create one
            Set objVersion = objCategory.Create()
        End If

        With objVersion

            .Name = p_sVersionName
            .Index = p_iVersionIndex

            If m_objFileSystem.FileExists(sTextureTGA) Then
                .Texture = sTextureTGA
            Else
                .Texture = ""
            End If

            If m_objFileSystem.FileExists(sTexturePNG) Then
                .Texture = sTexturePNG
            End If

            If m_objFileSystem.FileExists(sMovie) Then
                .Movie = sMovie
            Else
                .Movie = ""
            End If

        End With

    End If

End Sub


'   Animation
'---------------------------------------------------------------------
Sub GameFlow_Set()
	Dim objSponsor
	Dim objCategory
	Dim objVersion
	Dim sSponsor
	

	sSponsor =  SponsorManager.SponsoredItems.GetSponsoredItemByName("GameFlow").Sponsor
	
   	Log.LogEvent "Sponsor: Set GameFlow To '" & sSponsor & ".", "Debug", 0, "Renderer"

	' Get Sponsor object
	Set objSponsor = Sponsors.GetSponsorByName(sSponsor)
	
	If Not objSponsor Is Nothing Then

		' Get Category
		Set objCategory = objSponsor.GetCategoryByName("GameFlow")

		If Not objCategory Is Nothing Then

			' Get Version
			Set objVersion = objCategory.GetVersionByName("General")

			If Not objVersion Is Nothing Then

				' Check if Movie exists
				If objVersion.Movie <> "" Then
					'Viz_Update "GAMEFLOW_FILE", objVersion.Movie
					Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET GAMEFLOW_FILE=" & objVersion.Movie)
					Log.LogEvent "Sponsor: using Movie version (" & objVersion.Movie & ").", "Debug", 0, "Renderer"
				End If

			Else
				Log.LogEvent "Sponsor: Version not found.", "Debug", 0, "Renderer"
			End If

		End If

	Else
		Log.LogEvent "Sponsor: Sponsor not found.", "Debug", 0, "Renderer"
	End If

End Sub




'Sub ScoreSponsor_Set()
'	Dim objSponsor
'	Dim objCategory
'	Dim objVersion
'	Dim sSponsor
'	
'
'	sSponsor =  SponsorManager.SponsoredItems.GetSponsoredItemByName("ScoreSponsor").Sponsor
'	
'   	Log.LogEvent "Sponsor: Set ScoreSponsor To '" & sSponsor & ".", "Debug", 0, "Renderer"
'
'	' Get Sponsor object
'	Set objSponsor = Sponsors.GetSponsorByName(sSponsor)
'	
'	If Not objSponsor Is Nothing Then
'
'		' Get Category
'		Set objCategory = objSponsor.GetCategoryByName("ScoreSponsor")
'
'		If Not objCategory Is Nothing Then
'
'			' Get Version
'			Set objVersion = objCategory.GetVersionByName("General")
'
'			If Not objVersion Is Nothing Then
'
'				' Check if Movie exists
'				If objVersion.Movie <> "" Then
'					'Viz_Update "SCORESPON_FILE", objVersion.Movie
'					Viz_Send(VizLayer & "*FUNCTION*DataPool*Data SET SCORESPON_FILE=" & objVersion.Movie)
'					Log.LogEvent "Sponsor: using Movie version (" & objVersion.Movie & ").", "Debug", 0, "Renderer"
'				End If
'
'			Else
'				Log.LogEvent "Sponsor: Version not found.", "Debug", 0, "Renderer"
'			End If
'
'		End If
'
'	Else
'		Log.LogEvent "Score Sponsor: Sponsor not found.", "Debug", 0, "Renderer"
'	End If
'
'End Sub

' Sponsor Set Game Flow:
'		Loads appropriate sponsor based on preset or override
Sub Sponsor_Set(p_sSponsor)
	Dim objSponsor
	Dim objCategory
	Dim objVersion

	Dim objCategory2
	Dim objVersion2	
	
	Dim sSponsor

	' Store local copies of sponsor
	'sSponsor = Interface.Sponsor

  	Log.LogEvent "Sponsor: Set GameFlow To '" & p_sSponsor & ".", "Debug", 0, "Renderer"

	' Get Sponsor object
	Set objSponsor = Sponsors.GetSponsorByName(p_sSponsor)
	
	If Not objSponsor Is Nothing Then

		' Get Category
		Set objCategory = objSponsor.GetCategoryByName("GameFlow")		
		If Not objCategory Is Nothing Then
			' Get Version
			Set objVersion = objCategory.GetVersionByName("General")
			
			If Not objVersion Is Nothing Then
				' Check if Movie exists
				If objVersion.Movie <> "" Then
					Viz_Update "GAMEFLOW_FILE", objVersion.Movie 
					Log.LogEvent "Sponsor: using Movie version (" & objVersion.Movie & ").", "Debug", 0, "Renderer"
				End If
			Else
				Log.LogEvent "Sponsor: Version not found.", "Debug", 0, "Renderer"
			End If
			
		End If


		Set objCategory2 = objSponsor.GetCategoryByName("ScoreSponsor")
		If Not objCategory2 Is Nothing Then
			Set objVersion2 = objCategory2.GetVersionByName("General")

			If Not objVersion2 Is Nothing Then
				If objVersion2.Texture <> "" Then
					Viz_Update "SCORESPON_FILE", objVersion.Movie
	
					Log.LogEvent "Sponsor: using movie version (" & objVersion2.Movie & ").", "Debug", 0, "Renderer"
				End If
			End If
		End If
	Else
		Log.LogEvent "Sponsor: Sponsor not found.", "Debug", 0, "Renderer"
	End If

End Sub

' Sponsor Update:
'		Updates sponsor
Sub Sponsor_Update()

	Log.LogEvent "Sponsor: Update", "Debug", 0, "Renderer"
	Sponsor_Set("")
	
End Sub



Sub GetSponsorsSelected()
	Plugin.GameFlowSelected = Settings.SponsorGameFlow
	Plugin.SponsorSelected = Settings.SponsorFoxBox
End Sub


'Sub ClearSelctedSponsorList()
'	Plugin.ClearSelectedSponsors
'End Sub
'
'Sub RefreshSelectedSponsor(list)
'	Plugin.RefreshSelectedSponsor list
'End Sub
