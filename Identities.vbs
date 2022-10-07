'=====================================================================
' FSN 2008
' FOX Sports VBScript Identities Scripting Module
'---------------------------------------------------------------------
' Etienne Laneville (Etienne.Laneville@FOX.com)
' August 2008
'---------------------------------------------------------------------
' Procedures to handle FoxBox Identities
'=====================================================================

Option Explicit

'=====================================================================
'   PROCEDURES
'---------------------------------------------------------------------

' Create Identities:
'		Create Identities from available Elements
Sub Identities_CreateIdentities()
	Dim objFileSystem
	Dim objFile
	Dim objRootFolder
	Dim objFolder

	Dim objIdentity
	Dim sIdentityName
	Dim sFileName


	On Error Resume Next

	' Create File System object
	Set objFileSystem = CreateObject("Scripting.FileSystemObject")

	' Get Root Folder object
	Set objRootFolder = objFileSystem.GetFolder(COMMON_DIRECTORY & "Elements\Identities\")

	' Loop through all subfolders
	' Each subfolder represents an Identity
	For Each objFolder In objRootFolder.SubFolders
	
		' Store Identity Name
		sIdentityName = objFolder.Name

		' Check if Bug file exists
		'sFileName = objFolder.Path & "\" & sIdentityName & "_Bug.tga"
		sFileName = objFolder.Path & "\" & sIdentityName & "_Bug.png"

		If objFileSystem.FileExists(sFileName) Then

			' Retrieve Identity
			Set objIdentity = Identities.GetIdentityByName(sIdentityName)
			
			If objIdentity Is Nothing Then
				' Identity didn't exist, create one
				Set objIdentity = Identities.Create()
			End If
			
			With objIdentity
				.Name = sIdentityName
				.Description = sIdentityName
				.Logo = sFileName
			End With
		
			' Look for other files
			For Each objFile In objFolder.Files

				' Color file
				If UCase(objFileSystem.GetExtensionName(objFile)) = "COLOR" Then
					With objIdentity
						.Color.Init = objFileSystem.GetBaseName(objFile)
					End With
				End If

				' Family file
				If UCase(objFileSystem.GetExtensionName(objFile)) = "FAMILY" Then
					With objIdentity
						.Family = objFileSystem.GetBaseName(objFile)
					End With
				End If

			Next

			' Defaults
			If objIdentity.Color.Init = "0,0,0" Then objIdentity.Color.Init = "0,121,194"

		End If
	
	Next

End Sub

Sub SetNetworkLogoFile()
	Dim objIdentity
	' Retrieve style
        Set objIdentity = Identities.GetIdentityByName(Settings.Identity)
        If Not objIdentity Is Nothing Then
                ' Set style
                Ventuz_Update ".FoxBox.Network.File", "file:///" & objIdentity.Logo
        End If
End Sub
