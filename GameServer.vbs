'=====================================================================
' FSN Hockey 2005
' FOX Sports VBScript Constants Scripting Module
'---------------------------------------------------------------------
' Procedures called by Hockey FoxBox to handle broadcasting of
' FoxBox data out to TCP clients.
'=====================================================================

Option Explicit

'=====================================================================
'   DECLARATIONS
'---------------------------------------------------------------------

Dim m_sGameServer_LastPacket

'=====================================================================
'   PROCEDURES
'---------------------------------------------------------------------

'   Game Info Server
'---------------------------------------------------------------------

' Game Info Connect
'	A new client has connected to our server.
'	Broadcast out updates.
Sub GameInfo_Connect()

End Sub

' Game Info: Game
'		Broadcast out a game update.
Sub GameInfo_Game()

End Sub
    

' Game Info: Clock
'		Broadcast out a clock update.
Sub GameInfo_Clock()

End Sub

'   Simple Game Protocol
'---------------------------------------------------------------------

' Simple Game Protocol - Parse Data
'		Parse incoming data from Game Server
Sub SimpleGameProtocol_ParseData(p_sPacket)
	Dim arrPackets
	Dim sPacket
	Dim iStartPosition
	Dim iEndPosition
	Dim iCounter

	' Store last packet
	m_sGameServer_LastPacket = p_sPacket
	
	' Get delimiter positions
	iStartPosition = InStr(p_sPacket, "[")
	iEndPosition = InStr(p_sPacket, "]")
	
	If iStartPosition > 0 And iEndPosition > 0 Then

		' Remove delimiters
		sPacket = Mid(p_sPacket, iStartPosition + 1, iEndPosition - iStartPosition - 1)

		If InStr(sPacket, "|")  > 0 Then
			' Split packets
			arrPackets = Split(sPacket, "|")
			For iCounter = 1 to UBound(arrPackets) + 1
				SimpleGameProtocol_ParseGamePacket(arrPackets(iCounter - 1))
			Next
		Else
			SimpleGameProtocol_ParseGamePacket(sPacket)
		End If

		' Update CSS Games
		GameServer_UpdateCSS
			
	End If

End Sub

' Simple Game Protocol - Parse Game Packet
'		Parse game packet
Sub SimpleGameProtocol_ParseGamePacket(p_sPacket)
	Dim objGame
	Dim sUniqueID
	Dim arrValues
	Dim objTeam
	Dim lngTeamID

	On Error Resume Next
	
	' Validate packet
	If InStr(p_sPacket, ",") = 0 Then Exit Sub ' No data
	
	' Split packet into array
	arrValues = Split(p_sPacket, ",")
	
	' Get Game ID
	sUniqueID = arrValues(0)
	
	' Validate Game ID
	If sUniqueID = "" Then Exit Sub

	' Get Game
	Set objGame = Soccer.Games.GetGameByUniqueID(sUniqueID)
	
	If objGame Is Nothing Then
		' Create game
		Set objGame = Soccer.Games.Create
	End If

	With objGame

		' Store direct properties
		.UniqueID = sUniqueID
		.Clock.Value = arrValues(7)

		.Period = CInt(arrValues(8))
		.PeriodStart = CBool(arrValues(9))
		.PeriodEnd = CBool(arrValues(10))
		.Halftime = CBool(arrValues(11))
		.Final = CBool(arrValues(12))

		.HomeScore = CInt(arrValues(3))
		.VisitorsScore = CInt(arrValues(6))

		.Title = arrValues(13)
		.Location = arrValues(14)
		
		' Home team
		lngTeamID = CLng(arrValues(1))
		Set objTeam = Soccer.Teams.Team(lngTeamID, False)
		If .Home Is Nothing Then
			' Set team
			Set .Home = objTeam
		Else
			If .Home.ID <> objTeam.ID Then
				' Set team
				Set .Home = objTeam
			End If
		End If

		' Visitor team
		lngTeamID = CLng(arrValues(4))
		Set objTeam = Soccer.Teams.Team(lngTeamID, False)
		If .Visitors Is Nothing Then
			' Set team
			Set .Visitors = objTeam
		Else
			If .Visitors.ID <> objTeam.ID Then
				' Set team
				Set .Visitors = objTeam
			End If
		End If
	
	End With
	
End Sub

' Simple Game Protocol - Game Update
'		Send game packet to game server
Sub SimpleGameProtocol_GameUpdate()
	Dim sPacket

	On Error Resume Next

	' Check if Game Server socket is connected
	If Not GameServerSocket.Connected Then Exit Sub
	
	' Build packet
	sPacket = SimpleGameProtocol_BuildGamePacket(Game)
	
	' Send packet to Game Server
	GameServerSocket.Send sPacket
	
End Sub

' Simple Game Protocol - Build Game Packet
'		Build game packet from given game
Function SimpleGameProtocol_BuildGamePacket(p_objGame)
	Dim sPacket

	With p_objGame

		sPacket = .UniqueID & "," & _
			.Home.ID & "," & _
			.Home.Abbreviation & "," & _
			.HomeScore & "," & _
			.Visitors.ID & "," & _
			.Visitors.Abbreviation & "," & _
			.VisitorsScore & "," & _
			.Clock.Value & "," & _
			.Period & "," & _
			.PeriodStart & "," & _
			.PeriodEnd & "," & _
			.Halftime & "," & _
			.Final & "," & _
			.Title & "," & _
			.Location

	End With

	SimpleGameProtocol_BuildGamePacket = "[" & sPacket & "]"
	
End Function

'   Game Server
'---------------------------------------------------------------------

' Game Server: Game Update
'		Push local game data to Game Server
Sub GameServer_GameUpdate()
	Dim sXML

	On Error Resume Next

	SimpleGameProtocol_GameUpdate
	
	Exit Sub
	
	' Check if Game Server socket is connected
	If Not GameServerSocket.Connected Then Exit Sub

	' Get Game XML
	sXML = Game.SerializationXML
	
	' Wrap XML
	sXML = Soccer.WrapXML(sXML)
	
	' Push XML to Game Server
	GameServerSocket.Send sXML

End Sub

' Game Server: Update CSS
'		Store Game data into CSS objects
Sub GameServer_UpdateCSS()
	Dim objGame
	Dim objCSSGame

	On Error Resume Next
	
	For Each objGame In Soccer.Games
	
		' Get CSS Game
		Set objCSSGame = Interface.ctlCSS.CSS.Games.Game(objGame.ID) ' Match game by ID
		
		With objCSSGame

			.Home.Name = objGame.Home.Name
			.Home.City = objGame.Home.City
			.Home.Abbreviation = objGame.Home.DisplayAbbreviation
			.Home.Score = objGame.HomeScore

			.Visitors.Name = objGame.Visitors.Name
			.Visitors.City = objGame.Visitors.City
			.Visitors.Abbreviation = objGame.Visitors.DisplayAbbreviation
			.Visitors.Score = objGame.VisitorsScore

			If objGame.Final Then 
				.Status = "Final"
				.Clock = ""
			Else
				.Status = objGame.PeriodText
				.Clock = objGame.Clock.Value
			End If
			
			.Miscellaneous(1) = objGame.Title
			.Miscellaneous(2) = objGame.Location
			
		End With
		
	Next

End Sub