'=====================================================================
' FOX NE&O
' Sockets Module
'---------------------------------------------------------------------
' Etienne Laneville (Etienne.Laneville@FOX.com)
' June 2008
'---------------------------------------------------------------------
' Procedures to handle socket events
'=====================================================================

Option Explicit

'=====================================================================
'   PROCEDURES
'---------------------------------------------------------------------

' Primary Socket Connect:
'		Connection has been established by Primary socket
Sub PrimarySocket_Connect(p_sRemoteHostID)

End Sub

' Stats Socket Connect:
'		Connection has been established by Stats socket
Sub StatsSocket_Connect(p_sRemoteHostID)

End Sub

' Primary Socket Data Arrival:
'		Data has been received on socket
Sub PrimarySocket_DataArrival(p_sData)

    ' Log.LogEvent p_sData, "Debug", 0, "Primary Socket"

End Sub

'   Renderer
'---------------------------------------------------------------------

' Renderer Socket Connect:
'		Connection has been established by Renderer socket
Sub RendererSocket_Connect(p_sRemoteHostID)

	RendererSocket_ConnectionEstablished
End Sub

' Renderer Socket Data Arrival:
'		Data has been received on socket
Sub RendererSocket_DataArrival(p_sData)
	Call ProcessResponseFromTheEngine(p_sData)

'	Log.LogEvent p_sData, "Debug", 0, "Renderer Socket"

End Sub

'   Renderer Secondary
'---------------------------------------------------------------------
' Renderer Socket Connect:
'		Connection has been established by Renderer socket
Sub RendererSecondarySocket_Connect(p_sRemoteHostID)
	RendererSecondarySocket_ConnectionEstablished
End Sub

' Renderer Socket Data Arrival:
'		Data has been received on socket
Sub RendererSecondarySocket_DataArrival(p_sData)
	'
End Sub

'  Preview 
'---------------------------------------------------------------------

' Renderer Socket Connect:
'		Connection has been established by Renderer socket
Sub PreviewSocket_Connect(p_sRemoteHostID)
	PreviewSocket_ConnectionEstablished	
End Sub

' Renderer Socket Data Arrival:
'		Data has been received on socket
Sub PreviewSocket_DataArrival(p_sData)
	
	If Not RendererSocket.Connected Then
		Call ProcessResponseFromTheEngine(p_sData)
	End If

End Sub

Sub GetPreveiwSocketConnectionState()
	Plugin.previewSocketState = PreviewSocket.Connected
End Sub

'   Game Server
'---------------------------------------------------------------------

' Game Server Socket Connect:
'		Connection has been established by Game Server socket
Sub GameServerSocket_Connect(p_sRemoteHostID)

	' Send Game Data to Game Server
	SimpleGameProtocol_GameUpdate
	
End Sub

' Game Server Socket Data Arrival:
'		Data has been received on socket
Sub GameServerSocket_DataArrival(p_sData)

	' Log.LogEvent p_sData, "Debug", 0, "Game Server Data Arrival"

	' Use Simple Game Protocol to parse incoming data
	SimpleGameProtocol_ParseData p_sData

End Sub


Sub ConnectToPreviewSocket()

	Plugin.TcpSetProperties

	If  PreviewSocket.Connected Then
		DisconnectFromPreviewSocket
		PreviewSocket.Connect
	Else
	 	PreviewSocket.Connect
	End If
	
	Log.LogEvent "Establish Connection to Preview from Script.", "ConnectToPreviewSocket", 0, "PreviewSocket"

End Sub

Sub DisconnectFromPreviewSocket()
	 PreviewSocket.Disconnect
	Log.LogEvent "Disconnecting from Preview from Script.", "DisconnectFromPreviewSocket", 0, "PreviewSocket"
End Sub


Sub ConnectToSecondarySocket()

	Plugin.TcpSetProperties

	If  RendererSecondarySocket.Connected Then
		DisconnectFromSecondarySocket
		RendererSecondarySocket.Connect
	Else
	 	RendererSecondarySocket.Connect
	End If
	
	Log.LogEvent "Establish Connection to Renderer Secondary from Script.", "ConnectToSecondarySocket", 0, "RendererSecondary"

End Sub

Sub DisconnectFromSecondarySocket()
	 RendererSecondarySocket.Disconnect
	Log.LogEvent "Disconnecting from Secondary Renderer from Script.", "DisconnectFromSecondarySocket", 0, "RendererSecondary"
End Sub

Sub ConnectToRendererSocket()

	RendererSocket.RemoteHost = Settings.RendererSocket.Address
    	RendererSocket.RemotePort = Settings.RendererSocket.Port

	If  RendererSocket.Connected Then
		DisconnectFromRendererSocket
		RendererSocket.Connect
	Else
	 	RendererSocket.Connect
	End If
	
	Log.LogEvent "Establish Connection to Renderer Socket from Script.", "ConnectToRendererSocket", 0, "Renderer"

End Sub

Sub DisconnectFromRendererSocket()
	RendererSocket.Disconnect
	Log.LogEvent "Disconnecting from Renderer from Script.", "DisconnectFromRendererSocket", 0, "Renderer"
End Sub


Sub GetTcpConnectionStates()
	Plugin.GetSocketStates
	Plugin.TcpLedsRefresh

	Interface.TcpLedsRefresh
End Sub


Sub GetRendererSocketStateInfo()

	With Plugin
		.RendererSocketConnectionState = RendererSocket.Connected 
		.RendererSocketConnectionIPAddress = Settings.RendererSocket.Address
		.RendererSocketConnectionPort = Settings.RendererSocket.Port
	End With 

End Sub


