'=====================================================================
' FOX Sports VBScript Database Module
'---------------------------------------------------------------------
' Etienne Laneville (etiennel@fox.com)
' August 2001
'---------------------------------------------------------------------
' Functions to load and save to database
'=====================================================================

'=====================================================================
'   CONNECTION-RELATED PROCEDURES
'---------------------------------------------------------------------

' Database Set:
'       Set reference to database for Football objects
Sub Database_Set()
    Dim sSource
    
	With Soccer

		'sSource = APPLICATION_DIRECTORY & Option_League & "\Stats\" & Option_Database
		sSource = APPLICATION_DIRECTORY & "Stats\" & Option_Database

		.Database.Source = sSource
		.Database.Connect

	End With

    Log.LogEvent "Databases opened. " & sSource, "Message", 0, "Database"
    
End Sub

' Database Import:
'       Import all data from database
Sub Database_Import()
    Dim sngTimer
    
	sngTimer = Timer
	Soccer.DatabaseImport
	Log.LogEvent "Data Imported. " & FormatNumber((Timer - sngTimer), 2) & " seconds." , "Message", 0, "Database"
    
End Sub
