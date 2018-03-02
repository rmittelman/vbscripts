' AddAccessTrustedLocation.vbs
'
' Adds an Access trusted location for installed versions of Access.
'
' Expects command-line argument of path to add as trusted location
' ex: AddAccessTrustedLocation "C:\folder [\folder [... ]]"

' verify valid path sent
validLoc = WScript.Arguments.Count > 0
If validLoc Then
	Set fso = WScript.CreateObject("Scripting.FileSystemObject")
	trustedLoc = WScript.Arguments(0)
	validLoc = fso.FolderExists(trustedLoc)
End If
If Not validLoc Then
	msg = "Missing folder path to set as Access trusted location." & vbCrLf & "Drag desired folder and drop onto this script."
	VBSMsgBox msg, "Error", vbOK + vbExclamation, 5
	WScript.Quit
End If	

Set WshShell= WScript.CreateObject("WScript.Shell")
access_reg_path = "HKEY_CURRENT_USER\Software\Microsoft\Office\{v}.0\Access\Security\Trusted Locations\"

' check for various versions of Office.
' for each version found, set trusted location.

' Office 2010
regKey = Replace(access_reg_path, "{v}", "14")
isVersionThere = RegKeyExists(regKey)
If isVersionThere Then
	result = AddTrustedLoc(regKey, trustedLoc, 1)
End If

' Office 2013
regKey = Replace(access_reg_path, "{v}", "15")
isVersionThere = RegKeyExists(regKey)
If isVersionThere Then
	result = AddTrustedLoc(regKey, trustedLoc, 1)
End If

' Office 2016
regKey = Replace(access_reg_path, "{v}", "16")
isVersionThere = RegKeyExists(regKey)
If isVersionThere Then
	result = AddTrustedLoc(regKey, trustedLoc, 1)
End If

WScript.Quit

'Adds a trusted Office location
'parentKey:		Example: "HKEY_CURRENT_USER\.....\Trusted Locations\"
'path:			Full path to new trusted folder
'subFolders:	Send 1 to trust sub-folders, 0 to not trust.
Function AddTrustedLoc(parentKey, path, subFolders)
	
	AddTrustedLoc = False
	
	' find available location number
	foundEmptyLoc = False
	loc_key = parentKey & "Location{x}\"
	For i = 0 To 100
		locationKey = replace(loc_key, "{x}", CStr(i))
		If Not RegKeyExists(locationKey) Then
			foundEmptyLoc = True
			Exit For
		End If
	Next
	
	' exit if no locations available
	If Not foundEmptyLoc Then
		AddTrustedLoc = False
	
	' if location available, continue
	Else
		newKey = parentKey & locNo & "\"
		On Error Resume Next
		WshShell.RegWrite locationKey & "Path", path, "REG_SZ"
		If Err.Number = 0 Then WshShell.RegWrite locationKey & "AllowSubfolders", subFolders, "REG_DWORD"
		If Err.Number = 0 Then WshShell.RegWrite locationKey & "Date", FormatDateTime(Now()), "REG_SZ"
		If Err.Number = 0 Then WshShell.RegWrite locationKey & "Description", "Added by AddAccessTrustedLocation.vbs", "REG_SZ"
		AddTrustedLoc = (Err.Number = 0)
		On Error Goto 0
	End If
	
End Function

'Indicates whether a registry key exists
Function RegKeyExists(regKey)
	RegKeyExists = False
	On Error Resume Next
	WshShell.RegRead(regKey)
	RegKeyExists = (Err.Number = 0)
	On Error Goto 0
End Function

' <VBSMsgBox>
'  Display message box, retrieve value.
' </summary>
' <param name="message">Message to display.</param>
' <param name="title">Title caption.</param>
' <param name="flags">Flags to control buttons, icon, etc.</param>
' <param name="secondsToStay">If > 0, seconds for MsgBox to stay. Returns -1 result.</param>
' <remarks></remarks>
Function VBSMsgBox(message, title, flags, secondsToStay)

	Dim WshShell, result
	Set WshShell = WScript.CreateObject("WScript.Shell")
	result = WshShell.Popup(message, secondsToStay, title, flags)

	VBSMsgBox = result
	Set WshShell = Nothing

End Function
' </VBSMsgBox>
