' AccessTrustThisFolder.vbs
'
' Adds folder where script is running from as an Access trusted location for installed versions of Access.
'

Const HKEY_CLASSES_ROOT   = &H80000000
Const HKEY_CURRENT_USER   = &H80000001
Const HKEY_LOCAL_MACHINE  = &H80000002
Const HKEY_USERS          = &H80000003
Const ACCESS_KEY_PATH     = "Software\Microsoft\Office\{v}.0\Access\Security\Trusted Locations\"

Dim oReg, HIVE, computer
HIVE = HKEY_CURRENT_USER
computer = "."
Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & computer & "\root\default:StdRegProv")

Dim fso
Set fso = WScript.CreateObject("Scripting.FileSystemObject")

Dim myPath
myPath = WScript.ScriptFullName
trustedLoc = fso.GetParentFolderName(myPath)

' Make sure we're in the right folder!!!!!!!
If LCase(myPath) <> "c:\pols" And LCase(myPath) <> "c:\mlg" Then
	VBSMsgBox "This is not being run from proper location. Exitting...", "Error", vbExclamation, 3
	WScript.Quit
End If

' check for various versions of Office.
' for each version found, set trusted location.

' Office 2010
regKey = Replace(ACCESS_KEY_PATH, "{v}", "14")
If RegKeyExists(HIVE, regKey) Then
	If Not TrustedLocationExists(HIVE, regKey, trustedLoc) Then
		result = AddTrustedLoc(regKey, trustedLoc, 1)
	End If
End If

' Office 2013
regKey = Replace(ACCESS_KEY_PATH, "{v}", "15")
If RegKeyExists(HIVE, regKey) Then
	If Not TrustedLocationExists(HIVE, regKey, trustedLoc) Then
		result = AddTrustedLoc(regKey, trustedLoc, 1)
	End If
End If

' Office 2016
regKey = Replace(ACCESS_KEY_PATH, "{v}", "16")
If RegKeyExists(HIVE, regKey) Then
	If Not TrustedLocationExists(HIVE, regKey, trustedLoc) Then
		result = AddTrustedLoc(regKey, trustedLoc, 1)
	End If
End If

fso.DeleteFile myPath
WScript.Quit

'Adds a trusted Office location
'parentKey:		Example: "HKEY_CURRENT_USER\.....\Trusted Locations\"
'trustedLoc:	Full path to new trusted folder
'subFolders:	Send 1 to trust sub-folders, 0 to not trust.
Function AddTrustedLoc(parentKey, path, subFolders)
	
	Dim foundEmptyLoc, loc_key, locationKey
	
	AddTrustedLoc = False
	
	' find available location number
	foundEmptyLoc = False
	loc_key = parentKey & "Location{x}\"
	For i = 0 To 100
		locationKey = replace(loc_key, "{x}", CStr(i))
		If Not RegKeyExists(HIVE, locationKey) Then
			foundEmptyLoc = True
			Exit For
		End If
	Next
	
	' exit if no locations available
	If Not foundEmptyLoc Then
		AddTrustedLoc = False
	
	' if location available, continue
	Else
		oReg.CreateKey HIVE, locationKey
		oReg.SetStringValue HIVE, locationKey, "Path", trustedLoc
		oReg.SetDWORDValue HIVE, locationKey, "AllowSubfolders", subFolders
		oReg.SetStringValue HIVE, locationKey, "Date", FormatDateTime(Now())
		oReg.SetStringValue HIVE, locationKey, "Description", "Added by " & WScript.ScriptName
		AddTrustedLoc = True
	End If
	
End Function

'Indicates whether a registry key exists
Function RegKeyExists(HIVE, regKey)
	RegKeyExists = (oReg.EnumKey(HIVE, regKey, "", "") = 0)
End Function

'Indicates whether a trusted location exists
Function TrustedLocationExists(HIVE, regKey, trustedLoc)

	Dim arrSubKeys, subKey, strValue
	TrustedLocationExists = False
	
	oReg.EnumKey HIVE, regKey, arrSubKeys
	If IsArray(arrSubKeys) Then
		For Each subKey In arrSubKeys
			oReg.GetStringValue HIVE, regKey & "\" & subKey, "Path", strValue
			If Not IsEmpty(strValue) Then
				If strValue = trustedLoc Then
					TrustedLocationExists = True
					Exit For
				End If
			End If
		Next
	End If

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
