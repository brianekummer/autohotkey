;---------------------------------------------------------------------------------------------------------------------
; Public Functions
;   SlackStatusUpdate_Initialize
;   SlackStatusUpdate_SetSlackStatusBasedOnNetwork() {
;---------------------------------------------------------------------------------------------------------------------
#NoEnv
#Persistent



;---------------------------------------------------------------------------------------------------------------------
; Define global variables 
;---------------------------------------------------------------------------------------------------------------------
global SlackStatusUpdate_MySlackToken
global SlackStatusUpdate_OfficeNetworks
global SlackStatusUpdate_SlackStatuses
global SlackStatusUpdate_WindowTitles



;---------------------------------------------------------------------------------------------------------------------
; PUBLIC - Initialize global variables from Windows environment variables
;---------------------------------------------------------------------------------------------------------------------
SlackStatusUpdate_Initialize() 
{
	; Variables are read from environment variables, see "Slack Status Update Config.bat" for more details
	EnvGet, SlackStatusUpdate_MySlackToken, SLACK_TOKEN
	EnvGet, SlackStatusUpdate_OfficeNetworks, SLACK_OFFICE_NETWORKS
	
	slackStatusMeeting := SlackStatusUpdate_BuildSlackStatus("SLACK_STATUS_MEETING", "In a meeting|:spiral_calendar_pad:")
	slackStatusWorkingInOffice := SlackStatusUpdate_BuildSlackStatus("SLACK_STATUS_WORKING_OFFICE", "|")
	slackStatusWorkingRemotely := SlackStatusUpdate_BuildSlackStatus("SLACK_STATUS_WORKING_REMOTELY", "Working remotely|:house_with_garden:")
	slackStatusVacation := SlackStatusUpdate_BuildSlackStatus("SLACK_STATUS_VACATION", "Vacationing|:palm_tree:")
	slackStatusLunch := SlackStatusUpdate_BuildSlackStatus("SLACK_STATUS_LUNCH", "At lunch|:hamburger:")
	global SlackStatusUpdate_SlackStatuses := {"meeting": slackStatusMeeting, "workingInOffice": slackStatusWorkingInOffice, "workingRemotely": slackStatusWorkingRemotely, "vacation": slackStatusVacation, "lunch": slackStatusLunch}
	
	; Create an AHK "window group" named "SlackUpdateStatus_WindowTitles" that contains the pattern
	; to find the Slack window
	EnvGet, slackWindowTitle, SLACK_WINDOW_TITLE
	GroupAdd, SlackStatusUpdate_WindowTitles, %slackWindowTitle%
}



;---------------------------------------------------------------------------------------------------------------------
; PUBLIC - Text replacements while in Slack. Note that SlackUpdateStatus_WindowTitles is an AHK window group that
;          contains the names of all the windows that should match, avoiding need to hard-code something here.
;---------------------------------------------------------------------------------------------------------------------
#IfWinActive ahk_group SlackStatusUpdate_WindowTitles
::/lunch::
  SlackStatusUpdate_SetSlackStatusViaKeyboard(SlackStatusUpdate_SlackStatuses["lunch"])
	Return
::/wfh::
  SlackStatusUpdate_SetSlackStatusViaKeyboard(SlackStatusUpdate_SlackStatuses["workingRemotely"])
	Return
::/mtg::
  SlackStatusUpdate_SetSlackStatusViaKeyboard(SlackStatusUpdate_SlackStatuses["meeting"])
	Return
#IfWinActive



;---------------------------------------------------------------------------------------------------------------------
; PUBLIC - Set Slack status, based on what network user is connected to
;---------------------------------------------------------------------------------------------------------------------
SlackStatusUpdate_SetSlackStatusBasedOnNetwork() 
{
  ; Get current Slack status. Network errors are returned as emoji "???"
  mySlackStatusEmoji := SlackStatusUpdate_GetSlackStatusEmoji()

	While (mySlackStatusEmoji = "???") 
	{
	  Sleep, 30000
    mySlackStatusEmoji := SlackStatusUpdate_GetSlackStatusEmoji()
	}
  slackMeetingEmoji := SlackStatusUpdate_SlackStatuses["meeting"]["emoji"]

	If (mySlackStatusEmoji = slackMeetingEmoji) 
	{
	  ; I'm in a meeting, and I ASSUME that my Outlook addin will change my status back when the meeting ends
	}
	Else 
	{
		SlackStatusUpdate_CheckNetworkStatus()
	}
}	



;---------------------------------------------------------------------------------------------------------------------
; Private - Build a slack status object by reading the environment variable envVarName. If this variable is blank or 
;           not set, use the default value provided. 
;             - The text obtained from the environment variable and/or defaultValue should be a pipe-delimited string 
;               with the name and the emoji, such as "At lunch|:hamburger:". It does not matter the order of these.
;---------------------------------------------------------------------------------------------------------------------
SlackStatusUpdate_BuildSlackStatus(envVarName, defaultValue)
{
  EnvGet, slackStatus, %envVarName%
	
  If(slackStatus = "") 
	  slackStatus = %defaultValue%

	parts := StrSplit(slackStatus, "|")
 	slackStatus := RegExMatch(parts[1], "^:.*:$")
	  ? {"text": parts[2], "emoji": parts[1]}
	  : {"text": parts[1], "emoji": parts[2]}
	
	Return slackStatus
}



;---------------------------------------------------------------------------------------------------------------------
; Private - Sets my status in Slack via the Slack keyboard command "/status :emoji: :text:"
;---------------------------------------------------------------------------------------------------------------------
SlackStatusUpdate_SetSlackStatusViaKeyboard(slackStatus)
{
  slackText := slackStatus["text"]
  slackEmoji := slackStatus["emoji"]

  SendInput /status %slackEmoji% %slackText%{enter}
}



;---------------------------------------------------------------------------------------------------------------------
; Private - Check the network status. If user is not connected to any network, then keep checking every 30 seconds. 
;           Once the user is connected, the Slack status will be changed and this method will stop running.
;---------------------------------------------------------------------------------------------------------------------
SlackStatusUpdate_CheckNetworkStatus() {
	officeNetworkSearchExpr = i)connected-%SlackStatusUpdate_OfficeNetworks%

	done := False
	Loop
	{
	  ; Get user's network status. Will be one of these values:
		;    connected-NAME_OF_WIFI_NETWORK
		;    connected-ethernet
		;    disconnected
  	networkStatus := SlackStatusUpdate_GetNetworkStatus()

		If (RegExMatch(networkStatus, officeNetworkSearchExpr)) 
		{
			SlackStatusUpdate_SetSlackStatus(SlackStatusUpdate_SlackStatuses["workingInOffice"])
			done := True
		} 
		Else If (RegExMatch(networkStatus, "i)connected-")) 
		{
			SlackStatusUpdate_SetSlackStatus(SlackStatusUpdate_SlackStatuses["workingRemotely"])
			done := True
		}
		Else
		{
		  ; Wait for 30 seconds and check again
		  Sleep, 30000
		}
	}
	Until done
}	



;---------------------------------------------------------------------------------------------------------------------
; Private - Get the user's network status. Return values are:
;   disconnected.........User is not connected to a wifi network or via wired ethernet
;   connected-ethernet...User is connected to network via wired ethernet cable
;   connected-xxxxx......Where xxxxx is the SSID of the network
;---------------------------------------------------------------------------------------------------------------------
SlackStatusUpdate_GetNetworkStatus() 
{
  wifiStatus := SlackStatusUpdate_GetWifiStatus()
	Return RegExMatch(wifiStatus, "i)connected-")
	  ? wifiStatus
		: SlackStatusUpdate_GetEthernetStatus()
}



;---------------------------------------------------------------------------------------------------------------------
; Private - Get the user's wifi network status. Return values are:
;   disconnected......User is not connected to a wifi network
;   connected-xxxxx...Where xxxxx is the wifi network's SSID
;---------------------------------------------------------------------------------------------------------------------
SlackStatusUpdate_GetWifiStatus() 
{
	cmd = %comspec% /c netsh wlan show interface
	result := SlackStatusUpdate_RunWaitHidden(cmd)    ; Hidden, but get clipboard error many times when I login
	;result := SlackStatusUpdate_RunWaitOne(cmd)        ; IF this doesn't fix my errors, then maybe wait 30 sec after login

  ; Decide if the user is connected or not	
	networkStatus := RegExReplace(result, "s).*?\R\s+State\s+:(\V+).*", "$1")
	networkStatus := RegExReplace(networkStatus, "\s+")
	StringLower, networkStatus, networkStatus
	
  If (networkStatus = "connected")
	{
	  ; Get the name of the wifi network
		networkSSID := RegExReplace(result, "s).*?\R\s+SSID\s+:(\V+).*", "$1")
	  networkSSID := RegExReplace(networkSSID, "\s+")
	  StringLower, networkSSID, networkSSID
		
		networkStatus = %networkStatus%-%networkSSID%
	}
	
	Return networkStatus
}



;---------------------------------------------------------------------------------------------------------------------
; Private - Get the user's ethernet network status. Return values are:
;   disconnected.........User is not connected to wired ethernet
;   connected-ethernet...User is connected to wired ethernet
;---------------------------------------------------------------------------------------------------------------------
SlackStatusUpdate_GetEthernetStatus() 
{
	cmd = %comspec% /c netsh lan show interface
	ethernetStatus := SlackStatusUpdate_RunWaitHidden(cmd)    ; Hidden, but get clipboard error many times when I login
	;ethernetStatus := SlackStatusUpdate_RunWaitOne(cmd)

  Return RegExMatch(ethernetStatus, "i)name\s*: ethernet\R[a-zA-Z0-9 -:]*?\R[a-zA-Z0-9 -:]*?\R[a-zA-Z0-9 -:]*?\R\s*?state\s*: connected.")
	  ? "connected-ethernet" 
		: "disconnected"
}



;---------------------------------------------------------------------------------------------------------------------
; Private - Get the name of the emoji for my current Slack status using the Slack web API
;---------------------------------------------------------------------------------------------------------------------
SlackStatusUpdate_GetSlackStatusEmoji() 
{
	Try 
	{
	  webRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	  webRequest.Open("GET", "https://slack.com/api/users.profile.get?token="SlackStatusUpdate_MySlackToken)
    webRequest.Send()
	  results := webRequest.ResponseText

	  ; The JSON returned by Slack will look something like this: 
	  ;   ..."status_emoji":":house_with_garden:",... 
	  ;   or
    ;   ..."status_emoji":"",...
	  statusEmoji := SubStr(results, InStr(results, "status_emoji"))
	  statusEmoji := SubStr(statusEmoji, 1, InStr(statusEmoji, ","))
	  statusEmoji := RegExReplace(statusEmoji, "i)(status_emoji|""|,)")
	  statusEmoji := RegExReplace(statusEmoji, "i)::", ":")
  }
	Catch 
	{
	  statusEmoji = ???
	}

	Return statusEmoji
}



;---------------------------------------------------------------------------------------------------------------------
; Private - Set my Slack status via the Slack web API
;---------------------------------------------------------------------------------------------------------------------
SlackStatusUpdate_SetSlackStatus(slackStatus) 
{
  webRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
  data := "profile={'status_text': '"slackStatus["text"]"', 'status_emoji': '"slackStatus["emoji"]"'}"
	
	webRequest.Open("POST", "https://slack.com/api/users.profile.set")
  webRequest.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
  webRequest.SetRequestHeader("Authorization", "Bearer "SlackStatusUpdate_MySlackToken)

  webRequest.Send(data)
}





;<<<<<<<<<<====================  Utility functions  ====================>>>>>>>>>>


;---------------------------------------------------------------------------------------------------------------------
; Private - Run a DOS command. This code taken from AutoHotKey website: https://autohotkey.com/docs/commands/Run.htm
;---------------------------------------------------------------------------------------------------------------------
SlackStatusUpdate_RunWaitOne(command)
{
  shell := ComObjCreate("WScript.Shell")      ; WshShell object: http://msdn.microsoft.com/en-us/library/aew9yb99
  exec := shell.Exec(ComSpec " /C " command)  ; Execute a single command via cmd.exe
  Return exec.StdOut.ReadAll()                ; Read and return the command's output 
}



;---------------------------------------------------------------------------------------------------------------------
; Run a DOS command
;---------------------------------------------------------------------------------------------------------------------
SlackStatusUpdate_RunWaitHidden(cmd)
{
	Sleep, 250                ; KUMMER TRYING THIS TO PREVENT ERRORS READING FROM CLIPBOARD
  clipSaved := ClipboardAll	; Save the entire clipboard
  Clipboard = 

	Runwait %cmd% | clip,,hide
  output := Clipboard
	
	Sleep, 250                ; KUMMER TRYING THIS TO PREVENT ERRORS READING FROM CLIPBOARD
  Clipboard := clipSaved	  ; Restore the original clipboard. Note the use of Clipboard (not ClipboardAll).
  clipSaved =			          ; Free the memory in case the clipboard was very large

	Return output
}