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
; PUBLIC - Set Slack status, based on what wifi networks are available to the user. I can be connected to a network;
;          by a wired ethernet cable in the office or at home, so looking at what wifi networks are nearby/available 
;          seems like an accurate method of determining where I'm connected from. 
;---------------------------------------------------------------------------------------------------------------------
SlackStatusUpdate_SetSlackStatusBasedOnNetwork() 
{
  ; Get current Slack status (network errors are returned as emoji "???")
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
	  ; I'm not in a meeting, so we need to set my Slack status
		done := False
		Loop
		{
			If (Dllcall("Sensapi.dll\IsNetworkAlive", "UintP", lpdwFlags)) 
			{
			  ; I'm connected to a network
			  done := True
			  If (AmNearOfficeWifiNetwork()) 
				  SlackStatusUpdate_SetSlackStatus(SlackStatusUpdate_SlackStatuses["workingInOffice"])
				Else
				  SlackStatusUpdate_SetSlackStatus(SlackStatusUpdate_SlackStatuses["workingRemotely"])
			}
			Else
			{
				; Wait for 30 seconds and check again
				Sleep, 30000
			}
		}
		Until done
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
; Private - Am I near an office wifi network? Return true if any of the available wifi networks match the regular
;           expression in SlackStatusUpdate_OfficeNetworks.
;             - Command NETSH WLAN SHOW NETWORKS to show all the available wifi networks, which has output like this:
;                 Interface name : Wi-Fi
;                 There are 13 networks currently visible.
;                 
;                 SSID 1 : xfinitywifi
;                 Network type            : Infrastructure
;                 Authentication          : Open
;                 Encryption              : None
;                 
;                 ...
;---------------------------------------------------------------------------------------------------------------------
AmNearOfficeWifiNetwork()
{
	atWork := False

	cmd = %comspec% /c netsh wlan show networks
	allNetworks := SlackStatusUpdate_RunWaitHidden(cmd)
	
	; This may not be the BEST was to do it, but it works
  officeNetworkPattern = i)%SlackStatusUpdate_OfficeNetworks%
	pos=1
  While pos := RegExMatch(allNetworks, "i)\Rssid.+?:\s(.*)\R", oneNetwork, pos+StrLen(oneNetwork)) 
	{
	  ; oneNetwork is the line like "SSID x : network_ssid", so parse out the network's SSID
	  networkSSID := RegExReplace(oneNetwork, "\R.*?:\s(\V+)\R", "$1")

	  If RegExMatch(networkSSID, officeNetworkPattern)
	    atWork := True
  }	

	Return %atWork%
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