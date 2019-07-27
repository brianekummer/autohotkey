;========================================================================================================================
; Brian Kummer's AutoHotKey script
;
;
; Near the bottom of this script are a number of #include statements to include libraries of utility functions
;
;
; Miscellaneous
; -------------
;   - Definition of AutoHotKey keys: http://www.autohotkey.com/docs/KeyList.htm
;   - Send/SendInput modifiers: ^ = Ctrl, ! = Alt, + = Shift, # = Windows
;
;
; Known Issues
; ------------
;   - None
;
;
; To Do List
; ----------
; 1. Auto-generate shortcuts in header
;    NEED TO FIX PROBLEMS WHEN CAN'T OVERWRITE FILE
;    ***** THIS CODE IS NOT TRUSTWORTHY YET *****
;    BEWARE OF IT REPLACING TOO MUCH TEXT!!!!, going down to "Upon startup of this script, perform these actions"
;
;
; Future Ideas
; ------------
;   - When on TeleBYOD network, display a pop-up somewhere telling me of that. Use a timer to run 
;     "netsh wlan show interfaces | findstr TeleBYOD", which only returns text when I'm connected to that network
;   - Any use for text-to-speech? ComObjCreate("SAPI.SpVoice").Speak("Speak this phrase")
;   - Popup menus are useful- can I use them elsewhere?
;       - ADP for entering timesheet?
;
;
; Summary (7/22/2019)
; -------------------
; * Implements zooming of font size in multiple applications
; * Add Auto-Correct to EVERY Application (http://www.biancolo.com/articles/autocorrect)
; * Windows Virtual Desktop Improvements - Enhancements around Windows virtual desktops, such as assigning specific 
;   wallpapers to specific virtual desktops.  Based on code: https://github.com/Ciantic/VirtualDesktopAccessor
; * Slack Improvements
;     - Changing "/lunch" "/status :hamburger: At lunch", etc.
;     - Upon Windows login/unlock, set Slack status based on nearby wifi networks
; * Backup KeePass Database - Upon Windows login/unlock, if a specific USB drive is plugged in, backup KeePass 
;   database to it
;
;
; Shortcuts - Auto-Generated
; ------------------------------
; ^+WheelDown     Chrome (AHK)        Scroll through open tabs
; ^+WheelUp       Chrome (AHK)        Scroll through open tabs
;
; /lunch          Slack (AHK)         Set status to "/status :hamburger: At lunch"
; /mtg            Slack (AHK)         Set status to "/status :spiral_calendar_pad: In a meeting"
; /status         Slack (AHK)         Clear status
; /wfh            Slack (AHK)         Set status to "/status :house: Working remotely"
; ^WheelDown      Slack (AHK)         Decrease font size
; ^WheelUp        Slack (AHK)         Increase font size
;
; ^WheelDown      Typora (AHK)        Decrease font size
; ^WheelUp        Typora (AHK)        Increase font size
;
; #^1             V. Desktops (AHK)   Switch to virtual desktop #1
; #^2             V. Desktops (AHK)   Switch to virtual desktop #2
; #^3             V. Desktops (AHK)   Switch to virtual desktop #3
; #^4             V. Desktops (AHK)   Switch to virtual desktop #4
; #^5             V. Desktops (AHK)   Switch to virtual desktop #5
; #^6             V. Desktops (AHK)   Switch to virtual desktop #6
; #^7             V. Desktops (AHK)   Switch to virtual desktop #7
; #^8             V. Desktops (AHK)   Switch to virtual desktop #8
; #^9             V. Desktops (AHK)   Switch to virtual desktop #9
; #^+Left         V. Desktops (AHK)   Move active window to previous virtual desktop
; #^left          V. Desktops         Switch to previous virtual desktop
; #^+Right        V. Desktops (AHK)   Move active window to next virtual desktop
; #^right         V. Desktops         Switch to next virtual desktop
;
; ~$^s            VS Code (AHK)       After save AHK file in VS Code, reload the current script
;
; (login/unlock)  Windows (AHK)       Set Slack status based on nearby wifi networks; KeePass backup
; #1              Windows             1st app in the task bar
; #2              Windows             2nd app in the task bar
; #3              Windows             3rd app in the task bar
; #+4             Windows (AHK)       Timesheet (Shift-4 is $)
; #4              Windows             4th app in the task bar
; #5              Windows             5th app in the task bar
; #6              Windows             6th app in the task bar
; #7              Windows             7th app in the task bar
; #8              Windows             8th app in the task bar
; #9              Windows             9th app in the task bar
; #a              Windows             Windows Action Center
; #b              Windows (AHK)       git Bash (as admin)
; #c              Windows (AHK)       Outlook Calendar
; #down           Windows (AHK)       Minimize active window instead of making it unmaximized, then minimize
; #d              Windows             Windows desktop
; #e              Windows             Windows Explorer
; #g              Windows (AHK)       gTasks Pro
; #i              Windows (AHK)       Outlook Inbox
; #^j             Windows (AHK)       JIRA, smart (open a specific story)
; #j              Windows (AHK)       JIRA, current board
; #^k             Windows (AHK)       Slack, "jump to" dialog
; #k              Windows (AHK)       Slack
; #l              Windows             Lock workstation
; #m              Windows (AHK)       Music/media app
; #NumpadSub      Windows (AHK)       TEMP - price checks
; +printscreen    Windows             Take screenshot of the whole screen
; printscreen     Windows (AHK)       Start Windows snipping tool in rectangle mode
; #^p             Windows (AHK)       Personal cloud, smart (on 2nd virtual desktop in Edge)
; #p              Windows             Project (duplicate, extend, etc)
; #Space          Windows (AHK)       Toggle dark mode for active application
; #s              Windows (AHK)       Source code (BitBucket)
; #t              Windows (AHK)       Typora, with each virtual desktop having a different folder of files
; #up             Windows             Maximize active window
; #^v             Windows (AHK)       Visual Studio Code, smart (paste the selected text into the newly opened window)
; #v              Windows (AHK)       Visual Studio Code
; #v              Windows             Clipboard
; +WheelDown      Windows (AHK)       Turn system volume down
; +WheelUp        Windows (AHK)       Turn system volume up
; #w              Windows (AHK)       Wiki- Confluence
; XButton1        Windows (AHK)       Minimize current application
; XButton2        Windows (AHK)       Minimize app or close window/tab or close app
; #z              Windows (AHK)       noiZe
;========================================================================================================================









#NoEnv
#Persistent
#WinActivateForce


;---------------------------------------------------------------------------------------------------------------------
; Global variables - configured from Windows environment variables
;---------------------------------------------------------------------------------------------------------------------
Global WindowsLocalAppDataFolder
Global WindowsProgramFilesX86Folder
Global WindowsProgramFilesFolder
Global WindowsUserName
Global WindowsDnsDomain
Global WindowsUserProfile
Global UserEmailAddress

; These come from Windows-defined environment variables
EnvGet, WindowsLocalAppDataFolder, LOCALAPPDATA
EnvGet, WindowsProgramFilesX86Folder, PROGRAMFILES(X86)
EnvGet, WindowsProgramFilesFolder, PROGRAMFILES
EnvGet, WindowsUserName, USERNAME
EnvGet, WindowsDnsDomain, USERDNSDOMAIN
EnvGet, WindowsUserProfile, USERPROFILE
UserEmailAddress = %WindowsUserName%@%WindowsDnsDomain%

; These come from my own Windows environment variables. See "My Automations Config.bat" for details
Global BackupDriveSerialNumber
Global KeePassDBFilename
Global KeePassDBBackupFilename
EnvGet, JiraUrl, AHK_URL_JIRA
EnvGet, JiraMyProjectKeys, AHK_MY_PROJECT_KEYS_JIRA
EnvGet, JiraDefaultProjectKey, AHK_DEFAULT_PROJECT_KEY_JIRA
EnvGet, JiraDefaultRapidKey, AHK_DEFAULT_RAPID_KEY_JIRA
EnvGet, JiraDefaultSprint, AHK_DEFAULT_SPRINT_JIRA
EnvGet, SourceCodeUrl, AHK_URL_SOURCE_CODE
EnvGet, TimesheetUrl, AHK_URL_TIMESHEET
EnvGet, WikiUrl, AHK_URL_WIKI
EnvGet, PersonalCloudUrl, AHK_URL_PERSONAL_CLOUD
EnvGet, NoiseBrownMP3, AHK_MP3_NOISE_BROWN
EnvGet, NoiseRailroadMP3, AHK_MP3_NOISE_RAILROAD
EnvGet, BackupDriveSerialNumber, AHK_BACKUP_DRIVE_SERIAL_NUMBER
EnvGet, KeePassDBFilename, AHK_KEEPASS_DB_FILENAME
EnvGet, KeePassDBBackupFilename, AHK_KEEPASS_DB_BACKUP_FILENAME

; Commonly used folders	
Global MyDocumentsFolder
Global MyPersonalFolder
MyDocumentsFolder = %WindowsUserProfile%\Documents\
MyPersonalFolder = %WindowsUserProfile%\Personal\



;---------------------------------------------------------------------------------------------------------------------
; Upon startup of this script, perform these actions
;---------------------------------------------------------------------------------------------------------------------
  SetTitleMatchMode RegEx            ; Make windowing commands use regex
  OnExit("ExitFunc")                 ; Upon exit of this script, execute this function

  RunAsAdmin()
  StartInterceptingWindowsUnlock()

  ; Configure wallpapers for Windows virtual desktops. *REQUIRES* Windows environment variables 
	; AHK_VIRTUAL_DESKTOP_WALLPAPER_x such as:
	;   AHK_VIRTUAL_DESKTOP_WALLPAPER_1 = Main    |C:\Pictures\wallpaper1.jpg
  ;   AHK_VIRTUAL_DESKTOP_WALLPAPER_2 = Personal|C:\Pictures\wallpaper2.jpg
	; See "Virtual Desktop Accessor.ahk" for details
	VirtualDesktopAccessor_Initialize()

	; Configure Slack status updates based on the network. *REQUIRES* several Windows environment variables - see 
	; "Slack Status Update.ahk" for details
  SlackStatusUpdate_Initialize()
  SlackStatusUpdate_SetSlackStatusBasedOnNetwork()

	; Build the popup menu for starting a music/media app
  BuildMediaPlayerMenu()

  Return

	

;---------------------------------------------------------------------------------------------------------------------
; Upon exit of this script, perform these actions
;---------------------------------------------------------------------------------------------------------------------
ExitFunc(ExitReason, ExitCode)
{
  StopInterceptingWindowsUnlock()
}



;---------------------------------------------------------------------------------------------------------------------
; Hook into Windows system message WM_WTSSESSION_CHANGE so that when I log into the computer, I can do stuff
;   - Code taken from http://autohotkey.com/board/topic/23095-trouble-detecting-windows-unlock-only-when-compiled
;---------------------------------------------------------------------------------------------------------------------
StartInterceptingWindowsUnlock() 
{
  WM_WTSSESSION_CHANGE := 0x02B1
	NOTIFY_FOR_ALL_SESSIONS := 0
  OnMessage(WM_WTSSESSION_CHANGE, "OnWindowsUnlock")

  hw_ahk := FindWindowEx(0, 0, "AutoHotkey", a_ScriptFullPath " - AutoHotkey v" a_AhkVersion)
  result := DllCall("Wtsapi32.dll\WTSRegisterSessionNotification", "uint", hw_ahk, "uint", NOTIFY_FOR_ALL_SESSIONS)
}
StopInterceptingWindowsUnlock() 
{
  DllCall("Wtsapi32.dll\WTSUnRegisterSessionNotification", "uint", hwnd)
}
; (login/unlock)::     ; Windows|AHK|Set Slack status based on nearby wifi networks; KeePass backup
OnWindowsUnlock(wParam, lParam)
{
  WTS_SESSION_UNLOCK := 0x8
  If (wParam = WTS_SESSION_UNLOCK)
	{
		SendInput #a                                  		  ; Open Windows Action Center to show new notifications from my phone
    SlackStatusUpdate_SetSlackStatusBasedOnNetwork()		; If appropriate, then update my Slack status
	  BackupKeePass()                                    	; If backup drive is inserted, then backup my KeePass database
	}
}





;---------------------------------------------------------------------------------------------------------------------
; Temporary/experimental stuff goes here
;---------------------------------------------------------------------------------------------------------------------

; Win+Ctrl+V      to open VPN app
;^#v::
;  Run, "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Palo Alto Networks\GlobalProtect\GlobalProtect.lnk"
;  WinWait, GlobalProtect,, 10
;	Sleep, 500
;	WinActivate, GlobalProtect
;	Return

#NumpadSub::     ; Windows|AHK|TEMP - price checks
	; Dr Seuss's The Grinch (Illumination)
	Run, "https://www.amazon.com/Illumination-Presents-Dr-Seuss-Grinch/dp/B07JYR54B7/ref=sr_1_1?ie=UTF8&qid=1550496790&sr=8-1&keywords=dvd+illumination+grinch",, Max
	Run, "https://www.walmart.com/ip/Illumination-Presents-Dr-Seuss-The-Grinch-DVD/577298400",, Max
	Run, "https://www.bestbuy.com/site/illumination-presents-dr-seuss-the-grinch-dvd-2018/6310541.p?skuId=6310541",, Max

	; Sabrina the Teenage Witch - 1996 DVD - movie that started the series
	Run, "https://www.amazon.com/s/ref=nb_sb_noss?url=search-alias`%3Daps&field-keywords=sabrina+the+teenage+witch+dvd+1996+-season"
	Run, "https://www.ebay.com/sch/i.html?_from=R40&_trksid=m570.l1313&_nkw=sabrina+the+teenage+witch+dvd+-season&_sacat=0&LH_TitleDesc=0&_osacat=0&_odkw=sabrina+the+teenage+witch+dvd+1996+-season&LH_TitleDesc=0"
	
	; Pirates of Caribbean movies - low priority, have these on DVR
	;   - Already Own
	;       2017- Dead Men Tell No Tales
	;   - Need to Buy
	;       - ~$38 on Amazon or Target for each separately
	;       - If buy used through Amazon, looks like everyone charges shipping on each DVD, so can buy DVD for $1 + $4 in shipping :-(
	;       2003- Curse of the Black Pearl
	;       2006- Dead Man's Chest
	;       2007- At World's End
	;       2011- On Stranger Tides
	Run, "https://www.amazon.com/s/ref=nb_sb_ss_i_1_12?url=search-alias`%3Dmovies-tv&field-keywords=dvd+pirates+of+the+caribbean&sprefix=dvd+pirates+`%2Cmovies-tv`%2C131&crid=2CVU55IXFKKPA",, Max
	Run, "https://www.target.com/s?searchTerm=dvd+pirates+of+caribbean",, Max
	Run, "https://www.bestbuy.com/site/searchpage.jsp?id=pcat17071&st=pirates+of+the+caribbean+dvd",, Max
	
  ; Definitely
  ;   - Private Eyes (Jason Priestley)
	; Not sure 
	;   - Pinky & the Brain, Pinky, Elmyra & the Brain
  ;   - Animaniacs
	;   - Tiny Toons

	Return



;---------------------------------------------------------------------------------------------------------------------
; Backup KeePass database
;   - If the flash drive I use to backup my KeePass database is inserted, then back it up to that drive
;---------------------------------------------------------------------------------------------------------------------
BackupKeePass()
{
	keePassBackupDriveLetter := GetDriveLetter(BackupDriveSerialNumber, "REMOVABLE")
	If (keePassBackupDriveLetter)
	{
	  ; FileCopy - last parameter determines if overwrite
	  ; MsgBox options: 64 (Info icon) + 4096 (System Modal- always on top)
		FileCopy, %KeePassDBFilename%, %keePassBackupDriveLetter%:\%KeePassDBBackupFilename%, 1
		MsgBox, 4160, KeePass Database Backup, %KeePassDBFilename%`nhas been backed up to %keePassBackupDriveLetter%:\%KeePassDBBackupFilename%.
	}
}



;---------------------------------------------------------------------------------------------------------------------
; Dark mode
;---------------------------------------------------------------------------------------------------------------------
#Space::     ; Windows|AHK|Toggle dark mode for active application
	If WinActive("ahk_exe chrome.exe")
	{
		; In Chrome, need Dark Reader extension, which has an existing shortcut Alt+Shift+D that toggles between dark/light
	  SendInput !+D   
	}
	Else If WinActive("- KeePass")
	{
	  ; In KeePass, need KeeTheme plugin (https://github.com/xatupal/KeeTheme), which has an existing shortcut Ctrl+T 
		; that toggles Dark Theme on and off.
	  SendInput ^t   
	}
	Else If WinActive("- Visual Studio Code")
	{
		; In Visual Studio Code, there is no easy way to see which color theme is active, so I look at the color of a
		; specific pixel in the window, and if it's dim (blue < 55) then I ASSUME we're displaying a dark theme, else I
		; assume we're using a light theme.
		PixelGetColor, color, 20, 70
		blue:="0x" SubStr(color,3,2)  ; substr is to get the piece
		blue:=blue+0                  ; add 0 to convert it to decimal
		If blue < 55
		{
	    SendInput ^k
			Sleep, 100
			SendInput ^t
			Sleep, 100
			SendInput Light+{ENTER}
		}
		Else
		{
	    SendInput ^k
			Sleep, 100
			SendInput ^t
			Sleep, 100
			SendInput Dark+{ENTER}
		}
	}
	Else If WinActive("- IntelliJ IDEA")
	{
		; In IntelliJ IDEA, there is no easy way to see which color theme is active, so I look at the color of a specific 
		; pixel in the window, and if it's dim (blue < 55) then I ASSUME we're displaying a dark theme, else I assume 
		; we're using a light theme.
		PixelGetColor, color, 17, 80
		blue:="0x" SubStr(color,3,2)  ; substr is to get the piece
		blue:=blue+0                  ; add 0 to convert it to decimal
		SendInput ^!s
		Sleep, 500
		SendInput theme
		Sleep, 250
		SendInput {ENTER}
		Sleep, 250
		SendInput {DOWN}
		Sleep, 250
		If blue < 75
			SendInput {DOWN}{DOWN}      ; Switching to IntelliJ/Default/Light
		Else
			SendInput {UP}{UP}          ; Switching to Darcula
		SendInput {ENTER}
		Sleep, 250
		SendInput {ENTER}
	}
	Else If WinActive("ahk_class Notepad++")
  {
	  ; In Notepad++, need VS2015-Dark theme (https://github.com/Dark-Genser/NotepadPP_AHK_VS-theme) and can then go to
		; Settings => Style Configurator to toggle between themes "Default" and "VS2015-Dark"
		SendInput {ALT}ts{ENTER}
		WinWaitActive, Style Configurator,, 2
		ControlGet, currentTheme, Choice,, ComboBox1, Style Configurator
		ControlFocus, ComboBox1, Style Configurator
		If InStr(currentTheme, "Dark")
		  SendInput d    ; Select "Default"
		Else
		  SendInput vvv  ; Select "VS2015-Dark"
		SendInput {TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{ENTER}
  }	
	Else If WinActive("- Eclipse IDE")
	{
	  ; In Eclipse settings Preferences => Editor => Keys, I MANUALLY added a shortcut for Ctrl+Shift+8 to open 
		; Preferences => General => Appearance. So this code can use that shortcut to programmatically open
		; Preferences => Appearance and then toggle the Theme between "Classic" and "Dark".
		SendInput ^+{8}     
		WinWaitActive, Preferences,, 2
		ControlGet, currentTheme, Choice,, ComboBox1, Preferences
		ControlFocus, ComboBox1, Preferences
		If InStr(currentTheme, "Dark")
			SendInput c        ; Select "Classic"
		Else
			SendInput d        ; Select "Dark"
    SendInput !a         ; Apply changes
    SendInput {Escape}   ; Close dialog saying restart necessary for full effect
    SendInput {Escape}   ; Close Preferences dialog
	}
	Else If WinActive("Microsoft Visual Studio")
	{
	  ; In Visual Studio, Tools => Options, toggle between "Blue" and "Dark"
    SendInput !to    
		WinWaitActive, Options,, 2
		SendInput ^eVisual experience{TAB}
		Sleep, 500
		ControlGet, currentTheme, Choice,, ComboBox1, Options
		ControlFocus, ComboBox1, Options
		If InStr(currentTheme, "Dark")
			SendInput b     ; Select "Blue"
		Else
		  SendInput d     ; Select "Dark"
		SendInput {ENTER}
	}
	Else If WinActive("- Typora")
	{
		; In Typora, there is no easy way to see which theme is active:
		;   - Typora has a custom UI that is not accessible by AHK ControlGet/etc commands
		;   - I could not find where the name of the current theme is persisted, either in a file in
		;     C:\Program Files\Typora, or in the registry
		;   - While Typora uses Chrome to render, and the developer tools (Shift+F12) shows me see
		;     this in the header: 
		;       <link rel="stylesheet" href="C:\Users\Brian-Kummer\AppData\Roaming\Typora\themes\night.css" id="theme_css">
		;     I can't figure out how to access that from AHK
		; So I look at the color of a specific pixel in the window, and if it's dim (blue < 55) then
		; I ASSUME we're displaying a dark theme, else I assume we're using a light theme.
		PixelGetColor, color, 20, 70
		blue:="0x" SubStr(color,3,2)  ; substr is to get the piece
		blue:=blue+0                  ; add 0 to convert it to decimal
		SendInput !t
		Sleep, 100
		If blue < 55
			SendInput p                 ; Switch to light theme Pixyll
		Else
			SendInput nn{ENTER}         ; Switch to dark theme Night
	}
	Else If WinActive("ahk_exe explorer.exe")
	{
	  ; In Windows Explorer, light and dark mode is controlled by a registry key, and
		;   0 = dark mode
		;   1 = light mode
		RegRead, appsUseLightTheme, HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize, AppsUseLightTheme
		appsUseLightTheme:=!appsUseLightTheme
		RegWrite, REG_DWORD, HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize, AppsUseLightTheme, %appsUseLightTheme%
	}
	Else If WinActive("ahk_class VirtualConsoleClass")
	{
	  ; ConEmu Settings => Features => Colors, toggle between "Tomorrow Night Blue" (light) and "Cobalt2" (dark)
		SendInput #!p
		WinWaitActive, Settings.*,, 2
		SendInput ^fSchemes{ENTER}
		Sleep, 500
		ControlGetText, currentTheme, Edit20, Settings.*
		ControlFocus, Edit20, Settings.*
		If InStr(currentTheme, "Cobalt2")
			SendInput <Tomorrow Night Blue>{DOWN}
		Else
		  SendInput <Cobalt2>{DOWN}
		Sleep, 500
		ControlFocus, Save settings, Settings.*
		SendInput, {ENTER}
	}
	Else If WinActive("ahk_exe OUTLOOK.EXE")
	{
		; In Microsoft Outlook, there is no easy way to see which color theme is active, so I look at the color of a
		; specific pixel in the window, and if it's dim (blue < 55) then I ASSUME we're displaying a dark theme, else
		; I assume we're using a light theme.
		SendInput !f
		Sleep, 100
		SendInput d
		Sleep, 500
		PixelGetColor, color, 200, 200
		blue:="0x" SubStr(color,3,2)   ; substr is to get the piece
		blue:=blue+0                   ; add 0 to convert it to decimal
		SendInput y1
		If blue < 55
			SendInput c                  ; Switch to light theme Colorful
		Else
			SendInput b                  ; Switch to dark theme Black
		SendInput {ENTER}{ESCAPE}
	}
  Return
	
	
	
;---------------------------------------------------------------------------------------------------------------------
; Chrome
;---------------------------------------------------------------------------------------------------------------------
#IfWinActive ahk_exe chrome.exe
  ^+WheelUp::     ; Chrome|AHK|Scroll through open tabs
	  SendInput ^{PgUp}
		Return
  ^+WheelDown::   ; Chrome|AHK|Scroll through open tabs
	  SendInput ^{PgDn}
		Return
#IfWinActive



;---------------------------------------------------------------------------------------------------------------------
; Slack
;   - Ctrl+MouseWheel     Zoom in and out
;   - #k                  Open Slack
;   - #^k                 Open Slack and go to the "Jump to" window
;
; Functionality implemented in Slack-Status-Update.ahk
;   - /xxxxx              When in Slack, replace "/lunch", "/wfh", "/mtg" with the equivalent Slack command to change
;                         my status
;   - upon login/unlock   Upon Windows login/unlock, set Slack status based on nearby wifi networks
;---------------------------------------------------------------------------------------------------------------------
#IfWinActive ahk_group SlackStatusUpdate_WindowTitles
  ^WheelUp::     ; Slack|AHK|Increase font size
    SendInput ^{+}
		Return
  ^WheelDown::   ; Slack|AHK|Decrease font size
	  SendInput ^{-}
		Return
#IfWinActive

#k::             ; Windows|AHK|Slack
  OpenSlack()
  Return

#^k::            ; Windows|AHK|Slack, "jump to" dialog
  OpenSlack()
	SendInput ^k
  Return

OpenSlack()
{
  If Not WinExist("ahk_group SlackStatusUpdate_WindowTitles")
  {
    Run, "%WindowsLocalAppDataFolder%\slack\slack.exe"
	  WinWaitActive, Slack,, 2
  }
 	WinActivate, ahk_group SlackStatusUpdate_WindowTitles
	WinMaximize, A
  Return
}

	

;---------------------------------------------------------------------------------------------------------------------
; Adjust system volume
;   - Other keystroke combinations (Win+mousewheel, or Ctrl+mousewheel) affected the current window
;---------------------------------------------------------------------------------------------------------------------
+WheelUp::     ; Windows|AHK|Turn system volume up
	SendInput {Volume_Up 1}
	Return
+WheelDown::   ; Windows|AHK|Turn system volume down
  SendInput {Volume_Down 1}
	Return



;---------------------------------------------------------------------------------------------------------------------
; Extra mouse buttons
;   - XButton1 (front button) minimizes the current window
;   - XButton2 (rear button) depending on the active window, closes the active TAB/WINDOW, or minimizes or closes the 
;     active APPLICATION.
;
; This is based on the name of app/process, NOT the window title, or else it would minimize a browser with a tab whose
; title is "How to Use Slack".  Also, Microsoft Edge browser is more complex than a single process, so detecting it is
; more complex.
;---------------------------------------------------------------------------------------------------------------------
XButton1::     ; Windows|AHK|Minimize current application
  WinMinimize, A
	Return

XButton2::     ; Windows|AHK|Minimize app or close window/tab or close app
  WinGet, processName, ProcessName, A
  SplitPath, processName,,,, processNameNoExtension

	If RegExMatch(processNameNoExtension, "i)skype|outlook|wmplayer|slack|typora") 
	  or WinActive("gTasks Pro ahk_exe ApplicationFrameHost.exe") 
	  or WinActive("iHeartRadio ahk_exe ApplicationFrameHost.exe")
    or WinActive("ahk_exe Google Play Music Desktop Player.exe")
	{
		WinMinimize, A     ; Do not want to close these apps
	}
  Else If RegExMatch(processNameNoExtension, "i)chrome|iexplore|firefox|notepad++|ssms|devenv|eclipse|winmergeu|robo3t|code|idea64") 
	  or WinActive("Microsoft Edge ahk_exe ApplicationFrameHost.exe")
	{
    SendInput ^{f4}    ; Close a WINDOW/TAB/DOCUMENT
	}
  Else
	{
    SendInput !{f4}    ; Close the APP
	}
  Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Visual Studio Code
;
; TODO: Transition away from Notepad++ to VS Code
; 
;   - When editing any AutoHotKey script in Visual Studio Code, clicking Ctrl-S to save the script also causes 
;     AutoHotKey to reload the current script. This eliminates the need to right-click the AutoHotKey system tray icon
;     and select "Reload This Script".
;   - #v       Open Visual Studio Code
;   - #^v      Open Visual Studio Code and paste the selected text into the new Visual Studio Code window
;
; TODO:
;                   - I could have copied the selected text into a variable and done a "Send %variableName%", which 
;                     would send each character as if it was typed, which would be very SLOW. Instead, I paste the 
;                     selected text into the Notepad++ window.
;                   - If the pasted text looks like a recognized type of text, format it using a Notepad++ plugin.
;                     THESE PLUGINS -= MUST =- BE INSTALLED VIA PLUGINS => PLUGIN MANAGERS => SHOW PLUGIN MANAGER
;                       Type       Plugin Name   Plugin Hot Key     Comments
;                       --------   -----------   ----------------   ------------------------------------------------------
;                       HTML/XML   XML Tools     Ctrl+Shift+Alt+b   Also sets Notepad++ language to "XML" to enable folding
;                       JSON       JSTool        Ctrl+Alt+m
;                       SQL        SQLinForm     Alt+Shift+f
;---------------------------------------------------------------------------------------------------------------------
#IfWinActive .ahk - Visual Studio Code 
~$^s::      ; VS Code|AHK|After save AHK file in VS Code, reload the current script
  ;Reload
	AutoGenerateDocumentation(A_ScriptName)
  Reload
  Return
#IfWinActive

#v::        ; Windows|AHK|Visual Studio Code
	;Run "%WindowsProgramFilesFolder%\Notepad++\notepad++.exe"
	Run "C:\Users\Brian-Kummer\AppData\Local\Programs\Microsoft VS Code\Code.exe"
  Return

#^v::       ; Windows|AHK|Visual Studio Code, smart (paste the selected text into the newly opened window)
  ; Define the regular expression patterns to recognize different types of text
  ;   - Identifying SQL is complex. As a rough guess, look for any one of the following:
  ;       CREATE|ALTER...FUNCTION|PROCEDURE|VIEW|INDEX...AS BEGIN
  ;       DROP...FUNCTION|PROCEDURE|VIEW|INDEX
  ;       SELECT...FROM
  regExHtmlOrXml = s)<.*>.*/.*>
  regExJson = s)^\s*\[?\{.*\:.*\,.*\}
  regExSql = is)
  regExSql = %regExSql%((\b(create|alter)\b.*\b(function|procedure|view|index)\b.*\bas\s+begin\b)|
  regExSql = %regExSql%(\bdrop\b.*\b(function|procedure|view|index)\b)|
  regExSql = %regExSql%(\bselect\b.*\bfrom\b))+

	; Save the selected text to the clipboard
  ClipSaved := ClipboardAll
  Clipboard = 
  SendInput ^c
  ClipWait, 2
	
  If (!ErrorLevel)
  {
	  If Not WinExist("- Visual Studio Code") 
		{
		  ; Visual Studio Code isn't open, so start it, and wait up to 2 seconds for it to open
		  ;Run "%WindowsProgramFilesFolder%\Notepad++\notepad++.exe"
			Run "C:\Users\Brian-Kummer\AppData\Local\Programs\Microsoft VS Code\Code.exe"
		  WinWaitActive, Visual Studio Code,,2
    
		  WinGetTitle, notpadTitle, A
			foundPos := RegExMatch(notpadTitle, "i)Untitled-\d+ - Visual Studio Code")
		  If (foundPos == 0)
			  ; The active Visual Studio Code document is not new, so we need to create a new document to paste our text into
		    SendInput ^n
		}
		Else
		{
		  ; Visual Studio Code is already open, so activate it and create a new document
		  WinActivate, Visual Studio Code
		  SendInput ^n
		}
		
		; Paste in our text
    SendInput ^v
    Sleep, 750
		
		; TODO: Fix this with VS Code
    ; We can do some extra formatting of some types of data using Notepad++ plugins, 
    ; based on the type of the data we're viewing
    ;newText := Clipboard
    ;If RegExMatch(newText, regExHtmlOrXml)
    ;{
    ;  SendInput ^+!{b}    ; Plugins => XML Tools => Pretty print...
    ;  SendInput !{l}x     ; Language => XML
    ;}
    ;Else If RegExMatch(newText, regExJson)
    ;{
    ;  SendInput ^!{m}     ; Pluginx => JSTool =>    JSFormat
    ;}
    ;Else If RegExMatch(newText, regExSql)
    ;{
    ;  SendInput !+{f}     ; Plugins => SQLinForm => Format Selected SQL 
    ;}
  }
  Clipboard := ClipSaved	 ; Restore the original clipboard. Note the use of Clipboard (not ClipboardAll)
  ClipSaved =			       ; Free the memory in case the clipboard was very large
  Return	

	
	
;---------------------------------------------------------------------------------------------------------------------
; Personal cloud
;   - Open or activate my personal cloud website on the 2nd virtual desktop in Microsoft Edge
;---------------------------------------------------------------------------------------------------------------------
#^p::     ; Windows|AHK|Personal cloud, smart (on 2nd virtual desktop in Edge)
  If WinExist(".*Kummer Cloud ahk_exe ApplicationFrameHost.exe")
  {
	  WinActivate, Kummer Cloud
  }
	Else
	{
	  currentDesktopNumber := _GetCurrentDesktopNumber()
		If (currentDesktopNumber != 2)
      _ChangeDesktop(2) 
	  Run, "microsoft-edge:%PersonalCloudUrl%",, Max
	  WinWaitActive, Kummer Cloud,, 2
		WinMaximize, A
	}
  Return



;---------------------------------------------------------------------------------------------------------------------
; Typora
;   - Ctrl+MouseWheel   Zoom in and out
;   - #t                Open Typora if not already open. Virtual desktops "Main" and "Personal" open different folders 
;                       of files.
;---------------------------------------------------------------------------------------------------------------------
#IfWinActive - Typora 
  ^WheelUp::        ; Typora|AHK|Increase font size
	  SendInput ^+{=}
		Return
  ^WheelDown::      ; Typora|AHK|Decrease font size
	  SendInput ^+{-}
		Return
#IfWinActive

#t::                ; Windows|AHK|Typora, with each virtual desktop having a different folder of files
  ; Determine if Typora is running in this virtual desktop
	typoraIdOnThisDesktop := GetTyporaOnThisVirtualDesktop()
	If (typoraIdOnThisDesktop == 0)
	{
    n := _GetCurrentDesktopNumber()
	  If (n == 1)         ; Virtual desktop "Main"
    	Run "%WindowsProgramFilesFolder%\Typora\Typora.exe" "%WindowsUserProfile%\Documents\Always Open"
		Else If (n == 2)    ; Virtual desktop "Personal"
    	Run "%WindowsProgramFilesFolder%\Typora\Typora.exe" "%MyPersonalFolder%\Notes"
    Sleep, 1000
		
		typoraIdOnThisDesktop := GetTyporaOnThisVirtualDesktop()
	}
	
	WinActivate, ahk_id %typoraIdOnThisDesktop%
  WinMaximize, A
  Return	

GetTyporaOnThisVirtualDesktop()
{
  n := _GetCurrentDesktopNumber()
  typoraIdOnThisDesktop = 0

	WinGet, id, List, - Typora
  Loop, %id%
	{
		typoraId := id%A_Index%
	  isOnDesktop := IsWindowOnCurrentVirtualDesktop(typoraId)
	  If (isOnDesktop == 1) 
		{
		  typoraIdOnThisDesktop := typoraId
			Break
		}
	}
	Return typoraIdOnThisDesktop
}



;---------------------------------------------------------------------------------------------------------------------
; gTasks Pro
;---------------------------------------------------------------------------------------------------------------------
#g::              ; Windows|AHK|gTasks Pro
  If Not WinActive("gTasks Pro ahk_exe ApplicationFrameHost.exe") 
	{
    ; Note that Windows store apps are more complex than simple exe's
	  Run, "%MyPersonalFolder%\WindowsStoreAppLinks\gTasks Pro.lnk"
  }
	WinActivate, gTasks Pro
  Return

	

;---------------------------------------------------------------------------------------------------------------------
; Paste clipboard as plain text
;   - Code taken from https://autohotkey.com/board/topic/10412-paste-plain-text-and-copycut
;---------------------------------------------------------------------------------------------------------------------
;#^v::
;  Clip0 = %ClipBoardAll%
;  ClipBoard = %ClipBoard%      	; Convert to text
;  SendInput ^v                  ; For best compatibility: SendPlay
;  Sleep 50                      ; Don't change clipboard while it is pasted! (Sleep > 0)
;  ClipBoard = %Clip0%           ; Restore original ClipBoard
;  VarSetCapacity(Clip0, 0)     	; Free memory
;  Return

	

;---------------------------------------------------------------------------------------------------------------------
; git Bash
;   - Since ConEmu has multiple windows, cannot just use "WinActivate, Git Bash"
;---------------------------------------------------------------------------------------------------------------------
#b::              ; Windows|AHK|git Bash (as admin)
  If Not WinExist("ahk_class VirtualConsoleClass")
  {
    Run "%WindowsProgramFilesFolder%\ConEmu\ConEmu64.exe" -run {Shells::Git Bash}
    WinWait ahk_class VirtualConsoleClass	
  }
  WinActivate
  Return

	

;---------------------------------------------------------------------------------------------------------------------
; Timesheets website
;---------------------------------------------------------------------------------------------------------------------
#+4::             ; Windows|AHK|Timesheet (Shift-4 is $)
	Run, "%TimesheetUrl%",, Max
  Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Source code
;---------------------------------------------------------------------------------------------------------------------
#s::              ; Windows|AHK|Source code (BitBucket)
	Run, "%SourceCodeUrl%",, Max
  Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Wiki (Confluence)
;---------------------------------------------------------------------------------------------------------------------
#w::              ; Windows|AHK|Wiki- Confluence
	Run, "%WikiUrl%",, Max
  Return



;---------------------------------------------------------------------------------------------------------------------
; PrintScreen
;   - Open the Windows snipping tool in rectangular mode
;   - Image is stored in the clipboard
;   - Shift+PrintScreen still takes a screenshot of the whole screen
;---------------------------------------------------------------------------------------------------------------------
printscreen::     ; Windows|AHK|Start Windows snipping tool in rectangle mode
  Run snippingtool /clip
  Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Music/media
;---------------------------------------------------------------------------------------------------------------------
#m::              ; Windows|AHK|Music/media app
  If WinExist("ahk_exe Google Play Music Desktop Player.exe")
	{
	  WinActivate ahk_class Chrome_WidgetWin_1 ahk_exe i)google play music desktop player.exe
	}
  Else If WinExist("iHeartRadio ahk_exe ApplicationFrameHost.exe")
	{
	  WinActivate iHeartRadio ahk_exe ApplicationFrameHost.exe
	}
  Else If WinExist("White Noise ahk_exe ApplicationFrameHost.exe")
	{
	  WinActivate White Noise ahk_exe ApplicationFrameHost.exe
	}
  Else If WinExist("ahk_class WMPlayerApp")
	{
		WinActivate, ahk_class WMPlayerApp
	}
	Else 
	{
		Menu, MediaPlayerMenu, Show
	}
  Return  

PlayNoiseFile(whichNoiseMp3)
{
	Global WindowsProgramFilesX86Folder

  If Not WinExist("Noise.*Windows Media Player")
	{
	  ; I had a difficult time getting Windows Media Player to cooperate here, thus
		;   - /Task Library     I would have preferred to open NowPlaying
		;   - Sleep             Shouldn't need all of these, but couldn't get it to work without sleeps and WinActivate
		Run, "%WindowsProgramFilesX86Folder%\Windows Media Player\wmplayer.exe" "%whichNoiseMp3%" /Task Library
		WinWaitActive, Windows Media Player,, 2
		Sleep, 1000
	  WinActivate, Noise.*Windows Media Player
		Sleep, 2000
		WinMinimize 
	}
}

BuildMediaPlayerMenu()
{
  Global WindowsUserProfile
	Global MyPersonalFolder
	Global WindowsProgramFilesX86Folder

	Menu, MediaPlayerMenu, Add, &Google Play Music Desktop Player, MediaPlayerMenuHandler
	Menu, MediaPlayerMenu, Add, &iHeartRadio, MediaPlayerMenuHandler
	Menu, MediaPlayerMenu, Add, Windows &Media Player, MediaPlayerMenuHandler
	Menu, MediaPlayerMenu, Add
	Menu, MediaPlayerMenu, Add, &Brown Noise, MediaPlayerMenuHandler
	Menu, MediaPlayerMenu, Add, &White Noise App, MediaPlayerMenuHandler
	Menu, MediaPlayerMenu, Add, &Railroad Noise, MediaPlayerMenuHandler

	; I could not figure out how to extract an icon from a Windows Store app, AND it looks like using PNG files for icons
	; is unsupported, so I downloaded an icon from the internet and use it instead. The iHeartRadio shortcut lists this
	; as the icon: ClearChannelRadioDigital.iHeartRadio_6.0.34.0_x64__a76a11dkgb644?ms-resource://ClearChannelRadioDigital.iHeartRadio/Files/Assets/Square44x44Logo.png
	Menu, MediaPlayerMenu, Icon, &iHeartRadio, %MyPersonalFolder%\WindowsStoreAppLinks\iHeartRadio.jpg, 1, 32
	Menu, MediaPlayerMenu, Icon, &Google Play Music Desktop Player, %WindowsUserProfile%\AppData\Local\GPMDP_3\Update.exe, 1, 32
	Menu, MediaPlayerMenu, Icon, Windows &Media Player, %WindowsProgramFilesX86Folder%\Windows Media Player\wmplayer.exe, 1, 32
	Menu, MediaPlayerMenu, Icon, &Brown Noise, %MyPersonalFolder%\WindowsStoreAppLinks\Noise- Brown.jpg, 1, 32
	Menu, MediaPlayerMenu, Icon, &White Noise App, %MyPersonalFolder%\WindowsStoreAppLinks\Noise- White.jpg, 1, 32
	Menu, MediaPlayerMenu, Icon, &Railroad Noise, %MyPersonalFolder%\WindowsStoreAppLinks\Noise- Railroad.jpg, 1, 32
}

MediaPlayerMenuHandler:
  chosenMenuItem=%A_ThisMenuItem%
	If Instr(chosenMenuItem, "Windows")
	{
	  Run, "%WindowsProgramFilesX86Folder%\Windows Media Player\wmplayer.exe"
	}
	Else If Instr(chosenMenuItem, "Google")
	{
	  Run, "%WindowsUserProfile%\AppData\Local\GPMDP_3\Update.exe" --processStart "Google Play Music Desktop Player.exe"
	}
	Else If Instr(chosenMenuItem, "iHeartRadio")
	{
	  Run, "%MyPersonalFolder%\WindowsStoreAppLinks\iHeartRadio.lnk"
	}
	Else If Instr(chosenMenuItem, "White")
	{
	  Run, "%MyPersonalFolder%\WindowsStoreAppLinks\White Noise.lnk"
	}
	Else If Instr(chosenMenuItem, "Brown")
	{
	  PlayNoiseFile(NoiseBrownMP3)
	}
	Else If Instr(chosenMenuItem, "Railroad")
	{
	  PlayNoiseFile(NoiseRailroadMP3)
	}
	Return
		
	

;---------------------------------------------------------------------------------------------------------------------
; JIRA
;   - #j opens the current board
;   - #^j tries to determine which story to open
;       * If the highlighted text looks like a JIRA story number (e.g. PROJECT-1234), then open that story
;       * If the Git Bash window has text that looks like a JIRA story number, then open that story
;       * Last resort is to open the current board
;---------------------------------------------------------------------------------------------------------------------
#j::              ; Windows|AHK|JIRA, current board
	pathDefaultBoard = /secure/RapidBoard.jspa?rapidView=%JiraDefaultRapidKey%&projectKey=%JiraDefaultProjectKey%&sprint=%JiraDefaultSprint%
	Run, %JiraUrl%%pathDefaultBoard%
	Return

#^j::             ; Windows|AHK|JIRA, smart (open a specific story)
  regexStoryNumberWithoutProject = \b\d{1,5}\b
  regexStoryNumberWithProject = i)\b(%JiraMyProjectKeys%)([-_ ]|( - ))?\d{1,5}\b
	pathBrowse = /browse/
	pathDefaultBoard = /secure/RapidBoard.jspa?rapidView=%JiraDefaultRapidKey%&projectKey=%JiraDefaultProjectKey%&sprint=%JiraDefaultSprint%

  selectedText := GetSelectedTextUsingClipboard()

	; Search the selected text for something like PROJECT-1234
  RegExMatch(selectedText, regexStoryNumberWithProject, storyNumber)

  If StrLen(storyNumber) = 0
  { 
	  ; Search for just a number, and if found, add the default project name
    RegExMatch(selectedText, regexStoryNumberWithoutProject, storyNumber)
		If StrLen(storyNumber) > 0
		{
		  storyNumber = %JiraDefaultProjectKey%-%storyNumber%
		}
  }  

  If StrLen(storyNumber) = 0
  { 
	  ; Search for a ConEmu terminal with a JIRA story number
		WinGetTitle, git_window_title, ahk_exe i)\\conemu64\.exe$ ahk_class VirtualConsoleClass
		RegExMatch(git_window_title, regexStoryNumberWithProject, storyNumber)
	}

  If StrLen(storyNumber) = 0
  { 
	  ; Search for a Mintty terminal (comes with Git) with a JIRA story number
		WinGetTitle, git_window_title, ahk_exe i)\\mintty\.exe$ ahk_class mintty
		RegExMatch(git_window_title, regexStoryNumberWithProject, storyNumber)
  }  

	If StrLen(storyNumber) = 0
  {
	  ; Could not find any JIRA story number, go to a default JIRA board
		Run, %JiraUrl%%pathDefaultBoard%
  }
  Else
  {
	  ; Handle if there is an underscore or space instead of a hyphen, or no hyphen
	  storyNumber := RegExReplace(storyNumber, "[\s_]", "")
		If Not RegExMatch(storyNumber, "-")
		{
		  storyNumber := RegExReplace(storyNumber, "(\d+)", "-$1")
		}
		
    Run, %JiraUrl%%pathBrowse%%storyNumber%
  }
  Return
	
	
	
;---------------------------------------------------------------------------------------------------------------------
; Inbox
;---------------------------------------------------------------------------------------------------------------------
#i::              ; Windows|AHK|Outlook Inbox
  ActivateOrStartMicrosoftOutlook()
  WinActivate
	SendInput ^+I   ; Ctrl-Shift-I is "Switch to Inbox" shortcut key
	Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Calendar
;---------------------------------------------------------------------------------------------------------------------
#c::              ; Windows|AHK|Outlook Calendar
  ActivateOrStartMicrosoftOutlook()
  SendInput ^2	
  Return


	
;---------------------------------------------------------------------------------------------------------------------
; NOTE that Outlook is whiny, and the "instant search" feature (press Ctrl+E to search your mail items) refuses to run
; when Outlook is run as an administrator. Because we are running this AHK script as an administrator, we cannot
; simply run Outlook. Instead, we must run it as a standard user.
;---------------------------------------------------------------------------------------------------------------------
ActivateOrStartMicrosoftOutlook()
{
  global UserEmailAddress
	global WindowsProgramFilesX86Folder
	
  outlookTitle = i)%UserEmailAddress%\s-\sOutlook
  If Not WinExist(outlookTitle)	
  {
    outlookExe = %WindowsProgramFilesX86Folder%\Microsoft Office\root\Office16\OUTLOOK.EXE
	  ShellRun(outlookExe)
	  WinWaitActive, outlookTitle,,5
  }
  WinActivate
}	



;---------------------------------------------------------------------------------------------------------------------
; Minimize the active window. This overrides the existing Windows hotkey that:
;   - First time you use it, un-maximizes (restores) the window
;   - Second second time you use it, it minimizes the window
;---------------------------------------------------------------------------------------------------------------------
#down::           ; Windows|AHK|Minimize active window instead of making it unmaximized, then minimize
  WinMinimize, A
  Return


	
;---------------------------------------------------------------------------------------------------------------------
; Noise
;---------------------------------------------------------------------------------------------------------------------
#z::              ; Windows|AHK|noiZe
	PlayNoiseFile(NoiseBrownMP3)
	Return



	
; AutoGenerateDocumentation() 
; {
;   ; Read this script and build a list of shortcut keys
; 	Sleep, 1000     ; Give VS Code time to save the file before we start
; 	shortcut_details := []
; 	shortcut_array_keys := []
	
; 	line_num = 1
; 	Loop
; 	{
;     FileReadLine, line, %A_ScriptName%, %A_Index%
;     if ErrorLevel
;       break
; 		; When search for lines in this file, have to explicitly exclude the next line
; 		If (RegExMatch(line, ".*;?.*::.*;.*[\|].*[\|]") And InStr(line, "RegExMatch") == 0)
; 		{
; 			;MsgBox, 4, , Line #%A_Index% MATCHES: %line%. Continue?
;     	;IfMsgBox, No
;       ;  break
; 			current_shortcut_details := {array_key: "", raw_text: "", shortcut_keys: "", group_name: "", is_ahk: "", comments: ""}

; 			; Parse out shortcut keys, the group name, and the comments
; 			; TODO: Exclude commented code: If line starts with ; and is_ahk then is commented code and skip it
; 			RegExMatch(line, "[;]?(.*)`:`:.*\;(.*?)\|(.*)\|(.*)", info)
; 			shortcut_keys_all = %info1%
; 			group_name = %info2%
; 			is_ahk = %info3%
; 			comments = %info4%

; 			; For sorting, want keys to look like this for nice sorting purposes, keeping all varieties of
; 			; shortcut keys for a single letter together
; 			;   Windows|j|#^j|JIRA
; 			;   Windows|j|#j|JIRA
; 			shortcut_keys_stripped := RegExReplace(shortcut_keys_all, "[#^+!]")
; 			array_key = %group_name%|%shortcut_keys_stripped%|%shortcut_keys_all%|%line_num%
; 			current_shortcut_details.array_key := array_key
; 			shortcut_array_keys.Push(array_key)

;       ; Save the details of this shortcut key
; 			current_shortcut_details.raw_text := line
; 			current_shortcut_details.shortcut_keys := shortcut_keys_all
; 			current_shortcut_details.group_name := group_name
; 			current_shortcut_details.is_ahk := is_ahk
; 			current_shortcut_details.comments := comments
; 			shortcut_details.Push(current_shortcut_details)

; 			current_shortcut_details := ""  ; Release the object
; 			line_num++
; 		}
; 	}

;   shortcut_array_keys := sortArray(shortcut_array_keys)

; 	; Build the output. Use shortcut_array_keys to get shortcut keys in the desired order.
; 	output := ""
; 	last_group := ""
; 	for index, element in shortcut_array_keys
; 	{
; 		details_index := StrSplit(element, "|")[4]
;     details_object := shortcut_details[details_index]

; 		shortcut_keys := details_object.shortcut_keys
; 		group_name := details_object.group_name
; 		comments := details_object.comments

; 		padded_shortcut_keys := SubStr(shortcut_keys . "               ", 1, 15)
; 		padded_group_name := group_name
; 		If (details_object.is_ahk)
; 		  padded_group_name = %padded_group_name% (AHK)
; 	  padded_group_name := SubStr(padded_group_name . "                   ", 1, 19)
		
; 		If (last_group != "" And last_group != group_name)
; 		  output = %output%;`n
;  		output = %output%; %padded_shortcut_keys% %padded_group_name% %comments%`n

; 		last_group := group_name
; 	}
; 	FileDelete, C:\Temp\Out.txt
; 	FileAppend, %output%, C:\Temp\Out.txt

;   ; Backup our script
; 	source_filename = %A_ScriptName%
; 	FormatTime, CurrentDateTime,, yyyy-MM-ddThh-mm-ss
; 	target_filename = %A_ScriptName%.%CurrentDateTime%.ahk
;   FileCopy, %source_filename%, %target_filename%

;   ; Replace the "Shortcuts" documentation with our newly generated documentation
; 	FileRead, my_script, %A_ScriptName%
; 	title_dashes := Replicate("-", 10)
; 	end_of_comments_dashes := Replicate("=", 120)
; 	new_output := "; Shortcuts (Auto-Generated)`r`n; " . title_dashes . "`r`n" . output . ";" . end_of_comments_dashes . "`r`n"
;   ;test := RegExMatch(my_script, "s); Shortcuts`r`n; -----+`r`n;..*;")
;   ;msgbox test is %test%
; 	my_script := RegExReplace(my_script, "s); Shortcuts (Auto-Generated)`r`n; -----+`r`n;..*?;=====+", new_output)
; 	FileDelete, %source_filename%
; 	FileAppend, %my_script%, %source_filename%
; 	;Sleep, 1000

; 	MsgBox, Done regenerating shortcut documentation.
; 	Return
; }

; Replicate( Str, Count ) { ; By SKAN / CD: 01-July-2017 | goo.gl/U84K7J
;   Return StrReplace( Format( "{:0" Count "}", "" ), 0, Str )
; }

; sortArray(arr) {
; ; Code from https://autohotkey.com/board/topic/93570-sortarray/
; 	if	!IsObject(arr)
; 		return	0
; 	new :=	[]
; 	For each, item in arr
; 		list .=	item "`n"
; 	list :=	Trim(list,"`n")
; 	Sort, list, %options%
; 	Loop, parse, list, `n, `r
; 		new.Insert(A_LoopField)
; 	Return new
; }


;---------------------------------------------------------------------------------------------------------------------
; All of the following commented-out hotkeys, etc are here so that the auto-generated documentation includes them
;---------------------------------------------------------------------------------------------------------------------
#Include %A_ScriptDir%\My Automations Utilities.ahk
#Include %A_ScriptDir%\AutoGenerate Documentation.ahk
#Include %A_ScriptDir%\Slack Status Update.ahk
; /lunch::          ; Slack      |AHK|Set status to "/status :hamburger: At lunch"
; /mtg::            ; Slack      |AHK|Set status to "/status :spiral_calendar_pad: In a meeting"
; /wfh::            ; Slack      |AHK|Set status to "/status :house: Working remotely"
; /status::         ; Slack      |AHK|Clear status
#Include %A_ScriptDir%\Virtual Desktop Accessor.ahk
; #^1::             ; V. Desktops|AHK|Switch to virtual desktop #1
; #^2::             ; V. Desktops|AHK|Switch to virtual desktop #2
; #^3::             ; V. Desktops|AHK|Switch to virtual desktop #3
; #^4::             ; V. Desktops|AHK|Switch to virtual desktop #4
; #^5::             ; V. Desktops|AHK|Switch to virtual desktop #5
; #^6::             ; V. Desktops|AHK|Switch to virtual desktop #6
; #^7::             ; V. Desktops|AHK|Switch to virtual desktop #7
; #^8::             ; V. Desktops|AHK|Switch to virtual desktop #8
; #^9::             ; V. Desktops|AHK|Switch to virtual desktop #9
; #^+Left::         ; V. Desktops|AHK|Move active window to previous virtual desktop
; #^+Right::        ; V. Desktops|AHK|Move active window to next virtual desktop
; ----- Built-In Windows Shortcuts I Use -----
; #1::              ; Windows    |   |1st app in the task bar
; #2::              ; Windows    |   |2nd app in the task bar
; #3::              ; Windows    |   |3rd app in the task bar
; #4::              ; Windows    |   |4th app in the task bar
; #5::              ; Windows    |   |5th app in the task bar
; #6::              ; Windows    |   |6th app in the task bar
; #7::              ; Windows    |   |7th app in the task bar
; #8::              ; Windows    |   |8th app in the task bar
; #9::              ; Windows    |   |9th app in the task bar
; #a::              ; Windows    |   |Windows Action Center
; #d::              ; Windows    |   |Windows desktop
; #e::              ; Windows    |   |Windows Explorer
; #l::              ; Windows    |   |Lock workstation
; #p::              ; Windows    |   |Project (duplicate, extend, etc)
; #up::             ; Windows    |   |Maximize active window
; #v::              ; Windows    |   |Clipboard
; +printscreen::    ; Windows    |   |Take screenshot of the whole screen
; #^left::          ; V. Desktops|   |Switch to previous virtual desktop
; #^right::         ; V. Desktops|   |Switch to next virtual desktop


;---------------------------------------------------------------------------------------------------------------------
; Add auto-correct to EVERY application
;   http://www.biancolo.com/articles/autocorrect/
; For some reason, I have to add this at the bottom of my script, or it interferes with the other stuff in this script
;---------------------------------------------------------------------------------------------------------------------
#Include %A_ScriptDir%\Lib\AutoCorrect.ahk

;::brian::Brian      ; Doing this always capitalizes Brian-kummer@company.com
;::kummer::Kummer    ; Otherwise, it changes brian-kummer to brian-Kummer, and some apps REQUIRE a username to be all lowercase
;::i::I              ; Don't want to always capitalize i (e.g. "-i")

::1/4::¼             ; One-quarter
::1/2::½             ; One-half
::3/4::¾             ; Three-quarters
::~~::≈              ; Approximation

::nancy::Nancy
::manoj::Manoj
::doug::Doug
::bhawna::Bhawna
::andrew::Andrew
::barb::Barb
::steve::Steve
::sunil::Sunil
::prabhu::Prabhu
::shane::Shane
::shawn::Shawn
::albert::Albert
::vanita::Vanita
::rakesh::Rakesh
::oren::Oren
::sagar::Sagar
::clint::Clint
::adam::Adam
::madhan::Madhan
::tonya::Tonya
::damian::Damian
::sangeeta::Sangeeta
::vanya::Vanya
::sanath::Sanath
::kyle::Kyle
::senthil::Senthil
::mason::Mason
::tej::Tej

::telstrat::TelStrat
::Telstrat::TelStrat

::xt::XT
::cms::CMS
::iqtc::IQTC
::ocs::OCS
::wfh::WFH
;::pto::PTO
::optc::OPTC

::i'll::I'll
::i'd::I'd
::i've::I've
::i'm::I'm

::havent::haven't

::dicsuss::discuss
::disucss::discuss
::abd::and
::enetered::entered
::erro::error
::betwene::between
::aything::anything
::sugest::suggest
::wheer::where