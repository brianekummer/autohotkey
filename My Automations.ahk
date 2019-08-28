;========================================================================================================================
; Brian Kummer's AutoHotKey script
;
;
; Summary (7/22/2019)
; -------------------
; * Hotkeys to start or switch to applications and websites
; * Zooming of font size in multiple applications
; * Auto-Correct for every application (http://www.biancolo.com/articles/autocorrect)
; * Windows Virtual Desktop Improvements - Enhancements around Windows virtual desktops, such as assigning specific 
;   wallpapers to specific virtual desktops.  Based on code: https://github.com/Ciantic/VirtualDesktopAccessor
; * Slack Improvements
;     - Changing "/lunch" "/status :hamburger: At lunch", etc.
;     - Upon Windows login/unlock, set Slack status based on nearby wifi networks
; * Automatic backup of password databases when insert specific USB drive
; * Screen brightness controls for laptop and/or monitor
; * Toggling of dark mode in some applications
;
;
; Notes
; -------------
;   - Near the bottom of this script are a number of #include statements to include libraries of utility functions
;   - Definition of AutoHotKey keys: http://www.autohotkey.com/docs/KeyList.htm
;   - Send/SendInput modifiers: ^ = Ctrl, ! = Alt, + = Shift, # = Windows
;   - This looks helpful: http://www.daviddeley.com/autohotkey/xprxmp/autohotkey_expression_examples.htm
;
;
; Known Issues
; ------------
;
;
; To Do
; ------------
;   - Fix VS Code so can use it instead of Notepad++
;   - Package for Ed (simple and full-featured versions)
;   - Finish screen brightness
;       - Handle multiple monitors
;       - Try to use PowerShell and WMI instead of NirSoft utilities and ScreenBrightness AHK code
;   - What do, if anything, with gTasks Pro? Is useful, but do I need it on a hotkey?
;
;
; Future Ideas
; ------------
;   - When on TeleBYOD network, display a pop-up somewhere telling me of that. Use a timer to run 
;     "netsh wlan show interfaces | findstr TeleBYOD", which only returns text when I'm connected to that network
;   - Any use for text-to-speech? ComObjCreate("SAPI.SpVoice").Speak("Speak this phrase")
;   - Popup menus are useful- can I use them elsewhere?
;   - Are timed tooltips usefull somewhere?
;
;
; Decisions
; ---------
;   - For Chrome extensions
;       - I decided not to use "Add URL to Window Title" because there is no whitelist option, and having
;         URL on every toolbar is ugly. Adding the input field id and name is cool and could be useful for
;         multi-page logins (like timesheet) but that is not REQUIRED for what I need (yet).
;         https://github.com/erichgoldman/add-url-to-window-title
;
;
; Hotkeys - Auto-Generated
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
; ~$^s            VS Code (AHK)       After save AHK file, autogenerate documentation, reload the current script
;
; (login/unlock)  Windows (AHK)       Set Slack status based on nearby wifi networks
; (on device chan Windows (AHK)       Backup passwords databases when insert backup flash drive
; #,              Windows (AHK)       Reduce primary screen's brightness
; #.              Windows (AHK)       Increase primary screen's brightness
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
; #Capslock       Windows (AHK)       Toggle selected text between lower/upper/sentence/title case
; #c              Windows (AHK)       Outlook Calendar
; #down           Windows (AHK)       Minimize active window instead of making it unmaximized, then minimize
; #d              Windows             Windows desktop
; #Esc            Windows (AHK)       Open my work password database
; #e              Windows             Windows Explorer
; #^+g            Windows (AHK)       Grammarly (in browser)
; #^g             Windows (AHK)       Google phrase fix
; #g              Windows (AHK)       Simple Google search
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
; #^+s            Windows (AHK)       Schema
; #^s             Windows (AHK)       Schema source code
; #s              Windows (AHK)       Source code
; #t              Windows (AHK)       Typora, with each virtual desktop having a different folder of files
; #up             Windows             Maximize active window
; #^+u            Windows (AHK)       Generate random UUID (uppercase)
; #^u             Windows (AHK)       Generate random UUID (lowercase)
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
EnvGet, WindowsUserDomain, USERDOMAIN
EnvGet, WindowsDnsDomain, USERDNSDOMAIN
EnvGet, WindowsUserProfile, USERPROFILE
UserEmailAddress = %WindowsUserName%@%WindowsDnsDomain%

; These come from my own Windows environment variables. See "My Automations Config.bat" for details
Global BackupDriveSerialNumber
EnvGet, JiraUrl, AHK_URL_JIRA
EnvGet, JiraMyProjectKeys, AHK_MY_PROJECT_KEYS_JIRA
EnvGet, JiraDefaultProjectKey, AHK_DEFAULT_PROJECT_KEY_JIRA
EnvGet, JiraDefaultRapidKey, AHK_DEFAULT_RAPID_KEY_JIRA
EnvGet, JiraDefaultSprint, AHK_DEFAULT_SPRINT_JIRA
EnvGet, SourceCodeUrl, AHK_URL_SOURCE_CODE
EnvGet, SourceSchemaUrl, AHK_URL_SOURCE_SCHEMA
EnvGet, SchemaVaultUrl, AHK_URL_SCHEMA_VAULT
EnvGet, TimesheetUrl, AHK_URL_TIMESHEET
EnvGet, WikiUrl, AHK_URL_WIKI
EnvGet, PersonalCloudUrl, AHK_URL_PERSONAL_CLOUD
EnvGet, NoiseBrownMP3, AHK_MP3_NOISE_BROWN
EnvGet, NoiseRailroadMP3, AHK_MP3_NOISE_RAILROAD
EnvGet, BackupDriveSerialNumber, AHK_BACKUP_DRIVE_SERIAL_NUMBER
EnvGet, PasswordApp, AHK_PASSWRD_APP
EnvGet, PasswordDBExtension, AHK_PASSWRD_DB_EXTENSION
EnvGet, PasswordDBFilenameBackupPattern, AHK_PASSWRD_DB_FILENAME_BACKUP_PATTERN

; Commonly used folders	
Global MyDocumentsFolder
Global MyPersonalFolder
Global MyPersonalDocumentsFolder
MyDocumentsFolder = %WindowsUserProfile%\Documents\
MyPersonalFolder = %WindowsUserProfile%\Personal\
MyPersonalDocumentsFolder = %WindowsUserProfile%\Personal\Documents\

; Standardize user-defined environment variables
If SubStr(MyPersonalDocumentsFolder, 0) != "\"
	MyPersonalDocumentsFolder = %MyPersonalDocumentsFolder%\
If SubStr(PasswordDBExtension, 1, 1) != "."
	PasswordDBExtension = .%PasswordDBExtension%




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

  ; Cache the list of connected monitors, so that when we want to increase/decrease the brightness, we
	; do not have to waste time getting this list
  Global MonitorIds
	MonitorIds := GetMonitorIds()

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
  OnMessage(WM_WTSSESSION_CHANGE, "OnWindowsUnlock")

  WM_DEVICECHANGE := 0x0219
  OnMessage(WM_DEVICECHANGE, "OnDeviceChange")

	NOTIFY_FOR_ALL_SESSIONS := 0
  hw_ahk := FindWindowEx(0, 0, "AutoHotkey", a_ScriptFullPath " - AutoHotkey v" a_AhkVersion)
  result := DllCall("Wtsapi32.dll\WTSRegisterSessionNotification", "uint", hw_ahk, "uint", NOTIFY_FOR_ALL_SESSIONS)
}
StopInterceptingWindowsUnlock() 
{
  DllCall("Wtsapi32.dll\WTSUnRegisterSessionNotification", "uint", hwnd)
}

;---------------------------------------------------------------------------------------------------------------------
; Windows Event- login or unlock (WM_WTSSESSION_CHANGE)
;---------------------------------------------------------------------------------------------------------------------
; (login/unlock)::      ; Windows|AHK|Set Slack status based on nearby wifi networks
OnWindowsUnlock(wParam, lParam)
{
  WTS_SESSION_UNLOCK := 0x8
  If (wParam = WTS_SESSION_UNLOCK)
	{
		SendInput #a                                  		  ; Open Windows Action Center to show new notifications from my phone
		MonitorIds := GetMonitorIds()                       ; Get list of connected monitors
    SlackStatusUpdate_SetSlackStatusBasedOnNetwork()		; If appropriate, then update my Slack status
	}
}

;---------------------------------------------------------------------------------------------------------------------
; Windows Event- device change (WM_DEVICECHANGE)
;
; This message is generated by docking/undocking a laptop, adding/removing a monitor, inserting/removing removable 
; media, and a few other events.
;---------------------------------------------------------------------------------------------------------------------
; (on device change)::  ; Windows|AHK|Backup passwords databases when insert backup flash drive
OnDeviceChange(wParam, lParam)
{
  ; DEBUGGING- Log this to a file so I can decide if I should my code to only certain values of wParam 
	;   What are:
	;     - Insert/remove USB drive:  DBT_DEVNODES_CHANGED
	;     - Dock/undock:
	;     - Plugin/remove monitor:
  ; DBT_CONFIGCHANGECANCELED    := 0x0019  ; A request to change the current configuration (dock or undock) has been canceled.
  ; DBT_CONFIGCHANGED           := 0x0018  ; The current configuration has changed, due to a dock or undock.
  ; DBT_CUSTOMEVENT             := 0x8006  ; A custom event has occurred.
  ; DBT_DEVICEARRIVAL           := 0x8000  ; A device or piece of media has been inserted and is now available.
  ; DBT_DEVICEQUERYREMOVE       := 0x8001  ; Permission is requested to remove a device or piece of media. Any application can deny this request and cancel the removal.
  ; DBT_DEVICEQUERYREMOVEFAILED := 0x8002  ; A request to remove a device or piece of media has been canceled.
  ; DBT_DEVICEREMOVECOMPLETE    := 0x8004  ; A device or piece of media has been removed.
  ; DBT_DEVICEREMOVEPENDING     := 0x8003  ; A device or piece of media is about to be removed. Cannot be denied.
  ; DBT_DEVICETYPESPECIFIC      := 0x8005  ; A device-specific event has occurred.
  ; DBT_DEVNODES_CHANGED        := 0x0007  ; A device has been added to or removed from the system. [Kummer- add/remove usb drive]
  ; DBT_QUERYCHANGECONFIG       := 0x0017  ; Permission is requested to change the current configuration (dock or undock).
  ; DBT_USERDEFINED             := 0xFFFF  ; The meaning of this message is user-defined.
	; If (wParam == DBT_CONFIGCHANGECANCELED)
	;   FileAppend, %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min% - DBT_CONFIGCHANGECANCELED`n, C:\Temp\AHK_OnDeviceChange.log
	; Else If (wParam == DBT_CONFIGCHANGED)
	;   FileAppend, %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min% - DBT_CONFIGCHANGED`n, C:\Temp\AHK_OnDeviceChange.log
	; Else If (wParam == DBT_CUSTOMEVENT)
	;   FileAppend, %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min% - DBT_CUSTOMEVENT`n, C:\Temp\AHK_OnDeviceChange.log
	; Else If (wParam == DBT_DEVICEARRIVAL)
	;   FileAppend, %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min% - DBT_DEVICEARRIVAL`n, C:\Temp\AHK_OnDeviceChange.log
	; Else If (wParam == DBT_DEVICEQUERYREMOVE)
	;   FileAppend, %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min% - DBT_DEVICEQUERYREMOVE`n, C:\Temp\AHK_OnDeviceChange.log
	; Else If (wParam == DBT_DEVICEQUERYREMOVEFAILED)
	;   FileAppend, %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min% - DBT_DEVICEQUERYREMOVEFAILED`n, C:\Temp\AHK_OnDeviceChange.log
	; Else If (wParam == DBT_DEVICEREMOVECOMPLETE)
	;   FileAppend, %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min% - DBT_DEVICEREMOVECOMPLETE`n, C:\Temp\AHK_OnDeviceChange.log
	; Else If (wParam == DBT_DEVICEREMOVEPENDING)
	;   FileAppend, %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min% - DBT_DEVICEREMOVEPENDING`n, C:\Temp\AHK_OnDeviceChange.log
	; Else If (wParam == DBT_DEVICETYPESPECIFIC)
	;   FileAppend, %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min% - DBT_DEVICETYPESPECIFIC`n, C:\Temp\AHK_OnDeviceChange.log
	; Else If (wParam == DBT_DEVNODES_CHANGED)
	;   FileAppend, %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min% - DBT_DEVNODES_CHANGED`n, C:\Temp\AHK_OnDeviceChange.log
	; Else If (wParam == DBT_QUERYCHANGECONFIG)
	;   FileAppend, %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min% - DBT_QUERYCHANGECONFIG`n, C:\Temp\AHK_OnDeviceChange.log
	; Else If (wParam == DBT_USERDEFINED)
	;   FileAppend, %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min% - DBT_USERDEFINED`n, C:\Temp\AHK_OnDeviceChange.log

	MonitorIds := GetMonitorIds()    ; Get list of connected monitors
  BackupPasswordDatabases()        ; If backup drive is inserted, then backup my password databases
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
; Convert case
;---------------------------------------------------------------------------------------------------------------------
#Capslock::       ; Windows|AHK|Toggle selected text between lower/upper/sentence/title case
  ConvertCase()
	Return



;---------------------------------------------------------------------------------------------------------------------
; Simple Google search for selected text
;
; I could just search for the selected text using this command:
;   Run, "https://www.google.com/search?q=%clipboard%"
; But I almost always want to add some qualifier or something to my searches, so instead, open Google, prefill the 
; search with the selected text and a space, and let me take it from there.
;---------------------------------------------------------------------------------------------------------------------
#g::        ; Windows|AHK|Simple Google search
  clipback := ClipboardAll
  clipboard=
  Send ^c
  ClipWait, 0
	
	Run, "https://www.google.com"
	WinWaitActive, Google, , 5
	If !ErrorLevel
  {
	  Sleep, 1250
    Send ^v{Space}
		Sleep, 500
  }
	
	clipboard := clipback
  Return



;---------------------------------------------------------------------------------------------------------------------
; Use Google search to fix text
;   - Code adapted from https://jacksautohotkeyblog.wordpress.com/2016/09/22/fixing-grammar-problems-with-google-search-intermediate-autohotkey-tip/
;---------------------------------------------------------------------------------------------------------------------
#^g::       ; Windows|AHK|Google phrase fix
  clipback := ClipboardAll
  clipboard=
  Send ^c
  ClipWait, 0
  UrlDownloadToFile % "https://www.google.com/search?q=" . clipboard, temp
  FileRead, contents, temp
  FileDelete temp
  If (RegExMatch(contents, "Showing results for <a.*?>(.*?)</a>", match))
  {
    StringReplace, clipboard, match1, <b><i>,, All
    StringReplace, clipboard, clipboard, </i></b>,, All
  }
  Send ^v
  Sleep, 500
  clipboard := clipback
  Return



;---------------------------------------------------------------------------------------------------------------------
; Paste the selected text into a new document on Grammarly's web site
;
; To create a new document in Grammarly.com, navigate to https://app.grammarly.com/docs/new.
;   - It takes a few seconds to load
;   - When it is almost ready to allow editing, it changes the URL to something like this:
;     https://app.grammarly.com/ddocs/565623631. 
;   - Recent versions of Chrome don't allow us to easily or reliably get the URL from code
;   - I could press !d/^i/f6 to get to the Chrome omnibar so I can copy the URL, but there is no easy/consistent 
;     way to get back to the page content
;   - My solution is to use the Chrome extension called "Url in Title" by Guillaume Ryder, which adds the URL to the
;     title of the webpage, so AHK can watch for it. You can specify a whitelist so that only the whitelisted sites
;     display the URL in the title. I confligured the extension like this:
;       Tab title format: {title} - {protocol}://{hostname}{port}/{path}
;       Whitelist: https://app.grammarly.com
;
; Decisions
;   - Grammarly Windows app was my first choice, but is very keyboard unfriendly, so had to use the web page
;
; Dependencies
;   - Chrome extension "URL in Title". https://github.com/erichgoldman/add-url-to-window-title
;---------------------------------------------------------------------------------------------------------------------
#^+g::    ; Windows|AHK|Grammarly (in browser)
	clipback := ClipboardAll
  clipboard=
  Send ^c
  ClipWait, 0
  Run, "https://app.grammarly.com/docs/new"
	WinWaitActive, Grammarly - https://app.grammarly.com/ddocs, , 15
	If !ErrorLevel
  {
	  Sleep, 1250
    Send ^v
    Sleep, 500
  }
  clipboard := clipback
  Return



;---------------------------------------------------------------------------------------------------------------------
; Open my work password database, minimized, if it is not already running
;---------------------------------------------------------------------------------------------------------------------
#Esc::         ; Windows|AHK|Open my work password database
	StartPasswordManager()
  Return

StartPasswordManager() 
{
	Global PasswordApp
	Global MyPersonalDocumentsFolder
	Global WindowsUserDomain
	Global PasswordDBExtension

	SplitPath, PasswordApp , passwordAppExe,
	If !ProcessExists(passwordAppExe)
	{
		WorkPasswordDatabase = %MyPersonalDocumentsFolder%%WindowsUserDomain%%PasswordDBExtension%
	  Run, %ComSpec% /c ""`%AHK_PASSWRD_APP`%""" ""%WorkPasswordDatabase%" -pw-enc:`%AHK_PASSWRD_DB_PASSWRD`% -minimize,, Hide
	}
	Return
}



;---------------------------------------------------------------------------------------------------------------------
; Control screen brightness
;   - I WANT to use 
;       - #,/#. to update the PRIMARY screen so that most of the time I can do that and not have one command if I'm on
;         my laptop screen and another command if I'm on an external monitor.
;       - #^,/#^. to update a SECONDARY screen
;       - THIS WILL LIKELY NEED ADJUSTED WHEN I GO BACK TO MY MULTIPLE MONITORS SETUP
;
; Decisions
; ---------
;   - NirSoft's NirCmd isn't bad for changing laptop screen's brightness, just no on-screen display.
;     https://nircmd.nirsoft.net/changebrightness.html
;     Run, %MyPersonalFolder%PortableApps\nircmd\nircmdc.exe changebrightness %changeValue%,, Hide
;
; The Future 
; ----------
;   - Instead of using custom AHK code and NirSoft utilities, use Powershell and WMI
;   - Get-WmiObject win32_pnpentity | Where{$_.service -match "monitor"} | Select DeviceID, HardwareID, Name
;   - Get-WmiObject win32_desktopmonitor | Select Name, DeviceID, PNPDeviceID
;   - Get-WmiObject is discontinued!!!
;
; Dependencies: 
;   - BrightnessSetter. Only works on laptop screen.
;       https://github.com/qwerty12/AutoHotkeyScripts/tree/master/LaptopBrightnessSetter
;   - NirSoft's ControlMyMonitor. Only works on monitors.
;       https://www.nirsoft.net/utils/control_my_monitor.html
;---------------------------------------------------------------------------------------------------------------------
#,::        ; Windows|AHK|Reduce primary screen's brightness
	AdjustBrightness(-10)
  Return
  
#.::        ; Windows|AHK|Increase primary screen's brightness
	AdjustBrightness(10)
	Return

; TODO: Move this and AdjustBrightness to utilities class
;---------------------------------------------------------------------------------------------------------------------
; Get the list of monitor ids currently connected to the system
;   - Uses NirSoft's ControlMyMonitor- https://www.nirsoft.net/utils/control_my_monitor.html
;   - Instead of writing the list of monitors to a file, can specify an empty string and contents will be written to 
;     standard output. But that generates a black DOS box and I don't like it.
;       %MyPersonalFolder%PortableApps\ControlMyMonitor\ControlMyMonitor.exe /smonitors "" | more | findstr /c:"Monitor ID"
;---------------------------------------------------------------------------------------------------------------------
GetMonitorIds()
{
	monitorIds := []

  Run, %MyPersonalFolder%PortableApps\ControlMyMonitor\ControlMyMonitor.exe /smonitors "monitors.txt"
	Loop, Read, monitors.txt
  {
    RegExMatch(A_LoopReadLine, "Monitor ID: ""(.*)""", monitorId)
	 	If (monitorId1 != "") 
	 	{
	 		monitorIds.Push(monitorId1)
	 	}
  }
  FileDelete monitors.txt

  ;for index, element in MonitorIds
  ;  MsgBox % "Element number " . index . " is " . element

	Return monitorIds
}  

;---------------------------------------------------------------------------------------------------------------------
; Adjust the brightness of the primary screen
;   - Uses NirSoft's ControlMyMonitor- https://www.nirsoft.net/utils/control_my_monitor.html
;---------------------------------------------------------------------------------------------------------------------
AdjustBrightness(adjustment) {
  RunWait, %MyPersonalFolder%PortableApps\ControlMyMonitor\ControlMyMonitor.exe /GetValue Primary DF,, UseErrorLevel
	If (ErrorLevel == "0")
	{
    ; No external monitors
		BrightnessSetter.SetBrightness(adjustment)
	}
	Else
	{
    ; TODO- Loop through MonitorIds and replace Primary with monitor id
    Run, %MyPersonalFolder%PortableApps\ControlMyMonitor\ControlMyMonitor.exe /ChangeValue "Primary" 10 %adjustment% 
	}
}



;---------------------------------------------------------------------------------------------------------------------
; Backup password databases
;   - If the flash drive I use to backup my password database is inserted, then back them up to that drive
;---------------------------------------------------------------------------------------------------------------------
BackupPasswordDatabases()
{
	Global MyPersonalDocumentsFolder
	Global PasswordDBExtension
	Global PasswordDBFilenameBackupPattern

	backupDriveLetter := GetDriveLetter(BackupDriveSerialNumber, "REMOVABLE")
	If (backupDriveLetter)
	{
		index = 1
		Loop Files, %MyPersonalDocumentsFolder%*%PasswordDBExtension%
    {
		  destinationFilename = %backupDriveLetter%:\%PasswordDBFilenameBackupPattern%
		  destinationFilename := StrReplace(destinationFilename, "%INDEX%", index)
    
		  ; FileCopy - last parameter determines if overwrite
      FileCopy, %A_LoopFileFullPath%, %destinationFilename%, 1

	    ; MsgBox options: 64 (Info icon) + 4096 (System Modal- always on top)
		  MsgBox, 4160, Password Database Backup, %A_LoopFileFullPath%`nhas been backed up to %destinationFilename%.
	    index++
    }
	}
}



;---------------------------------------------------------------------------------------------------------------------
; Dark mode
;
; In some applications, there is no easy way to determine what the current theme is. There's no textbox (or there is 
; but AHK can't access it), or the theme is a control that's checked/unchecked in a menu, and it's not persisted 
; somewhere we can easily get to it (either in the registry or a file). In those cases, I have to get the color of
; some pixel and if that pixel is a dark color, then I ASSUME we're displaying a dark theme, else I ASSUME that we're 
; using a light theme.
;---------------------------------------------------------------------------------------------------------------------
#Space::     ; Windows|AHK|Toggle dark mode for active application
	If WinActive("ahk_exe chrome.exe")
	{
		; In Chrome, extension Dark Reader has an existing shortcut Alt+Shift+D that toggles between dark/light
		; Dependencies: Chrome extension Dark Reader (https://darkreader.org)
	  SendInput !+D   
	}
	Else If WinActive("- KeePass")
	{
	  ; KeePass plugin KeeTheme has an existing shortcut Ctrl+T that toggles Dark Theme on and off.
	  ; Dependencies: KeePass plugin KeeTheme (https://github.com/xatupal/KeeTheme)
	  SendInput ^t   
	}
	Else If WinActive("- Visual Studio Code")
	{
		; In Visual Studio Code, have to look at a pixel to determine which theme is active
		;   "Light+ (default light)" is the light theme
		;   "Dark+ (default dark)" is the dark theme
		blue := GetPixelsBlueValue(20, 70)
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
		; In IntelliJ IDEA, have to look at a pixel to determine which theme is active
		;   "IntelliJ/Default/Light" is the light theme
		;   "Darcula" is the dark theme
		blue := GetPixelsBlueValue(17, 80)
		SendInput ^!s

		;Wait for "Settings" window
		WinWaitActive, Settings, , 10
		If !ErrorLevel
  	{
			; Search for appearance section
			Sleep, 500
			SendInput appearance
			Sleep, 500
			SendInput {ENTER}
			Sleep, 250
			SendInput {DOWN}
			Sleep, 500

      ; Then search for theme and move to the theme dropdown box
			SendInput ^atheme
			Sleep, 500
			SendInput {ENTER}
			Sleep, 250

			; Change the theme and exit the dialog
			SendInput {DOWN}
			Sleep, 250
			If blue < 75
				SendInput {DOWN}{DOWN}
			Else
				SendInput {UP}{UP}
			Sleep, 250
			SendInput {ENTER}
			Sleep, 250
			SendInput {ENTER}
	  }
	}
	Else If WinActive("ahk_class Notepad++")
  {

	  ; In Notepad++, go to Settings => Style Configurator to toggle between themes
		;   "Default" is the light theme
		;   "VS2015-Dark" is the dark theme
		; Dependencies: Notepad++ theme VS2015-Dark
		SendInput {ALT}ts{ENTER}
		WinWaitActive, Style Configurator,, 2
		ControlGet, currentTheme, Choice,, ComboBox1, Style Configurator
		ControlFocus, ComboBox1, Style Configurator
		If InStr(currentTheme, "Dark")
		  SendInput d
		Else
		  SendInput vvv
		SendInput {TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{ENTER}
  }	
	Else If WinActive("- Eclipse IDE")
	{
	  ; In Eclipse settings Preferences => Editor => Keys, I MANUALLY added a shortcut for Ctrl+Shift+8 to open 
		; Preferences => General => Appearance. So this code can use that shortcut to programmatically open
		; Preferences => Appearance and then toggle the Theme between "Classic" and "Dark".
		;   "Classic" is the light theme
		;   "Dark" is the dark theme
		SendInput ^+{8}     
		WinWaitActive, Preferences,, 2
		ControlGet, currentTheme, Choice,, ComboBox1, Preferences
		ControlFocus, ComboBox1, Preferences
		If InStr(currentTheme, "Dark")
			SendInput c
		Else
			SendInput d
    SendInput !a         ; Apply changes
    SendInput {Escape}   ; Close dialog saying restart necessary for full effect
    SendInput {Escape}   ; Close Preferences dialog
	}
	Else If WinActive("Microsoft Visual Studio")
	{
	  ; In Visual Studio, Tools => Options
		;   "Blue" is the light theme
		;   "Dark" is the dark theme
    SendInput !to    
		WinWaitActive, Options,, 2
		SendInput ^eVisual experience{TAB}
		Sleep, 500
		ControlGet, currentTheme, Choice,, ComboBox1, Options
		ControlFocus, ComboBox1, Options
		If InStr(currentTheme, "Dark")
			SendInput b
		Else
		  SendInput d
		SendInput {ENTER}
	}
	Else If WinActive("- Typora")
	{
		; In Typora, have to look at a pixel to determine which theme is active
		;   "Pixyll" is the light theme
		;   "Night" is the dark theme
		blue := GetPixelsBlueValue(20, 70)
		SendInput !t
		Sleep, 100
		If blue < 55
			SendInput p
		Else
			SendInput nn{ENTER}
	}
	Else If WinActive("ahk_exe explorer.exe")
	{
	  ; In Windows Explorer, light and dark mode is controlled by a registry key, and
		;   1 is the light mode
		;   0 is the dark mode
		RegRead, appsUseLightTheme, HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize, AppsUseLightTheme
		appsUseLightTheme:=!appsUseLightTheme
		RegWrite, REG_DWORD, HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize, AppsUseLightTheme, %appsUseLightTheme%
	}
	Else If WinActive("ahk_class VirtualConsoleClass")
	{
	  ; ConEmu Settings => Features => Colors
		;   "Tomorrow Night Blue" is the light theme
		;   "Cobalt2" is the dark theme
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
		; In Microsoft Outlook, have to look at a pixel to determine which theme is active
		;   "Colorful" is the light theme
		;   "Black" is the dark theme
		SendInput !f
		Sleep, 100
		SendInput d
		Sleep, 500
		blue := GetPixelsBlueValue(200, 200)
		SendInput y1
		If blue < 55
			SendInput c
		Else
			SendInput b
		SendInput {ENTER}{ESCAPE}
	}
  Return

GetPixelsBlueValue(x, y) 
{
	PixelGetColor, color, x, y
	blue:="0x" SubStr(color,3,2)   ; substr is to get the piece
	blue:=blue+0                   ; add 0 to convert it to decimal

	Return blue
}


	
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
; Generate and output a UUID/GUID
;---------------------------------------------------------------------------------------------------------------------
#^u::     ; Windows|AHK|Generate random UUID (lowercase)
  newGUID := CreateGUID()
	StringLower, newGUID, newGUID
  SendInput %newGUID%
	Return

#^+u::    ; Windows|AHK|Generate random UUID (uppercase)
  newGUID := CreateGUID()
	StringUpper, newGUID, newGUID
  SendInput %newGUID%
	Return



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
    ;Run, "%WindowsLocalAppDataFolder%\slack\slack.exe"
		Run, "%WindowsProgramFilesFolder%\Slack\Slack.exe"
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
;   - If VS Code is running as administrator and I right-click a file and choose "Open with Code", the open fails with
;     the message "A second instance of Code is already running as Administrator." The problem is that the instance
;     of VS Code being opened by the shell is running as a standard user. The solution I'm using is for AHK to start
;     VS Code as a standard user, and to manually open VS Code as administrator in the very few cases that's necessary.
;   - When editing any AutoHotKey script in VS Code, clicking Ctrl-S to save the script also causes AutoHotKey to 
;     reload the current script. This eliminates the need to right-click the AutoHotKey system tray icon and select 
;     "Reload This Script".
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
;
; COMMAND LINE OPTIONS - https://code.visualstudio.com/docs/editor/command-line
;   -n or --new-window	  				Opens a new session of VS Code instead of restoring the previous session (default).
;   -r or --reuse-window					Forces opening a file or folder in the last active window.
;   --install-extension <ext>			Install an extension. Provide the full extension name publisher.extension as an argument. Use --force argument to avoid prompts.
;   --uninstall-extension <ext>		Uninstall an extension. Provide the full extension name publisher.extension as an argument.
;   --disable-extensions					Disable all installed extensions. Extensions will still be visible in the Disabled section of the Extensions view but they will never be activated.
;   --list-extensions							List the installed extensions.
;   --show-versions								Show versions of installed extensions, when using --list-extensions
;   --enable-proposed-api <ext>		Enables proposed api features for an extension. Provide the full extension name publisher.extension as an argument.
;---------------------------------------------------------------------------------------------------------------------
#IfWinActive .ahk - Visual Studio Code 
~$^s::      ; VS Code|AHK|After save AHK file, autogenerate documentation, reload the current script
	AutoGenerateDocumentation(A_ScriptName)
  Reload
  Return
#IfWinActive

#v::        ; Windows|AHK|Visual Studio Code
	;Run "%WindowsLocalAppDataFolder%\Programs\Microsoft VS Code\Code.exe"

  vsCodeExe = %WindowsLocalAppDataFolder%\Programs\Microsoft VS Code\Code.exe
  ShellRun(vsCodeExe)
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
			Run "%WindowsLocalAppDataFolder%\Programs\Microsoft VS Code\Code.exe"
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
;#g::   
;           ; Windows|AHK|gTasks Pro
;  If Not WinActive("gTasks Pro ahk_exe ApplicationFrameHost.exe") 
;	{
;    ; Note that Windows store apps are more complex than simple exe's
;	  Run, "%MyPersonalFolder%\WindowsStoreAppLinks\gTasks Pro.lnk"
;  }
;	WinActivate, gTasks Pro
;  Return

	

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
;   - Start the password manager in case we have to login
;   - If the login page appears within 30 seconds then login, else assume I was logged in automatically by Active 
;     Directory. 
;---------------------------------------------------------------------------------------------------------------------
#+4::             ; Windows|AHK|Timesheet (Shift-4 is $)
  StartPasswordManager()
	Run, "%TimesheetUrl%",, Max

	WinWaitActive, Login, , 30
  If !ErrorLevel
  {
		SendInput ^!a
	}
  Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Source code, schema (source), and schema vault
;---------------------------------------------------------------------------------------------------------------------
#s::              ; Windows|AHK|Source code
	Run, "%SourceCodeUrl%",, Max
  Return

#^s::             ; Windows|AHK|Schema source code
	Run, "%SourceSchemaUrl%",, Max
  Return

#^+s::            ; Windows|AHK|Schema
	Run, "%SchemaVaultUrl%",, Max
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



;---------------------------------------------------------------------------------------------------------------------
; All of the following commented-out hotkeys, etc are here so that the auto-generated documentation includes them
;---------------------------------------------------------------------------------------------------------------------
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
#include %A_ScriptDir%\Lib\BrightnessSetter.ahk
#include %A_ScriptDir%\Lib\Convert Case.ahk





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