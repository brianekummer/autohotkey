﻿;---------------------------------------------------------------------------------------------------------------------
; Brian Kummer's AutoHotKey script
;
; - Near the bottom of this script are a number of #include statements to include libraries of utility functions
;
; Miscellaneous:
;   - Definition of AutoHotKey keys: http://www.autohotkey.com/docs/KeyList.htm
;   - Send, SendInput modifiers: ^ = Ctrl, ! = Alt, + = Shift, # = Windows
;
; Summary (3/22/2018)
; -------------------
;   Add auto-correct to EVERY application (http://www.biancolo.com/articles/autocorrect/)
;
;   Modifying System Behavior:
;     printscreen               Starts Windows snipping tool in rectangle mode. Shift+Prt Scr still takes a screenshot of the whole screen.
;     Win+down                  Minimize the active window instead of making it unmaximized, then minimize
;     Shift+mousewheel          Adjusts system volume 
;
;   Mouse Behavior
;     XButton1                  Minimizes the current application
;     XButton2                  Depending on the active application, minimizes the window, closes the window/tab (Ctrl-F4),
;                               or closes the application (Alt-F4)
;
;   Other hotkeys
;     Win+c                     Calendar- Activate Outlook and goto Calendar
;     Win+i                     Inbox- Activate Outlook and goto my Inbox
;     Win+k                     slacK- Activate Slack and opens "Jump to"
;     Win+m                     windows Media player
;     Win+n                     Notepad++- Opens Notepad++
;     Win+p                     Progress- Activate Outlook and goto Tasks
;     Win+t                     Activate Typora
;     Win+z                     noiZe- Open SimplyNoise.com
;
;   Modifying application behavior
;     Typora                    Ctrl+mousewheel to zoom
;     Notepad++                 After save ahk file in Notepad++, reload the current script in AutoHotKey
;     Slack                     Typing "/lunch" gets changed to "/status :hamburger: At lunch"
;                               Typing "/mtg" gets changed to "/status :spiral_calendar_pad: In a meeting"
;                               Typing "/wfh" gets changed to "/status :house: Working remotely"
;
;  For development (Win-Shift and some key)
;     Win+Shift+c               Command prompt- Open a command prompt (because we're using RunAsAdmin() in this script is as admin)
;     Win+Shift+a               ADP
;     Win+Shift+b               BitBucket
;     Win+Shift+g               Git
;     Win+Shift+j               JIRA
;     Win+Shift+n               Notepad++- Open Notepad++, and paste the selected text into the newly opened window
;     Win+Shift+v               Visual Studio 2017
;     Win+Shift+w               Wiki

	; OLD JUNK---
;     Win-k:                    KLOVE- Open K-Love in Media Player, or if already open, open K-Love website

;---------------------------------------------------------------------------------------------------------------------
#NoEnv
#Persistent



;---------------------------------------------------------------------------------------------------------------------
; Global variables - configured from Windows environment variables
;---------------------------------------------------------------------------------------------------------------------

; These come from Windows-defined environment variables
EnvGet, WindowsLocalAppDataFolder, LOCALAPPDATA
EnvGet, WindowsProgramFilesX86Folder, PROGRAMFILES(X86)
EnvGet, WindowsProgramFilesFolder, PROGRAMFILES
EnvGet, WindowsUserName, USERNAME
EnvGet, WindowsDnsDomain, USERDNSDOMAIN
UserEmailAddress = %WindowsUserName%@%WindowsDnsDomain%

; These come from my own Windows environment variables; see "My Automations Config.bat" for details
EnvGet, BitBucketUrl, AHK_URL_BITBUCKET
EnvGet, JiraUrl, AHK_URL_JIRA
EnvGet, TimesheetUrl, AHK_URL_TIMESHEET
EnvGet, WikiUrl, AHK_URL_WIKI
EnvGet, CentrifyUrl, AHK_URL_CENTRIFY
EnvGet, CitrixUrl, AHK_URL_CITRIX
	
	


;---------------------------------------------------------------------------------------------------------------------
; Upon startup of this script, perform these actions
;---------------------------------------------------------------------------------------------------------------------
  SetTitleMatchMode RegEx            ; Make windowing commands use regex
  OnExit("ExitFunc")                 ; Upon exit of this script, execute this function

  RunAsAdmin()
  StartInterceptingWindowsUnlock()

  ; Configure wallpapers for Windows virtual desktops - 
	; REQUIRES Windows environment variables AHK_VIRTUAL_DESKTOP_WALLPAPER_x such as:
	;   AHK_VIRTUAL_DESKTOP_WALLPAPER_1 = Main    |C:\Pictures\wallpaper1.jpg
  ;   AHK_VIRTUAL_DESKTOP_WALLPAPER_2 = Personal|C:\Pictures\wallpaper2.jpg
	; See "Virtual Desktop Accessor.ahk" for details
	VirtualDesktopAccessor_Initialize()

	; Configure Slack status updates based on the network
  ; REQUIRES some Windows environment variables - see "Slack Status Update.ahk" for details
  SlackStatusUpdate_Initialize()
  SlackStatusUpdate_SetSlackStatusBasedOnNetwork()
	
  Return

	

;---------------------------------------------------------------------------------------------------------------------
; Upon exit of this script, perform these actions
;---------------------------------------------------------------------------------------------------------------------
ExitFunc(ExitReason, ExitCode)
{
  StopInterceptingWindowsUnlock()
}



;---------------------------------------------------------------------------------------------------------------------
; Hook into Windows system message WM_WTSSESSION_CHANGE so that when I log into the computer, I can do stuff. Code 
; taken from http://autohotkey.com/board/topic/23095-trouble-detecting-windows-unlock-only-when-compiled
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
OnWindowsUnlock(wParam, lParam)
{
  WTS_SESSION_UNLOCK := 0x8
  If (wParam = WTS_SESSION_UNLOCK)
	{
		SendInput #a                                      ; Open Windows Action Center to show any new notifications from my phone
    SlackStatusUpdate_SetSlackStatusBasedOnNetwork()  ; If appropriate, update my Slack status
	}
}




	

;---------------------------------------------------------------------------------------------------------------------
; Temporary stuff goes here. Uses Win+(dash on numeric keypad) as hotkey.
;---------------------------------------------------------------------------------------------------------------------
#NumpadSub::
	Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Slack
; -----
;   Win+k:    Open Slack and go to the "Jump to" window
;
; Functionality in Slack-Status-Update.ahk
; -----------------------------------------
;   /xxxxx    When in Slack, replace "/lunch", "/wfh", "/mtg" with the equivalent Slack command to change my status
;---------------------------------------------------------------------------------------------------------------------
#k::
  If Not WinExist("ahk_group SlackStatusUpdate_WindowTitles")
  {
    Run, "%WindowsLocalAppDataFolder%\slack\slack.exe"
  	
		WinActivate, ahk_group SlackStatusUpdate_WindowTitles
	  Sleep, 1000
  	WinMaximize
  }
  WinActivate
	;WinMaximize
  Return

+#k::
  If Not WinExist("ahk_group SlackStatusUpdate_WindowTitles")
  {
    Run, "%WindowsLocalAppDataFolder%\slack\slack.exe"
  	WinActivate, ahk_group SlackStatusUpdate_WindowTitles
	  Sleep, 1000
  	WinMaximize
  }
	WinActivate
	;WinMaximize
	SendInput ^k
  Return
	
	


;---------------------------------------------------------------------------------------------------------------------
; Shift+mousewheel: Change system volume
;   Other keystroke combinations (Win+mousewheel, or Ctrl+mousewheel) affect the current window
;
;   Because I also use Win-10-virtual-desktop-enhancer (https://github.com/sdias/win-10-virtual-desktop-enhancer),
;   which has an option to change desktops when you mouse over the system tray and scroll the mouse wheel, I had
;   to turn that setting off in its settings.ini (set TaskbarScrollSwitching=0)
;---------------------------------------------------------------------------------------------------------------------
+WheelUp::   SendInput {Volume_Up 1}
+WheelDown:: SendInput {Volume_Down 1}



;---------------------------------------------------------------------------------------------------------------------
; Extra mouse buttons
;   - XButton1 minimizes the current window
;   - XButton2, depending on the active window, closes a WINDOW, or minimizes or closes the active application.
;
; It is based on the name of app/process, NOT the window title, or else it would minimize a browser with a tab 
; whose title is "How to Use Slack". Also, Microsoft Edge browser is more complex than a single process, so detecting
; it is more complex.
;---------------------------------------------------------------------------------------------------------------------
XButton1::
  WinMinimize, A
	Return

XButton2::
  WinGet, processName, ProcessName, A
  SplitPath, processName,,,, processNameNoExtension

	If RegExMatch(processNameNoExtension, "i)skype|outlook|wmplayer|slack|typora")
		WinMinimize, A     ; Do not want to close these apps
  Else If RegExMatch(processNameNoExtension, "i)chrome|iexplore|firefox|notepad++|ssms|devenv") or WinActive("Microsoft Edge ahk_exe ApplicationFrameHost.exe")
    SendInput ^{f4}         ; Close a WINDOW/TAB/DOCUMENT
  Else
    SendInput !{f4}         ; Close the APP
  Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Notepad++
;
; When editing any AutoHotKey script in Notepad++, clicking Ctrl-S to save the script automatically forces AutoHotKey
; to reload the current script. This eliminates the need to right-click the AutoHotKey system tray icon and select
; "Reload This Script".
;
; Win+n          Open Notepad++
; Win+Shift+n    Open Notepad++ and paste the selected text into the new Notepad++ window
;                  - I could have copied the selected text into a variable and done a "Send %variableName%", which 
;                    would send each character as if it was typed, which would be very SLOW. Instead, I paste the 
;                    selected text into the Notepad++ window.
;                  - If the pasted text looks like a recognized type of text, format it using a Notepad++ plugin.
;                    THESE PLUGINS -= MUST =- BE INSTALLED VIA PLUGINS => PLUGIN MANAGERS => SHOW PLUGIN MANAGER
;                      Type       Plugin Name   Plugin Hot Key     Comments
;                      --------   -----------   ----------------   ------------------------------------------------------
;                      HTML/XML   XML Tools     Ctrl+Shift+Alt+b   Also sets Notepad++ language to "XML" to enable folding
;                      JSON       JSTool        Ctrl+Alt+m
;                      SQL        SQLinForm     Alt+Shift+f
;---------------------------------------------------------------------------------------------------------------------
#IfWinActive .ahk - Notepad++ 
~$^s:: 
  Reload
  Return
#IfWinActive

#n::
	Run "%WindowsProgramFilesX86Folder%\Notepad++\notepad++.exe"
  Return

+#n::
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

  ClipSaved := ClipboardAll
  Clipboard = 
  SendInput ^c
  ClipWait, 1
  If (!ErrorLevel)
  {
    Run "%WindowsProgramFilesX86Folder%\Notepad++\notepad++.exe"
    WinWaitActive, "Notepad++",,2     ; Wait up to 2 seconds for Notepad++ to open

    SendInput ^v
    Sleep, 500
		
    ; We can do some extra formatting of some types of data using Notepad++ plugins, 
    ; based on the type of the data we're viewing
    newText := Clipboard
    If RegExMatch(newText, regExHtmlOrXml)
    {
      SendInput ^+!{b}    ; Plugins => XML Tools => Pretty print...
      SendInput !{l}x     ; Language => XML
    }
    Else If RegExMatch(newText, regExJson)
    {
      SendInput ^!{m}     ; Pluginx => JSTool =>    JSFormat
    }
    Else If RegExMatch(newText, regExSql)
    {
      SendInput !+{f}     ; Plugins => SQLinForm => Format Selected SQL 
    }
  }
  Clipboard := ClipSaved	; Restore the original clipboard. Note the use of Clipboard (not ClipboardAll)
  ClipSaved =			        ; Free the memory in case the clipboard was very large
  Return	



;---------------------------------------------------------------------------------------------------------------------
; Typora
;   Ctrl+mousewheel:  Zoom in and out
;   Win+T:            Open Typora if not already open
; Note that a "•" character will be displayed after the filename if the file has unsaved changes. Since I'm editing
; NextCloud notes, those usually have a .txt extension.
;---------------------------------------------------------------------------------------------------------------------
#IfWinActive \.(md|txt)•? - Typora 
  ^WheelUp::   SendInput ^+{=}
  ^WheelDown:: SendInput ^+{-}
#IfWinActive

#t::
  If Not WinExist("\.(md|txt)•? - Typora")
	{
  	Run "%WindowsProgramFilesFolder%\Typora\Typora.exe"
	}
  WinActivate, Typora
  WinMaximize, Typora
  Return



;---------------------------------------------------------------------------------------------------------------------
; Win+Ctrl+v: Paste the clipboard as plain text. https://autohotkey.com/board/topic/10412-paste-plain-text-and-copycut
;---------------------------------------------------------------------------------------------------------------------
^#v::                            ; Text–only paste from ClipBoard
  Clip0 = %ClipBoardAll%
  ClipBoard = %ClipBoard%       ; Convert to text
  SendInput ^v                       ; For best compatibility: SendPlay
  Sleep 50                      ; Don't change clipboard while it is pasted! (Sleep > 0)
  ClipBoard = %Clip0%           ; Restore original ClipBoard
  VarSetCapacity(Clip0, 0)      ; Free memory
  Return

	

;---------------------------------------------------------------------------------------------------------------------
; Win+Shift+g: Open Git. ConEmu has multiple windows, so cannot just use "WinActivate, Git Bash"
;---------------------------------------------------------------------------------------------------------------------
+#g::
  If Not WinExist("ahk_class VirtualConsoleClass")
  {
    Run "%WindowsProgramFilesFolder%\ConEmu\ConEmu64.exe" -run {Shells::Git Bash}
    WinWait ahk_class VirtualConsoleClass
  }
  WinActivate
  Return

	

;---------------------------------------------------------------------------------------------------------------------
; Win+Shift+v: Run VISUAL STUDIO
;---------------------------------------------------------------------------------------------------------------------
+#v::
	If Not WinExist("Microsoft Visual Studio")
	{
	  Run *RunAs "%WindowsProgramFilesX86Folder%\Microsoft Visual Studio\2017\Enterprise\Common7\IDE\devenv.exe"
	}
  WinActivate, Microsoft Visual Studio
  Return  

	
	
;---------------------------------------------------------------------------------------------------------------------
; K-Love radio:
;   #k:               Use Windows Media Player to open a playlist to play K-Love. If it's already open, open a browser
;                     to K-Love's web site so I can see what song is playing. 
;                       Unused right now: http://www.klove.com/listen/player.aspx
;   Media Play Pause: When listening to K-LOVE, pause is disabled, so map the pause button to press media stop. This 
;                     requires using the Windows Media Player plug-in Windows Media Player Plus!
;---------------------------------------------------------------------------------------------------------------------
;#k::
;  If WinExist("Encouraging K-LOVE")
;    Run http://www.klove.com
;	Else
;		Run "C:\Program Files (x86)\Windows Media Player\wmplayer.exe" /task MediaLibrary /prefetch:1 /play "%MyMemoryCardDrive%\Music\Playlists\K-Love Radio.wpl"
;  Return
;#IfWinExist Encouraging K-LOVE
;Media_Play_Pause::
;  Send {Media_Stop}
;	Return
;#IfWinExist






;---------------------------------------------------------------------------------------------------------------------
; Win+Shift+a: ADP - Go to ADP website to enter my timesheet
;---------------------------------------------------------------------------------------------------------------------
+#a::
	Run, "%TimesheetUrl%",, Max
  Return

	
;---------------------------------------------------------------------------------------------------------------------
; Win+Shift+b: BitBucket
;---------------------------------------------------------------------------------------------------------------------
+#b::
	Run, "%BitBucketUrl%",, Max
  Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Win+Shift+w: Wiki
;---------------------------------------------------------------------------------------------------------------------
+#w::
	Run, "%WikiUrl%",, Max
  Return



;---------------------------------------------------------------------------------------------------------------------
; Win+Shift+y: centrifY
;---------------------------------------------------------------------------------------------------------------------
+#y::
	Run, "%CentrifyUrl%",, Max
  Return


	
;---------------------------------------------------------------------------------------------------------------------
; Win+Shift+x: citriX
;---------------------------------------------------------------------------------------------------------------------
+#x::
	Run, "%CitrixUrl%",, Max
  Return

	
  
;---------------------------------------------------------------------------------------------------------------------
; Prt Scr: Open the Windows snipping tool in rectangular mode. Image is stored in the clipboard.
;          Shift+Prt Scr still takes a screenshot of the whole screen.
;---------------------------------------------------------------------------------------------------------------------
printscreen::
  Run snippingtool /clip
  Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Win+m: Windows Media Player
;---------------------------------------------------------------------------------------------------------------------
#m::
	If Not WinExist("Windows Media Player")
	{
	  Run "%WindowsProgramFilesX86Folder%\Windows Media Player\wmplayer.exe"
	}
  WinActivate, ahk_class WMPlayerApp
  Return  

	

;---------------------------------------------------------------------------------------------------------------------
; Win+Shift+j: JIRA- Open JIRA
;   - If the highlighted text looks like a JIRA story number (e.g. TRAN|IQTC|OCS|DA-x[xxx]), then open that story
;   - If the Git Bash window has text that looks like a JIRA story number, then open that story
;---------------------------------------------------------------------------------------------------------------------
+#j::
  regexString := "i)\b(tran|iqtc|ocs|da)[-_]?\d{1,4}\b"
  RegExMatch(GetSelectedTextUsingClipboard(), regexString, storyId)

  If StrLen(storyId) = 0
  { 
	  ; Search for a ConEmu terminal with a JIRA story number
		WinGetTitle, git_window_title, ahk_exe i)\\conemu64\.exe$ ahk_class VirtualConsoleClass
		RegExMatch(git_window_title, regexString, storyId)
	}

  If StrLen(storyId) = 0
  { 
	  ; Search for a Mintty terminal (comes with Git) with a JIRA story number
		WinGetTitle, git_window_title, ahk_exe i)\\mintty\.exe$ ahk_class mintty
		RegExMatch(git_window_title, regexString, storyId)
  }  
		
	If StrLen(storyId) = 0
  {
	  ; Could not find any JIRA story number
		Run, %JiraUrl%/secure/RapidBoard.jspa?rapidView=235&projectKey=IQTC
  }
  Else
  {
	  ; Handle if there is an underscore instead of a hyphen, or if there is no hyphen
		storyId := StrReplace(storyId, "_", "-")
		If Not RegExMatch(storyId, "-")
		{
		  storyId := RegExReplace(storyId, "(\d+)", "-$1")
		}
		
    Run, %JiraUrl%/browse/%storyId%
  }

	
	
;---------------------------------------------------------------------------------------------------------------------
; Win+i: INBOX - Activate Outlook and goto my Inbox
;---------------------------------------------------------------------------------------------------------------------
#i::	
  ActivateOrStartMicrosoftOutlook()
  WinActivate
	SendInput ^+I    ;Ctrl-Shift-I is "Switch to Inbox" shortcut key
	Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Win+C: CALENDAR - Activate Outlook and goto my Calendar
;---------------------------------------------------------------------------------------------------------------------
#c::
  ActivateOrStartMicrosoftOutlook()
  SendInput ^2	
  Return


	
;---------------------------------------------------------------------------------------------------------------------
; Win+p: PROGRESS - Activate Outlook and goto Tasks
;---------------------------------------------------------------------------------------------------------------------
#p::
  ActivateOrStartMicrosoftOutlook()
  SendInput ^4
  Return



;---------------------------------------------------------------------------------------------------------------------
; NOTE that Outlook is whiny, and the "instant search" feature (press Ctrl+E to search your mail items) 
; refuses to run when Outlook is run as an administrator. Because we are running this AHK script as an
; administrator, we cannot simply run Outlook. Instead, we must run it as a standard user.
;---------------------------------------------------------------------------------------------------------------------
ActivateOrStartMicrosoftOutlook()
{
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
; Win+down: MINIMIZE - Minimize the active window. This overrides the existing Windows hotkey that:
;                        - First time you use it, un-maximizes (restores) the window
;                        - Second second time you use it, it minimizes the window
;---------------------------------------------------------------------------------------------------------------------
#down::
  WinMinimize, A
  Return


	
;---------------------------------------------------------------------------------------------------------------------
; Win+Shift+c: COMMAND prompt - Open a command prompt. Because we're using RunAsAdmin() in this script to run as 
;                               administrator, the command prompt has administrator privileges. I am using ConEmu as a
;                               replacement console (https://code.google.com/p/conemu-maximus5)
;---------------------------------------------------------------------------------------------------------------------
#+c::
	Run %WindowsProgramFilesFolder%\ConEmu\ConEmu64.exe -run {Shells::cmd (Admin)}
  Return

  

;---------------------------------------------------------------------------------------------------------------------
; Win+Shift+s: SURROUND - Open Surround
;---------------------------------------------------------------------------------------------------------------------
+#s::
  Run %WindowsProgramFilesFolder%\Seapine\Surround SCM\Surround SCM Client.exe
  Return
	

	
;---------------------------------------------------------------------------------------------------------------------
; Win+z: noiZe - Toggle SimplyNoise.com
;
; I'm having problems w/this site crashing Chrome, so I'll use IE for this instead. Plus, it segregates it out into a 
; separately controllable window.
;---------------------------------------------------------------------------------------------------------------------
#z::
  If Not WinExist("SimplyNoise")
    Run, iexplore.exe http://www.simplynoise.com,, Min
  Else 
	{
	  WinActivate, SimplyNoise
    WinClose, SimplyNoise
	}
	Return


	
	
	
	
	
#Include %A_ScriptDir%\My Automations Utilities.ahk
#Include %A_ScriptDir%\Slack Status Update.ahk
#Include %A_ScriptDir%\Virtual Desktop Accessor.ahk



;---------------------------------------------------------------------------------------------------------------------
; Add auto-correct to EVERY application
;   http://www.biancolo.com/articles/autocorrect/
; For some reason, I have to add this at the bottom of my script, or it interferes with the other stuff in this script
;---------------------------------------------------------------------------------------------------------------------
#Include %A_ScriptDir%\Lib\AutoCorrect.ahk

;::brian::Brian      -- doing this always capitalizes Brian-kummer@company.com
;::kummer::Kummer    -- Otherwise, it changes brian-kummer to brian-Kummer, and Jenkins can't handle that as a login
;::i::I              -- Don't want to always capitalize i (e.g. "-i")

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
::pto::PTO
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