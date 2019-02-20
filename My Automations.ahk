;---------------------------------------------------------------------------------------------------------------------
; Brian Kummer's AutoHotKey script
;
;
; Near the bottom of this script are a number of #include statements to include libraries of utility functions
;
;
; Miscellaneous
; -------------
;   - Definition of AutoHotKey keys: http://www.autohotkey.com/docs/KeyList.htm
;   - Send, SendInput modifiers: ^ = Ctrl, ! = Alt, + = Shift, # = Windows
;
;
; Known Issues
; ------------
;   - None
;
;
; To Do
; -----
;   - None
;
;
; Future Ideas
; ------------
;   - Text to speech: ComObjCreate("SAPI.SpVoice").Speak("Speak this phrase")
;
;
; Summary (1/23/2019)
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
;   Modifying application behavior
;     Typora                    Ctrl+mousewheel to zoom
;     Notepad++                 After save ahk file in Notepad++, reload the current script in AutoHotKey
;     Slack                     Typing "/lunch" gets changed to "/status :hamburger: At lunch"
;                               Typing "/mtg" gets changed to "/status :spiral_calendar_pad: In a meeting"
;                               Typing "/wfh" gets changed to "/status :house: Working remotely"
;
;   Shortcuts
;     Win+c                     outlook Calendar
;     Win+g                     gTasks Pro
;     Win+i                     outlook Inbox
;     Win+k                     slacK
;       Win+Shift+k               slack "Jump to" dialog
;     Win+m                     windows Media player
;     Win+n                     Notepad++
;       Win+Shift+n               open Notepad++, and paste the selected text into the newly opened window
;     Win+t                     Typora, with each virtual desktop having a different folder of files
;     Win+z                     noiZe- open SimplyNoise.com
;
;     Win+Ctrl+v:               Paste the clipboard as plain text
;     Win+Shift+c               Command prompt (as admin)
;     Win+Shift+a               ADP
;     Win+Shift+b               BitBucket
;     Win+Shift+g               Git Bash (as admin)
;     Win+Shift+j               JIRA
;     Win+Shift+p               Personal cloud
;     Win+Shift+v               Visual Studio 2017
;     Win+Shift+w               Wiki
;     Win+Shift+x               citriX
;     Win+Shift+y               centrifY
;
;
;   Windows Virtual Desktop Improvements
;   ------------------------------------
;   Enhancements around Windows virtual desktops, which includes assigning specific wallpapers to specific virtual 
;   desktops. Based on code: https://github.com/Ciantic/VirtualDesktopAccessor
;
;---------------------------------------------------------------------------------------------------------------------
#NoEnv
#Persistent
#WinActivateForce


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
EnvGet, PersonalCloudUrl, AHK_URL_PERSONAL_CLOUD
	
	


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
	Run, "https://www.amazon.com/Pirates-Caribbean-Stranger-Johnny-Depp/dp/B005COPWZW/ref=sr_1_1?m=A3KHMEC8P675KP&s=merchant-items&ie=UTF8&qid=1550692697&sr=1-1&keywords=pirates+caribbean+dvd",, Max
	Run, "https://www.amazon.com/Pirates-Caribbean-At-Worlds-End/dp/B000U7WV1Y/ref=sr_1_2?m=A3KHMEC8P675KP&s=merchant-items&ie=UTF8&qid=1550692697&sr=1-2&keywords=pirates+caribbean+dvd",, Max
	Run, "https://www.amazon.com/Pirates-Caribbean-Curse-Two-Disc-Collectors/dp/B00005JM5E/ref=sr_1_3?m=A3KHMEC8P675KP&s=merchant-items&ie=UTF8&qid=1550692697&sr=1-3&keywords=pirates+caribbean+dvd",, Max
	Run, "https://www.amazon.com/Pirates-Caribbean-Dead-Mans-Chest/dp/B000I0RQVI/ref=sr_1_4?m=A3KHMEC8P675KP&s=merchant-items&ie=UTF8&qid=1550692697&sr=1-4&keywords=pirates+caribbean+dvd",, Max
	Return

  ; Chuck
	Run, "https://www.amazon.com/Chuck-Complete-Various/dp/B009GYTP0W",, Max
	Run, "https://www.ebay.com/itm/Chuck-Chuck-Seasons-1-5-The-Complete-Series-New-DVD-Boxed-Set-Collectors/302168821419?epid=129869237&hash=item465aaa4eab%3Ag%3Ar1MAAOSwBY1bVoS8&_sacat=0&_nkw=chuck+tv+levi+dvd+complete&_from=R40&rt=nc&_trksid=m570.l1313",, Max
	Run, "https://www.walmart.com/ip/Chuck-The-Complete-Series-Collector-Set-DVD/21907403",, Max

	; Dragons - Race to the Edge - Release dates: 1+2 2/12/2019, 3+4 3/5/2019, 5+6 3/26/2019
	Run, "https://www.bestbuy.com/site/dragons-race-to-the-edge-seasons-1-2-dvd/34451312.p?skuId=34451312",, Max
	Run, "https://www.target.com/p/dragons-race-to-the-edge-season-1-2-dvd/-/A-54323862",, Max
	Run, "https://www.amazon.com/Dragons-Race-Seasons-Jay-Baruchel/dp/6317635579/ref=sr_1_3?ie=UTF8&qid=1550496450&sr=8-3&keywords=dragons+race+to+edge+season",, Max
	Run, "https://www.walmart.com/ip/Dragons-Race-to-the-Edge-Seasons-1-2-DVD/533584634",, Max

	Run, "https://www.bestbuy.com/site/dragons-race-to-the-edge-seasons-3-4-dvd/34475213.p?skuId=34475213",, Max
	Run, "https://www.target.com/p/dragons-race-to-the-edge-seasons-3-4-dvd/-/A-54396281",, Max
	Run, "https://www.amazon.com/Dragons-Race-Edge-Seasons/dp/B07MPK2XSV/ref=sr_1_1?ie=UTF8&qid=1550496450&sr=8-1&keywords=dragons+race+to+edge+season",, Max
	Run, "https://www.walmart.com/ip/Dragons-Race-To-The-Edge-Seasons-3-And-4-DVD/822275189",, Max

	Run, "https://www.bestbuy.com/site/dragons-race-to-the-edge-seasons-5-6-dvd/34475204.p?skuId=34475204",, Max
	Run, "https://www.target.com/p/dragons-race-to-the-edge-seasons-5-dvd/-/A-54419845",, Max
	Run, "https://www.amazon.com/Dragons-Race-Edge-Seasons/dp/B07MS59ZGR/ref=sr_1_2?ie=UTF8&qid=1550496450&sr=8-2&keywords=dragons+race+to+edge+season",, Max
	Run, "https://www.walmart.com/ip/Dragons-Race-To-The-Edge-Seasons-5-And-6-DVD/877831331",, Max
	
	; Private Eyes (Jason Priestley)
	;Run, "",, Max
	
	; Dr Seuss's The Grinch (Illumination)
	Run, "https://www.amazon.com/Illumination-Presents-Dr-Seuss-Grinch/dp/B07JYR54B7/ref=sr_1_1?ie=UTF8&qid=1550496790&sr=8-1&keywords=dvd+illumination+grinch",, Max
	Run, "https://www.walmart.com/ip/Illumination-Presents-Dr-Seuss-The-Grinch-DVD/577298400",, Max
	Run, "https://www.bestbuy.com/site/illumination-presents-dr-seuss-the-grinch-dvd-2018/6310541.p?skuId=6310541",, Max
	
	; Pirates of Caribbean
	;   2003- Curse of the Black Pearl
	;   2006- Dead Man's Chest
	;   2007- At World's End
	;   2011- On Stranger Tides
	;   2017- Dead Men Tell No Tales
	; ~$45 on amazon for each separately
	Run, "https://www.ebay.com/itm/Pirates-of-the-Caribbean-5-Movie-Collection-DVD-9-Disc-Dead-Men-Tell-No-Tales/273704386159?hash=item3fba0d7e6f:g:b40AAOSwJ5dcZ4C6",, Max
	Run, "https://www.amazon.com/s/ref=nb_sb_ss_i_1_12?url=search-alias`%3Dmovies-tv&field-keywords=dvd+pirates+of+the+caribbean&sprefix=dvd+pirates+`%2Cmovies-tv`%2C131&crid=2CVU55IXFKKPA",, Max
	Run, "https://www.target.com/s?searchTerm=dvd+pirates+of+caribbean",, Max
	Run, "https://www.bestbuy.com/site/searchpage.jsp?id=pcat17071&st=pirates+of+the+caribbean+dvd",, Max
	
	; First 4 Pirates movies as used DVDs from same store for < $7 + > $16 in shipping :-(
	Run, "https://www.amazon.com/Pirates-Caribbean-Stranger-Johnny-Depp/dp/B005COPWZW/ref=sr_1_1?m=A3KHMEC8P675KP&s=merchant-items&ie=UTF8&qid=1550692697&sr=1-1&keywords=pirates+caribbean+dvd",, Max
	Run, "https://www.amazon.com/Pirates-Caribbean-At-Worlds-End/dp/B000U7WV1Y/ref=sr_1_2?m=A3KHMEC8P675KP&s=merchant-items&ie=UTF8&qid=1550692697&sr=1-2&keywords=pirates+caribbean+dvd",, Max
	Run, "https://www.amazon.com/Pirates-Caribbean-Curse-Two-Disc-Collectors/dp/B00005JM5E/ref=sr_1_3?m=A3KHMEC8P675KP&s=merchant-items&ie=UTF8&qid=1550692697&sr=1-3&keywords=pirates+caribbean+dvd",, Max
	Run, "https://www.amazon.com/Pirates-Caribbean-Dead-Mans-Chest/dp/B000I0RQVI/ref=sr_1_4?m=A3KHMEC8P675KP&s=merchant-items&ie=UTF8&qid=1550692697&sr=1-4&keywords=pirates+caribbean+dvd",, Max
	Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Slack
;   Ctrl+mousewheel   Zoom in and out
;   Win+k             Open Slack
;   Win+Shift+k       Open Slack and go to the "Jump to" window
;
; Functionality in Slack-Status-Update.ahk
;   /xxxxx            When in Slack, replace "/lunch", "/wfh", "/mtg" with the equivalent Slack command to change my status
;---------------------------------------------------------------------------------------------------------------------
#IfWinActive ahk_group SlackStatusUpdate_WindowTitles
  ^WheelUp::   SendInput ^{+}
  ^WheelDown:: SendInput ^{-}
#IfWinActive

#k::
  OpenSlack()
  Return

+#k::
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
; Shift+mousewheel   Change system volume
;   Other keystroke combinations (Win+mousewheel, or Ctrl+mousewheel) affect the current window
;---------------------------------------------------------------------------------------------------------------------
+WheelUp::   SendInput {Volume_Up 1}
+WheelDown:: SendInput {Volume_Down 1}



;---------------------------------------------------------------------------------------------------------------------
; Extra mouse buttons
;   - XButton1 (front button) minimizes the current window
;   - XButton2 (rear button) depending on the active window, closes a WINDOW, or minimizes or closes the active APPLICATION.
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

	If RegExMatch(processNameNoExtension, "i)skype|outlook|wmplayer|slack|typora") or WinActive("gTasks Pro ahk_exe ApplicationFrameHost.exe")
		WinMinimize, A     ; Do not want to close these apps
  Else If RegExMatch(processNameNoExtension, "i)chrome|iexplore|firefox|notepad++|ssms|devenv") or WinActive("Microsoft Edge ahk_exe ApplicationFrameHost.exe")
    SendInput ^{f4}    ; Close a WINDOW/TAB/DOCUMENT
  Else
    SendInput !{f4}    ; Close the APP
  Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Notepad++
;
; When editing any AutoHotKey script in Notepad++, clicking Ctrl-S to save the script automatically forces AutoHotKey
; to reload the current script. This eliminates the need to right-click the AutoHotKey system tray icon and select
; "Reload This Script".
;
; Win+n           Open Notepad++
; Win+Shift+n     Open Notepad++ and paste the selected text into the new Notepad++ window
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

	; Save the selected text to the clipboard
  ClipSaved := ClipboardAll
  Clipboard = 
  SendInput ^c
  ClipWait, 2
	
  If (!ErrorLevel)
  {
	  If Not WinExist("- Notepad++") 
		{
		  ; Notepad++ isn't open, so start it, and wait up to 2 seconds for it to open
		  Run "%WindowsProgramFilesX86Folder%\Notepad++\notepad++.exe"
		  WinWaitActive, Notepad++,,2
    
		  WinGetTitle, notpadTitle, A
			foundPos := RegExMatch(notpadTitle, "i)new \d+ - Notepad++")
		  If (foundPos == 0)
			  ; The active Notepad++ document is not new, so we need to create a new document
				; to paste our text into
		    SendInput ^n
		}
		Else
		{
		  ; Notepad++ is already open, so activate it and create a new document
		  WinActivate, Notepad++
		  SendInput ^n
		}
		
		; Paste in our text
    SendInput ^v
    Sleep, 750
		
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
; Personal cloud
;   Win+Shift+P   Open or activate my personal cloud website in Microsoft Edge
;---------------------------------------------------------------------------------------------------------------------
+#p::
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
;   Ctrl+mousewheel   Zoom in and out
;   Win+T             Open Typora if not already open. Virtual desktops "Main" and "Personal" open different folders 
;                     of files.
;---------------------------------------------------------------------------------------------------------------------
#IfWinActive - Typora 
  ^WheelUp::   SendInput ^+{=}
  ^WheelDown:: SendInput ^+{-}
#IfWinActive

#t::
  ; Determine if Typora is running in this virtual desktop, while ignoring the instance of my to do list
	typoraIdOnThisDesktop := GetTyporaOnThisVirtualDesktop()
	If (typoraIdOnThisDesktop == 0)
	{
    n := _GetCurrentDesktopNumber()
	  If (n == 1)         ; Virtual desktop "Main"
    	Run "%WindowsProgramFilesFolder%\Typora\Typora.exe" "C:\Users\Brian-Kummer\Documents\Always Open"
		Else If (n == 2)    ; Virtual desktop "Personal"
    	Run "%WindowsProgramFilesFolder%\Typora\Typora.exe" "C:\Users\Brian-Kummer\Personal\Notes"
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
; Win+g           Activate/open gTasks Pro to handle my tasks
;---------------------------------------------------------------------------------------------------------------------
#g::
  ; Windows store apps are more complex than simple exe's
  If Not WinActive("gTasks Pro ahk_exe ApplicationFrameHost.exe") 
	{
	   Run, "C:\Users\Brian-Kummer\Personal\WindowsStoreAppLinks\gTasks Pro.lnk"
  	 WinWaitActive, gTasks Pro,,2
  }
	WinActivate, gTasks Pro
	;WinMaximize, A
  Return

	

;---------------------------------------------------------------------------------------------------------------------
; Win+Ctrl+v      Paste clipboard as plain text. https://autohotkey.com/board/topic/10412-paste-plain-text-and-copycut
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
; Win+Shift+g     Open Git. ConEmu has multiple windows, so cannot just use "WinActivate, Git Bash"
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
; Win+Shift+v     Run VISUAL STUDIO
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
;   Win+k         Use Windows Media Player to open a playlist to play K-Love. If it's already open, open a browser
;                 to K-Love's web site so I can see what song is playing. 
;                   Unused right now: http://www.klove.com/listen/player.aspx
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
; Win+Shift+a     ADP - Go to ADP website to enter my timesheet
;---------------------------------------------------------------------------------------------------------------------
+#a::
	Run, "%TimesheetUrl%",, Max
  Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Win+Shift+b     BitBucket
;---------------------------------------------------------------------------------------------------------------------
+#b::
	Run, "%BitBucketUrl%",, Max
  Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Win+Shift+w     Wiki
;---------------------------------------------------------------------------------------------------------------------
+#w::
	Run, "%WikiUrl%",, Max
  Return



;----------------------------------	-----------------------------------------------------------------------------------
; Win+Shift+y     centrifY
;---------------------------------------------------------------------------------------------------------------------
+#y::
	Run, "%CentrifyUrl%",, Max
  Return


	
;---------------------------------------------------------------------------------------------------------------------
; Win+Shift+x     citriX
;---------------------------------------------------------------------------------------------------------------------
+#x::
	Run, "%CitrixUrl%",, Max
  Return

	
  
;---------------------------------------------------------------------------------------------------------------------
; Prt Scr         Open the Windows snipping tool in rectangular mode. Image is stored in the clipboard.
;                 Shift+Prt Scr still takes a screenshot of the whole screen.
;---------------------------------------------------------------------------------------------------------------------
printscreen::
  Run snippingtool /clip
  Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Win+m           Music - Google Play Music Player or Windows Media Player
;---------------------------------------------------------------------------------------------------------------------
#m::
  ; If Google Play Music Player is running, activate that
  If WinExist("ahk_exe Google Play Music Desktop Player.exe")
	{
	  ; This process has multiple windows, we want the one with class Chrome_WidgetWin_1
	  WinActivate ahk_class Chrome_WidgetWin_1 ahk_exe i)google play music desktop player.exe
		Return
	}
	
	; Else use Windows Media Player
	If Not WinExist("Windows Media Player")
	{
	  Run "%WindowsProgramFilesX86Folder%\Windows Media Player\wmplayer.exe"
	}
  WinActivate, ahk_class WMPlayerApp
  Return  

	

;---------------------------------------------------------------------------------------------------------------------
; Win+Shift+j     JIRA- Open JIRA
;   - If the highlighted text looks like a JIRA story number (e.g. TRAN|IQTC|OCS|DA-x[xxx]), then open that story
;   - If the Git Bash window has text that looks like a JIRA story number, then open that story
;---------------------------------------------------------------------------------------------------------------------
+#j::
  regexString := "i)\b(tran|iqtc|ocs|da)([-_ ]|( - ))?\d{1,4}\b"
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
		Run, %JiraUrl%/secure/RapidBoard.jspa?rapidView=235&projectKey=IQTC&selectedIssue=IQTC-2631&sprint=1138
  }
  Else
  {
	  ; Handle if there is an underscore or space instead of a hyphen, or no hyphen
	  storyId := RegExReplace(storyId, "[\s_]", "")
		If Not RegExMatch(storyId, "-")
		{
		  storyId := RegExReplace(storyId, "(\d+)", "-$1")
		}
		
    Run, %JiraUrl%/browse/%storyId%
  }
  Return
	
	
	
;---------------------------------------------------------------------------------------------------------------------
; Win+i           INBOX - Activate Outlook and goto my Inbox
;---------------------------------------------------------------------------------------------------------------------
#i::	
  ActivateOrStartMicrosoftOutlook()
  WinActivate
	SendInput ^+I    ;Ctrl-Shift-I is "Switch to Inbox" shortcut key
	Return

	
	
;---------------------------------------------------------------------------------------------------------------------
; Win+C           CALENDAR - Activate Outlook and goto my Calendar
;---------------------------------------------------------------------------------------------------------------------
#c::
  ActivateOrStartMicrosoftOutlook()
  SendInput ^2	
  Return


	
;---------------------------------------------------------------------------------------------------------------------
; NOTE that Outlook is whiny, and the "instant search" feature (press Ctrl+E to search your mail items) 
; refuses to run when Outlook is run as an administrator. Because we are running this AHK script as an
; administrator, we cannot simply run Outlook. Instead, we must run it as a standard user.
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
; Win+down        MINIMIZE - Minimize the active window. This overrides the existing Windows hotkey that:
;                   - First time you use it, un-maximizes (restores) the window
;                   - Second second time you use it, it minimizes the window
;---------------------------------------------------------------------------------------------------------------------
#down::
  WinMinimize, A
  Return


	
;---------------------------------------------------------------------------------------------------------------------
; Win+Shift+c     COMMAND prompt - Open a command prompt. Because we're using RunAsAdmin() in this script to run as 
;                                  administrator, the command prompt has administrator privileges. I am using ConEmu
;                                  as a replacement console (https://code.google.com/p/conemu-maximus5)
;---------------------------------------------------------------------------------------------------------------------
#+c::
	Run %WindowsProgramFilesFolder%\ConEmu\ConEmu64.exe -run {Shells::cmd (Admin)}
  Return

  

;---------------------------------------------------------------------------------------------------------------------
; Win+Shift+s     SURROUND - Open Surround
;---------------------------------------------------------------------------------------------------------------------
+#s::
  Run %WindowsProgramFilesFolder%\Seapine\Surround SCM\Surround SCM Client.exe
  Return
	

	
;---------------------------------------------------------------------------------------------------------------------
; Win+z           noiZe - Toggle SimplyNoise.com
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