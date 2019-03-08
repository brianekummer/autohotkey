;---------------------------------------------------------------------------------------------------------------------
;
; My Automations General Purpose Utility Functions
;
;---------------------------------------------------------------------------------------------------------------------


FindWindowEx(p_hw_parent, p_hw_child, p_class, p_title)
{
  Return, DllCall("FindWindowEx", "uint", p_hw_parent, "uint", p_hw_child, "str", p_class, "str", p_title)
}


;---------------------------------------------------------------------------------------------------------------------
; Run a DOS command. This code taken from AutoHotKey website: https://autohotkey.com/docs/commands/Run.htm
;---------------------------------------------------------------------------------------------------------------------
RunWaitOne(command)
{
  shell := ComObjCreate("WScript.Shell")      ; WshShell object: http://msdn.microsoft.com/en-us/library/aew9yb99
  exec := shell.Exec(ComSpec " /C " command)  ; Execute a single command via cmd.exe
  Return exec.StdOut.ReadAll()                ; Read and return the command's output 
}


;---------------------------------------------------------------------------------------------------------------------
; Run a DOS command
;---------------------------------------------------------------------------------------------------------------------
RunWaitHidden(cmd)
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


;---------------------------------------------------------------------------------------------------------------------
; Search the active window for an X button (XButtonImageName) to click to close the window, then restore the mouse 
; cursor to its original position.
;    - We only need to search the topmost SearchTopmostPixels (WindowY + SearchTopmostPixels) and rightmost 
;      SearchRightmostPixels (WindowX + WindowWidth - SearchRightmostPixels) pixels.
;    - XButtonImageName is assumed to be in %A_ScriptDir% (the same folder as this script)
;---------------------------------------------------------------------------------------------------------------------
ClickXToCloseActiveWindow(SearchTopmostPixels, SearchRightmostPixels, XButtonImageName)
{
  WinGetPos, WindowX, WindowY, WindowWidth, WindowHeight, A
  MouseGetPos, CurrentMouseX, CurrentMouseY
  
  XButtonImageName = %A_ScriptDir%\%XButtonImageName%
  ImageSearch, CloseButtonX, CloseButtonY, WindowX + WindowWidth - SearchRightmostPixels, WindowY, WindowX + WindowWidth, WindowY + SearchTopmostPixels, *15 %XButtonImageName%
  If ErrorLevel = 0
  {
    MouseMove, CloseButtonX + 3, CloseButtonY + 10, 0
    Click
    MouseMove, CurrentMouseX, CurrentMouseY, 0
  }
  Return 
}


;---------------------------------------------------------------------------------------------------------------------
; Determine if the mouse is over a window with the specified title. Taken from 
; http://l.autohotkey.net/docs/commands/_If.htm#Examples
;---------------------------------------------------------------------------------------------------------------------
MouseIsOver(WinTitle)
{
  MouseGetPos,,, Win
  Return WinExist(WinTitle . " ahk_id " . Win)
}

  
;---------------------------------------------------------------------------------------------------------------------
; Get the handle to the Windows System Tray
;   Based on Sean's GetTrayBar(): http://www.autohotkey.com/forum/topic17314.html
;---------------------------------------------------------------------------------------------------------------------
GetTrayBarHwnd()
{
  WinGet, ControlList, ControlList, ahk_class Shell_TrayWnd
  RegExMatch(ControlList, "(?<=ToolbarWindow32)\d+(?!.*ToolbarWindow32)", nTB)

  Loop, %nTB%
  {
    ControlGet, hWnd, hWnd,, ToolbarWindow32%A_Index%, ahk_class Shell_TrayWnd
    hParent := DllCall("GetParent", "Uint", hWnd)
    WinGetClass, sClass, ahk_id %hParent%
    If sClass != SysPager
      Continue
    Return hWnd
  }

  Return 0
}


;---------------------------------------------------------------------------------------------------------------------
; Unformat a date/time into YYYYMMDDHH24MISS format
; Taken from http://www.autohotkey.com/board/topic/65302-unformattime-inverter-for-formattime-command
; This is currently not used. It was used by my code to check Allegheny Badgers web site for schedule updates
;---------------------------------------------------------------------------------------------------------------------
UnformatTime(str, format)
{
  static pf   := "t,y,M,d,h,H,m,s"
       , MMM  := {Jan:"01", Feb:"02", Mar:"03", Apr:"04", May:"05", Jun:"06", Jul:"07", Aug:"08", Sep:"09", Oct:10, Nov:11, Dec:12}
	     , MMMM := {January:"01", February:"02", March:"03", April:"04", May:"05", June:"06", July:"07", August:"08", September:"09", October:10, November:11, December:12}
  StringCaseSense, On
  RegExMatch(str, RegExReplace(format, "\w+", "\w+"), timeStr)
  n := 0
  Loop, parse, pf, `,
  {
	  Pos :=	1, t :=	format
	  if !RegExMatch(format, A_LoopField "+", r)
	    if InStr("h,t", A_LoopField)
		    continue
	    else _1 := ""
	  else
	  {
		  While Pos := RegExMatch(format, "\w+", m, Pos + StrLen(m))
		    StringReplace, t, t, %m%, %	(m==r) ? "(\w+)" : "\w+"
		  RegExMatch(timeStr, t, _)
		  if (A_LoopField="t" && _1~="i)p")
  		  n := 12
	  }
	  res .= (A_LoopField="t") ? ""
		   : !_1 ? ((r~="y") ? "0000" : "00")
		   : (r~="y") ? ((r="y" && StrLen(_1)=1) ? 200 _1 : (StrLen(r)<4) ? ((_1>SubStr(A_YYYY,-1)) ? 19 _1 : 20 _1) : _1)
		   : (r~="M") ? ((r="MMM") ? MMM[_1] : (r="MMMM") ? MMMM[_1] : (Abs(_1)<10) ? "0" SubStr(_1,0) : _1)
		   : (r~="h") ? ((n && _1=12) ? "00" : n ? Abs(_1)+n : (Abs(_1)<10) ? "0" SubStr(_1,0) : _1)
		   : (StrLen(r)=1 && Abs(_1)<10) ? "0" SubStr(_1,0) : _1
  }
  StringCaseSense, Off
  Return res
}

  
;---------------------------------------------------------------------------------------------------------------------
; Pad a date's month and day with leading zeroes, if necessary
;---------------------------------------------------------------------------------------------------------------------
DatePad(InDate){
  Global DateCheck:= SubStr(InDate, 3, 1)
  IfNotEqual,DateCheck,/
    InDate := "0" . InDate
  DateCheck := SubStr(InDate, 6, 1)
  IfNotEqual,DateCheck,/
    InDate := SubStr(InDate, 1, 3) . "0" . SubStr(InDate, 4)

  Return InDate
}


;---------------------------------------------------------------------------------------------------------------------
; Wait for the browser cursor to change, denoting that the web page is loaded
;---------------------------------------------------------------------------------------------------------------------
WaitForBrowserCursor()
{
  Loop 
  { 
    Sleep, 1000
    If A_Cursor not contains AppStarting,Wait
      Break
  }
}


;---------------------------------------------------------------------------------------------------------------------
; Get the text that is currently selected by using the clipboard, while preserving the clipboard's current contents.
;---------------------------------------------------------------------------------------------------------------------
GetSelectedTextUsingClipboard()
{
  SelectedText =
  ClipSaved := ClipboardAll	; Save the entire clipboard to a variable of your choice
  Clipboard = 
  Send, ^c
  ClipWait, 1
  SelectedText := Clipboard
  Clipboard := ClipSaved	; Restore the original clipboard. Note the use of Clipboard (not ClipboardAll).
  ClipSaved =			    ; Free the memory in case the clipboard was very large
  
  Return SelectedText
}


;---------------------------------------------------------------------------------------------------------------------
; Add months - https://autohotkey.com/board/topic/15263-working-with-months/?p=99784
;---------------------------------------------------------------------------------------------------------------------
AddMonth(d,a) { ; d=yyyyMM, a=#months to add (can be negative)
   M := SubStr(d,5,2) + a

   Return SubStr(d,1,4) + floor((M-1)/12) . SubStr(0 . REM(M,12), -1, 2)
}
REM(n,m) { ; ~mod(n,m), but 1 <= REM(n,m) <= m, even if n < 0
   Return n > 0 ? mod(n-1,m)+1 : mod(n,m)+m
}


;---------------------------------------------------------------------------------------------------------------------
; Encode and decode URLs - https://autohotkey.com/board/topic/17367-url-encoding-and-decoding-of-special-characters
;---------------------------------------------------------------------------------------------------------------------
uriDecode(str) {
    Loop
 If RegExMatch(str, "i)(?<=%)[\da-f]{1,2}", hex)
    StringReplace, str, str, `%%hex%, % Chr("0x" . hex), All
    Else Break
 Return, str
}
UriEncode(Uri, RE="[0-9A-Za-z]"){
    VarSetCapacity(Var,StrPut(Uri,"UTF-8"),0),StrPut(Uri,&Var,"UTF-8")
    While Code:=NumGet(Var,A_Index-1,"UChar")
    Res.=(Chr:=Chr(Code))~=RE?Chr:Format("%{:02X}",Code)
    Return,Res  
}



;---------------------------------------------------------------------------------------------------------------------
;  ShellRun by Lexikos
;	   https://autohotkey.com/board/topic/72812-run-as-standard-limited-user/page-2#entry522235
;    requires: AutoHotkey_L
;    license: http://creativecommons.org/publicdomain/zero/1.0/
;
;  Credit for explaining this method goes to BrandonLive:
;  http://brandonlive.com/2008/04/27/getting-the-shell-to-run-an-application-for-you-part-2-how/
;
;  Shell.ShellExecute(File [, Arguments, Directory, Operation, Show])
;  http://msdn.microsoft.com/en-us/library/windows/desktop/gg537745
;---------------------------------------------------------------------------------------------------------------------
ShellRun(prms*)
{
    shellWindows := ComObjCreate("{9BA05972-F6A8-11CF-A442-00A0C90A8F39}")

    desktop := shellWindows.Item(ComObj(19, 8)) ; VT_UI4, SCW_DESKTOP                

    ; Retrieve top-level browser object.
    if ptlb := ComObjQuery(desktop
        , "{4C96BE40-915C-11CF-99D3-00AA004AE837}"  ; SID_STopLevelBrowser
        , "{000214E2-0000-0000-C000-000000000046}") ; IID_IShellBrowser
    {
        ; IShellBrowser.QueryActiveShellView -> IShellView
        if DllCall(NumGet(NumGet(ptlb+0)+15*A_PtrSize), "ptr", ptlb, "ptr*", psv:=0) = 0
        {
            ; Define IID_IDispatch.
            VarSetCapacity(IID_IDispatch, 16)
            NumPut(0x46000000000000C0, NumPut(0x20400, IID_IDispatch, "int64"), "int64")

            ; IShellView.GetItemObject -> IDispatch (object which implements IShellFolderViewDual)
            DllCall(NumGet(NumGet(psv+0)+15*A_PtrSize), "ptr", psv
                , "uint", 0, "ptr", &IID_IDispatch, "ptr*", pdisp:=0)

            ; Get Shell object.
            shell := ComObj(9,pdisp,1).Application

            ; IShellDispatch2.ShellExecute
            shell.ShellExecute(prms*)

            ObjRelease(psv)
        }
        ObjRelease(ptlb)
    }
}



;---------------------------------------------------------------------------------------------------------------------
;  Is Windows locked?
;---------------------------------------------------------------------------------------------------------------------
WindowsIsLocked()
{
  Return !DllCall("User32\OpenInputDesktop","int",0*0,"int",0*0,"int",0x0001L*1)
}
