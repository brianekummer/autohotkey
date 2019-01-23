;---------------------------------------------------------------------------------------------------------------------
; Windows Virtual Desktop
;
; Enhancements around Windows virtual desktops, which includes assigning specific wallpapers to specific virtual 
; desktops. Based on code: https://github.com/Ciantic/VirtualDesktopAccessor
;
; Windows already provides these hot keys
;   Win+Ctrl+Left          Switch to the next virtual desktop
;   Win+Ctrl+Right         Switch to the previous virtual desktop
;
; This script adds these hot keys
;   - Win+Ctrl+1, Win+Ctrl+2, etc to switch desktops
;   - Win+Ctrl+Shift+Left|Right moves the current window to that virtual desktop and switches to that virtual desktop
;
; Public functions
; ---------------
; VirtualDesktopAccessor_Initialize(virtualDesktops)
;---------------------------------------------------------------------------------------------------------------------
#NoEnv
#Persistent

#Include, lib\tooltip.ahk



;---------------------------------------------------------------------------------------------------------------------
; Define global variables 
;---------------------------------------------------------------------------------------------------------------------
global hVirtualesktopAccessor
global GetCurrentDesktopNumberProc
global GetDesktopCountProc
global GoToDesktopNumberProc
global MoveWindowToDesktopNumberProc
global RegisterPostMessageHookProc
global UnregisterPostMessageHookProc



;---------------------------------------------------------------------------------------------------------------------
; PUBLIC - Declare global variables and initialize the hooks into Windows to monitor virtual desktop changes.
;          Windows environment variables named AHK_VIRTUAL_DESKTOP_WALLPAPER_x are expected to contain a pipe-delimited 
;          string containing the desktop name and the path to the wallpaper, such as:
;            AHK_VIRTUAL_DESKTOP_WALLPAPER_1 = Main    |C:\Pictures\wallpaper1.jpg
;            AHK_VIRTUAL_DESKTOP_WALLPAPER_2 = Personal|C:\Pictures\wallpaper2.jpg
;            ...
;            AHK_VIRTUAL_DESKTOP_WALLPAPER_x = xxxxx   |C:\Pictures\wallpaperx.jpg
;---------------------------------------------------------------------------------------------------------------------
VirtualDesktopAccessor_Initialize()
{
  DetectHiddenWindows, On
  hwnd:=WinExist("ahk_pid " . DllCall("GetCurrentProcessId","Uint"))
  hwnd+=0x1000<<32

  global hVirtualDesktopAccessor := DllCall("LoadLibrary", "Str", "Virtual Desktop Accessor.dll", "Ptr") 
	global IsWindowOnCurrentVirtualDesktopProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "IsWindowOnCurrentVirtualDesktop", "Ptr")
  global IsPinnedWindowProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "IsPinnedWindow", "Ptr")
  global GetCurrentDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GetCurrentDesktopNumber", "Ptr")
  global GetDesktopCountProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GetDesktopCount", "Ptr")
  global GoToDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GoToDesktopNumber", "Ptr")
  global MoveWindowToDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "MoveWindowToDesktopNumber", "Ptr")

  global RegisterPostMessageHookProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "RegisterPostMessageHook", "Ptr")
  global UnregisterPostMessageHookProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "UnregisterPostMessageHook", "Ptr")

  DllCall(RegisterPostMessageHookProc, Int, hwnd, Int, 0x1400 + 30)
  OnMessage(0x1400 + 30, "VirtualDesktopAccessor_VirtualDesktopChanged")
}



;---------------------------------------------------------------------------------------------------------------------
; PUBLIC - Hot keys that are used
;   Win+Ctrl+x             Switches to virtual desktop #x
;   Win+Ctrl+Shift+Left    Moves current window to the virtual desktop to the left and switches to it
;   Win+Ctrl+Shift+Right   Moves current window to the virtual desktop to the right and switches to it
;
; Windows already provides these hot keys
;   Win+Ctrl+Left          Switch to the next virtual desktop
;   Win+Ctrl+Right         Switch to the previous virtual desktop
;---------------------------------------------------------------------------------------------------------------------
#^1::      _ChangeDesktop(1)
#^2::      _ChangeDesktop(2)
#^3::      _ChangeDesktop(3)
#^4::      _ChangeDesktop(4)
#^5::      _ChangeDesktop(5)
#^6::      _ChangeDesktop(6)
#^7::      _ChangeDesktop(7)
#^8::      _ChangeDesktop(8)
#^9::      _ChangeDesktop(9)
#^+Left::  _MoveAndSwitchToDesktop(_GetPreviousDesktopNumber())
#^+Right:: _MoveAndSwitchToDesktop(_GetNextDesktopNumber())


;---------------------------------------------------------------------------------------------------------------------
; PUBLIC - Is a given window on the current virtual desktop
;---------------------------------------------------------------------------------------------------------------------
IsWindowOnCurrentVirtualDesktop(hIsWindowOnCurrentVirtualDesktopProc, hWnd) 
{
  isOnCurrentVirtualDesktop := DllCall(hIsWindowOnCurrentVirtualDesktopProc, UInt, hWnd, UInt)
	Return %isOnCurrentVirtualDesktop%
}



;---------------------------------------------------------------------------------------------------------------------
; PUBLIC - Move a window to virtual desktop n
;---------------------------------------------------------------------------------------------------------------------
MoveWindowToDesktop(hWnd, n:=1) 
{
  activeHwnd := _GetCurrentWindowID()
  DllCall(MoveWindowToDesktopNumberProc, UInt, hWnd, UInt, n-1)
}


;---------------------------------------------------------------------------------------------------------------------
; Private - This function is called every time a virtual desktop changes. Because it is called via OnMessage, it runs
;           in a separate thread from the rest of the AHK script.
;---------------------------------------------------------------------------------------------------------------------
VirtualDesktopAccessor_VirtualDesktopChanged(wParam, lParam, msg, hwnd) 
{
 	new_desktop_index := lParam+1

  ; Get desktop name and path to wallpaper from Windows environment variable AHK_VIRTUAL_DESKTOP_WALLPAPER_x
	envVarName = AHK_VIRTUAL_DESKTOP_WALLPAPER_%new_desktop_index%
	EnvGet, virtualDesktop, %envVarName%
	parts := StrSplit(virtualDesktop, "|")
	new_desktop_name := parts[1]
	new_wallpaper := parts[2]
	
  ; Let AHK's auto-trim feature trim leading and trailing spaces 
	new_desktop_name = %new_desktop_name%
	new_wallpaper = %new_wallpaper%
	
  ; Switch the wallpaper
  DllCall("SystemParametersInfo", UInt, 0x14, UInt, 0, Str, new_wallpaper, UInt, 1)

  ; Display a toast message
  params := {}
  params.message := new_desktop_name
  params.lifespan := 1500
  params.position := 1
  params.fontSize := 18
  params.fontWeight := 700
  params.fontColor := "0xFFFFFF"
  params.backgroundColor := "0x1F1F1F"
  Toast(params)
}

	
	
;---------------------------------------------------------------------------------------------------------------------
; Private - Move the currently active window to desktop n, and then switch to that virtual desktop
;---------------------------------------------------------------------------------------------------------------------
_MoveAndSwitchToDesktop(n:=1) 
{
  _MoveCurrentWindowToDesktop(n)
  _ChangeDesktop(n)
}



;---------------------------------------------------------------------------------------------------------------------
; Private - Change the current virtual desktop to n
;---------------------------------------------------------------------------------------------------------------------
_ChangeDesktop(n:=1) 
{
  If (n == 0) 
	{
    n := 10
  }
  DllCall(GoToDesktopNumberProc, Int, n-1)
}



;---------------------------------------------------------------------------------------------------------------------
; Private - Get the current virtual desktop number
;---------------------------------------------------------------------------------------------------------------------
_GetCurrentDesktopNumber() 
{
  Return DllCall(GetCurrentDesktopNumberProc) + 1
}



;---------------------------------------------------------------------------------------------------------------------
; Private - Get the number of virtual desktops
;---------------------------------------------------------------------------------------------------------------------
_GetNumberOfDesktops() 
{
  Return DllCall(GetDesktopCountProc)
}



;---------------------------------------------------------------------------------------------------------------------
; Private - Get the number of the next virtual desktop
;---------------------------------------------------------------------------------------------------------------------
_GetNextDesktopNumber() 
{
  i := _GetCurrentDesktopNumber()
  i := (i == _GetNumberOfDesktops() ? i : i + 1)

  Return i
}



;---------------------------------------------------------------------------------------------------------------------
; Private - Get the number of the previous virtual desktop
;---------------------------------------------------------------------------------------------------------------------
_GetPreviousDesktopNumber() 
{
  i := _GetCurrentDesktopNumber()
  i := (i == 1 ? i : i - 1)

  Return i
}



;---------------------------------------------------------------------------------------------------------------------
; Private - Get the window ID of the currently active window
;---------------------------------------------------------------------------------------------------------------------
_GetCurrentWindowID() 
{
  WinGet, activeHwnd, ID, A
  Return activeHwnd
}



;---------------------------------------------------------------------------------------------------------------------
; Private - Move the currently active window to virtual desktop n
;---------------------------------------------------------------------------------------------------------------------
_MoveCurrentWindowToDesktop(n:=1) 
{
  activeHwnd := _GetCurrentWindowID()
  DllCall(MoveWindowToDesktopNumberProc, UInt, activeHwnd, UInt, n-1)
}


