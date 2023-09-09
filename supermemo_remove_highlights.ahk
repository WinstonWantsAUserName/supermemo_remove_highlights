#Requires AutoHotkey v1.1.1+  ; so that the editor would recognise this script as AHK V1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#if (WinActive("ahk_class TElWind") && InStr(ControlGetFocus("A"), "Internet Explorer_Server"))
^!+h::
  KeyWait Ctrl
  KeyWait Alt
  KeyWait Shift
  ClipSaved := ClipboardAll
  LongCopy := A_TickCount, Clipboard := "", LongCopy -= A_TickCount  ; LongCopy gauges the amount of time it takes to empty the clipboard which can predict how long the subsequent ClipWait will need
  send ^c
  ClipWait, LongCopy ? 0.6 : 0.2, True
  if (ErrorLevel) {
    ToolTip("No text found.")
    return
  } else {
    ClipboardGet_HTML(t)
    RegExMatch(t, "s)<!--StartFragment-->\K.*(?=<!--EndFragment-->)", t)
  }
  t := StrReplace(t, " class=Highlight")
  send % "{text}" . t := StrReplace(t, "`r`n")
  send % "+{left " . StrLen(t) . "}^+1"
  Clipboard := ClipSaved
return

ControlGetFocus(WinTitle:="", WinText:="", ExcludeTitle:="", ExcludeText:="") {
  ControlGetFocus, v, % WinTitle, % WinText, % ExcludeTitle, % ExcludeText
  Return, v
}

ToolTip(text, perma:=false, period:=-2000, command:="", n:=20) {
  PrevCoordModeTT := A_CoordModeToolTip
  CoordMode, ToolTip, Screen
  if (command = "center") {
    x := A_ScreenWidth / 3, y := A_ScreenHeight / 2
  } else {
    x := A_ScreenWidth / 3, y := A_ScreenHeight / 5 * 4
  }
  ToolTip, % text, % x, % y, % n
  RemoveTTFunc := Func("RemoveToolTip").Bind(n)
  SetTimer, % RemoveTTFunc, off
  if (!perma)
    SetTimer, % RemoveTTFunc, % period
  CoordMode, ToolTip, % PrevCoordModeTT
}

RemoveToolTip(n:=20) {
  Tooltip,,,, % n
}

ClipboardGet_HTML( byref Data ) {  ; www.autohotkey.com/forum/viewtopic.php?p=392624#392624
  If CBID := DllCall( "RegisterClipboardFormat", Str,"HTML Format", UInt )
  If DllCall( "IsClipboardFormatAvailable", UInt,CBID ) <> 0
    If DllCall( "OpenClipboard", UInt,0 ) <> 0
    If hData := DllCall( "GetClipboardData", UInt,CBID, UInt )
        DataL := DllCall( "GlobalSize", UInt,hData, UInt )
      , pData := DllCall( "GlobalLock", UInt,hData, UInt )
      , VarSetCapacity( data, dataL * ( A_IsUnicode ? 2 : 1 ) ), StrGet := "StrGet"
      , A_IsUnicode ? Data := %StrGet%( pData, dataL, 0 )
                    : DllCall( "lstrcpyn", Str,Data, UInt,pData, UInt,DataL )
      , DllCall( "GlobalUnlock", UInt,hData )
  DllCall( "CloseClipboard" )
  Return dataL ? dataL : 0
}
