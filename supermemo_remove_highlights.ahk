#Requires AutoHotkey v1.1.1+  ; so that the editor would recognise this script as AHK V1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#if (WinActive("ahk_class TElWind"))
^!+h::
  KeyWait Ctrl
  KeyWait Alt
  KeyWait Shift
  if (IsSMEditingHTML()) {
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
  } else if (!IsSMEditingText() && ControlGet(,, "Internet Explorer_Server1", "A")) {
    MsgBox, 3,, Do you want to remove all highlights?
    if (IfMsgBox("No") || IfMsgBox("Cancel"))
      return
    send ^{f7}  ; save read point
    if (!IsSMEditingText()) {
      send ^t
      WaitSMTextFocus()
      if (!IsSMEditingHTML()) {
        ToolTip("HTML not found.")
        return
      }
    }
    if (!SMSaveHTML(, 2500)) {
      ToolTip("Time out.")
      return
    }
    if (!HTML := FileRead(HTMLPath := Copy(,,, "!{f12}fc"))) {
      ToolTip("File not found.")
      return
    }
    FileDelete % HTMLPath
    FileAppend, % StrReplace(HTML, " class=Highlight"), % HTMLPath
    send !{home}
    SMWaitFileLoad()
    send !{left}
    SMWaitFileLoad()
    send {esc}
  }
return

ControlGetFocus(WinTitle:="", WinText:="", ExcludeTitle:="", ExcludeText:="") {
  ControlGetFocus, v, % WinTitle, % WinText, % ExcludeTitle, % ExcludeText
  Return, v
}

ControlGet(Cmd:="Hwnd", Value:="", Control:="", WinTitle:="", WinText:="", ExcludeTitle:="", ExcludeText:="") {
  Control := Control ? Control : ControlGetFocus(WinTitle)
  ControlGet, v, % Cmd, % Value, % Control, % WinTitle, % WinText, % ExcludeTitle, % ExcludeText
  Return, v
}

IfContains(ByRef var, MatchList, StrCaseSense:=false) {
  PrevStringCaseSense := A_StringCaseSense
  StringCaseSense % StrCaseSense ? "on" : "off"
  If var contains %MatchList%
    ret := true
  StringCaseSense % PrevStringCaseSense
  return ret
}

IfMsgBox(ByRef ButtonName) {
  IfMsgBox, % ButtonName
    return true
}

WaitSMTextFocus(Timeout:=0) {
  StartTime := A_TickCount
  loop {
    if (IsSMEditingText()) {
      return true
    } else if (TimeOut && (A_TickCount - StartTime > TimeOut)) {
      return false
    }
  }
}

IsSMEditingText() {
  return IfContains(ControlGetFocus("A"), "Internet Explorer_Server,TMemo,TRichEdit")
}

IsSMEditingHTML() {
  return IfContains(ControlGetFocus("A"), "Internet Explorer_Server")
}

SMSaveHTML(method:=0, timeout:=0) {
  timeout := timeout ? timeout / 1000 : ""
  SMOpenNotepad(method, timeout)
  WinWaitNotActive, ahk_class TElWind,, % timeout
  WinClose, ahk_class Notepad
  WinActivate, ahk_class TElWind
  WinWaitActive, ahk_class TElWind
  return !ErrorLevel
}

SMOpenNotepad(method:=0, timeout:=0) {
  SMExitText(true, timeout)
  if (method) {
    send !{f12}fw
  } else {
    send ^+{f6}
  }
}

SMExitText(ReturnToComp:=false, timeout:=0) {
  ret := 1
  if (IsSMEditingText()) {
    if (SMHasTwoComp()) {
      send ^t
      if (ReturnToComp) {
        send !{f12}fl
      }
      ret := 2
    }
    send {esc}
    if (!SMWaitTextExit(timeout))
      return 0
  }
  return ret
}

SMHasTwoComp() {
  return ((ControlGet(,, "Internet Explorer_Server2", "ahk_class TElWind") && ControlGet(,, "Internet Explorer_Server1", "ahk_class TElWind"))
        || (ControlGet(,, "TMemo2", "ahk_class TElWind") && ControlGet(,, "TMemo1", "ahk_class TElWind")))
}

SMWaitTextExit(Timeout:=0) {
  StartTime := A_TickCount
  loop {
    if (WinActive("ahk_class TElWind") && !IsSMEditingText()) {
      return true
    ; Choices because reference could update
    } else if (TimeOut && (A_TickCount - StartTime > TimeOut)) {
      return false
    }
  }
}

FileRead(Filename) {
  FileRead, v, % Filename
  Return, v
}

SMWaitFileLoad(timeout:=0, add:="", PrepareStatBar:=true) {  ; used for reloading or waiting for an element to load
  ; Move mouse because this function requires status bar text detection
  if (PrepareStatBar)
    PrepareStatBar(1)
  match := "^(\s+)?(Priority|Int|Downloading|\(\d+ item\(s\)" . add . ")"
  if (timeout == -1) {
    ret := (WinGetText("ahk_class TStatBar") ~= match)
  } else {
    StartTime := A_TickCount
    loop {
      while (WinExist("ahk_class Internet Explorer_TridentDlgFrame"))  ; sometimes could happen on YT videos
        WinClose
      if (WinGetText("ahk_class TStatBar") ~= match) {
        ret := true
        break
      } else if (timeout && (A_TickCount - StartTime > timeout)) {
        ret := false
        break
      }
    }
  }
  if (PrepareStatBar)
    PrepareStatBar(2)
  return ret
}

PrepareStatBar(step, x:=0, y:=0) {
  static
  RestoreStatBar := false
  if (step == 1) {
    if (!WinGetText("ahk_class TStatBar")) {
      PostMessage, 0x0111, 313,,, ahk_class TElWind
      RestoreStatBar := true
    }
    PrevCoordModeMouse := A_CoordModeMouse
    CoordMode, Mouse, Screen
    MouseGetPos, xSaved, ySaved
    MouseMove, % x, % y, 0
  } else if (step == 2) {
    MouseMove, xSaved, ySaved, 0
    CoordMode, Mouse, % PrevCoordModeMouse
    if (RestoreStatBar)
      PostMessage, 0x0111, 313,,, ahk_class TElWind
  }
}

WinGetText(WinTitle:="", WinText:="", ExcludeTitle:="", ExcludeText:="") {
  WinGetText, v, % WinTitle, % WinText, % ExcludeTitle, % ExcludeText
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

; Clip() - Send and Retrieve Text Using the Clipboard
; Originally by berban - updated February 18, 2019 - modified by Winston
; https://www.autohotkey.com/boards/viewtopic.php?f=6&t=62156
Clip(Text:="", Reselect:=false, RestoreClip:=true, HTML:=false, Method:=0, KeysToSend:="") {
  if (RestoreClip)
    ClipSaved := ClipboardAll
  If (Text = "") {
    LongCopy := A_TickCount, Clipboard := "", LongCopy -= A_TickCount  ; LongCopy gauges the amount of time it takes to empty the clipboard which can predict how long the subsequent ClipWait will need
    send % KeysToSend ? KeysToSend : (Method ? "^{Ins}" : "^c")
    ClipWait, LongCopy ? 0.6 : 0.2, True
    if (!ErrorLevel) {
      if (HTML) {
        ClipboardGet_HTML(clipped)
        RegExMatch(clipped, "s)<!--StartFragment-->\K.*(?=<!--EndFragment-->)", clipped)
      } else {
        Clipped := Clipboard
      }
    }
  }
  if (RestoreClip)  ; for scripts that restore clipboard at the end
    Clipboard := ClipSaved
  If (Text = "")
    Return Clipped
}

Copy(RestoreClip:=true, HTML:=false, CopyMethod:=0, KeysToSend:="") {
  return Clip(,, RestoreClip, HTML, CopyMethod, KeysToSend)
}
