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
    Clip(StrReplace(Copy(false, true), " class=Highlight"),, false, "sm")
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
        ClipboardGet_HTML(Clipped)
        RegExMatch(Clipped, "s)<!--StartFragment-->\K.*(?=<!--EndFragment-->)", Clipped)
      } else {
        Clipped := Clipboard
      }
    }
  } Else {
    if (HTML && (HTML != "sm")) {
      SetClipboardHTML(text)
    } else {
      Clipboard := ""
      Clipboard := Text
      ClipWait
    }
    if (HTML = "sm") {
      PasteHTML()
    } else {
      send % KeysToSend ? KeysToSend : (Method ? "+{Ins}" : "^v")
      while (DllCall("GetOpenClipboardWindow"))
        sleep 1
      ; Sleep 20  ; Short sleep in case Clip() is followed by more keystrokes such as {Enter}
    }
  }
  If (Text && Reselect)
    send % "+{Left " . StrLen(ParseLineBreaks(text)) . "}"
  if (RestoreClip)  ; for scripts that restore clipboard at the end
    Clipboard := ClipSaved
  If (Text = "")
    Return Clipped
}

Copy(RestoreClip:=true, HTML:=false, CopyMethod:=0, KeysToSend:="") {
  return Clip(,, RestoreClip, HTML, CopyMethod, KeysToSend)
}

PasteHTML() {
  send {AppsKey}xp  ; Paste HTML
  while (DllCall("GetOpenClipboardWindow"))
    sleep 1
  WinWaitNotActive, ahk_class TElWind,, 0.3
  WinWaitActive, ahk_class TElWind
}

ParseLineBreaks(str) {
  if (this.SM.IsEditingHTML()) {  ; not perfect
    if (StrLen(str) != InStr(str, "`r`n") + 1) {  ; first matched `r`n not at the end
      str := RegExReplace(str, "D)(?<=[ ])\r\n$")  ; removing the very last line break if there's a space before it
      str := RegExReplace(str, "(?<![ ])\r\n$")  ; remove line breaks at end of line if there isn't a space before it
      str := StrReplace(str, "`r`n`r`n", "`n")  ; parse all paragraph tags (<P>)
    }
    str := StrReplace(str, "`r")  ; parse all line breaks (<BR>)
    str := RegExReplace(str, this.move.hr)  ; parse horizontal lines
  } else {
    str := StrReplace(str, "`r")
  }
  return str
}

; https://www.autohotkey.com/boards/viewtopic.php?t=80706
SetClipboardHTML(HtmlBody, HtmlHead:="", AltText:="") {       ; v0.67 by SKAN on D393/D42B
  Local  F, Html, pMem, Bytes, hMemHTM:=0, hMemTXT:=0, Res1:=1, Res2:=1   ; @ tiny.cc/t80706
  Static CF_UNICODETEXT:=13,   CFID:=DllCall("RegisterClipboardFormat", "Str","HTML Format")

  If ! DllCall("OpenClipboard", "Ptr",A_ScriptHwnd)
    Return 0
  Else DllCall("EmptyClipboard")

  If (HtmlBody!="")
  {
    Html     := "Version:0.9`r`nStartHTML:00000000`r`nEndHTML:00000000`r`nStartFragment"
        . ":00000000`r`nEndFragment:00000000`r`n<!DOCTYPE>`r`n<html>`r`n<head>`r`n"
              ; . HtmlHead . "`r`n</head>`r`n<body>`r`n<!--StartFragment-->`r`n"
              . HtmlHead . "`r`n</head>`r`n<body>`r`n<!--StartFragment-->"
                . HtmlBody . "<!--EndFragment-->`r`n</body>`r`n</html>"
                ; . HtmlBody . "`r`n<!--EndFragment-->`r`n</body>`r`n</html>"

    Bytes    := StrPut(Html, "utf-8")
    hMemHTM  := DllCall("GlobalAlloc", "Int",0x42, "Ptr",Bytes+4, "Ptr")
    pMem     := DllCall("GlobalLock", "Ptr",hMemHTM, "Ptr")
    StrPut(Html, pMem, Bytes, "utf-8")

    F := DllCall("Shlwapi.dll\StrStrA", "Ptr",pMem, "AStr","<html>", "Ptr") - pMem
    StrPut(Format("{:08}", F), pMem+23, 8, "utf-8")
    F := DllCall("Shlwapi.dll\StrStrA", "Ptr",pMem, "AStr","</html>", "Ptr") - pMem
    StrPut(Format("{:08}", F), pMem+41, 8, "utf-8")
    F := DllCall("Shlwapi.dll\StrStrA", "Ptr",pMem, "AStr","<!--StartFra", "Ptr") - pMem
    StrPut(Format("{:08}", F), pMem+65, 8, "utf-8")
    F := DllCall("Shlwapi.dll\StrStrA", "Ptr",pMem, "AStr","<!--EndFragm", "Ptr") - pMem
    StrPut(Format("{:08}", F), pMem+87, 8, "utf-8")

    DllCall("GlobalUnlock", "Ptr",hMemHTM)
    Res1  := DllCall("SetClipboardData", "Int",CFID, "Ptr",hMemHTM)
  }

  If (AltText!="")
  {
    Bytes    := StrPut(AltText, "utf-16")
    hMemTXT  := DllCall("GlobalAlloc", "Int",0x42, "Ptr",(Bytes*2)+8, "Ptr")
    pMem     := DllCall("GlobalLock", "Ptr",hMemTXT, "Ptr")
    StrPut(AltText, pMem, Bytes, "utf-16")
    DllCall("GlobalUnlock", "Ptr",hMemTXT)
    Res2  := DllCall("SetClipboardData", "Int",CF_UNICODETEXT, "Ptr",hMemTXT)
  }

  DllCall("CloseClipboard")
  hMemHTM := hMemHTM ? DllCall("GlobalFree", "Ptr",hMemHTM) : 0

  Return (Res1 & Res2)
}
