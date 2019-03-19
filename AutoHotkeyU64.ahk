#InstallKeybdHook

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;           OUTLOOK           ;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;


; New Outlook email
^!m:: Run "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" /c ipm.note


; New Outlook appointment
^!n:: Run "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" /c ipm.appointment

; Bring Outlook to front
^!b:: Run "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" /recycle


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;          VOL CTRL           ;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; Not currently using these because volume control built into my keyboard

; "CTRL + ALT + PLUS"  for volume up
;^!NumpadAdd::Send {Volume_Up 1}

; "CTRL + ALT + MINUS"  for volume down
;^!NumpadSub::Send {Volume_Down 1}


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;        TEXT REPLACE         ;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; Conference call dial-in (for Outlook calendar invites, with built-in tabs)
:*:cc1::
    IniRead, ConfDialInVar, %A_ScriptDir%\autohotkey_secret.ini, mysection, ConfDialInKey
    IniRead, ConfCodeVar, %A_ScriptDir%\autohotkey_secret.ini, mysection, ConfCodeKey
    SendInput %ConfDialInVar%;%ConfCodeVar%{#}`tDial-in: %ConfDialInVar%`nCode: %ConfCodeVar%{#}
Return

; Conference call dial-in for regular text boxes
:*:cc2::
    IniRead, ConfDialInVar, %A_ScriptDir%\autohotkey_secret.ini, mysection, ConfDialInKey
    IniRead, ConfCodeVar, %A_ScriptDir%\autohotkey_secret.ini, mysection, ConfCodeKey
    SendInput Dial-in: %ConfDialInVar%`nCode: %ConfCodeVar%{#}
Return

; Personal email address
:*:@@p::
    IniRead, PersonalEmailVar, %A_ScriptDir%\autohotkey_secret.ini, mysection, PersonalEmailKey
    SendInput %PersonalEmailVar%
Return

; Work email address 1
:*:@@a::
    IniRead, WorkEmailVar1, %A_ScriptDir%\autohotkey_secret.ini, mysection, WorkEmailKey1
    SendInput %WorkEmailVar1%
Return

; Work email address 2
:*:@@r::
    IniRead, WorkEmailVar2, %A_ScriptDir%\autohotkey_secret.ini, mysection, WorkEmailKey2
    SendInput %WorkEmailVar2%
Return

; Example of standard text replace (does not work with variables
; which is why I didn't use it above)
;:*:replaceme::raw text that would be replaced


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;         REMAP KEYS          ;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

Insert::Delete
F1::F2

; Remap the CapsLock key to Escape, except when Outlook is open (was having issues
; exiting myself out of draft emails by hitting CapsLock-N which closes out the draft
; without saving). Had to put this at the end of the file because otherwise for some
; reason it was messing up my text replace setup

#IfWinActive ahk_class rctrl_renwnd32 ; Outlook
{
	CapsLock::
	return
}
#IfWinNotActive ahk_class rctrl_renwnd32 ; Outlook
{
	CapsLock::Esc
	return
}
