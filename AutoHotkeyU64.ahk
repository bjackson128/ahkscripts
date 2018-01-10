
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

:*:cc1::
    IniRead, ConfDialInVar, %A_ScriptDir%\autohotkey_secret.ini, mysection, ConfDialInKey
    IniRead, ConfCodeVar, %A_ScriptDir%\autohotkey_secret.ini, mysection, ConfCodeKey
    SendInput %ConfDialInVar%;%ConfCodeVar%{#}`t`t`t`t`t`t`t`t`tDial-in: %ConfDialInVar%`nCode: %ConfCodeVar%{#}
Return

:*:cc2::
    IniRead, ConfDialInVar, %A_ScriptDir%\autohotkey_secret.ini, mysection, ConfDialInKey
    IniRead, ConfCodeVar, %A_ScriptDir%\autohotkey_secret.ini, mysection, ConfCodeKey
    SendInput Dial-in: %ConfDialInVar%`nCode: %ConfCodeVar%{#}
Return

:*:@@::
    IniRead, PersonalEmailVar, %A_ScriptDir%\autohotkey_secret.ini, mysection, PersonalEmailKey
    SendInput %PersonalEmailVar%
Return

:*:@w::
    IniRead, WorkEmailVar, %A_ScriptDir%\autohotkey_secret.ini, mysection, WorkEmailKey
    SendInput %WorkEmailVar%
Return

;Example of standard text replace (does not work with variables)
;:*:replaceme::raw text that would be replaced
