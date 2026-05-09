#Requires AutoHotkey v2.0
#SingleInstance Force

; If you have some script that is always running in the background, you can add the below code to it.  
; **Be sure to change the 'Run' path so it is correct! 

^+t:: ; Launch TextToolbox — copies selection then opens TTB
{
    Send "^c"          ; Copy selected text (Ctrl+C)
    Sleep 100          ; Brief pause so clipboard is populated before ClipWait checks
    ClipWait 2         ; Wait up to 2 secs for clipboard to have content
    Run "D:\AutoHotkey\MasterScript\TextToolbox\TextToolbox.exe" ; Path to your exe
}