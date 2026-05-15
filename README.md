# UniqueInstance-for-AHK
AutoHotkey v2 程序权限提升、单实例运行管理库，让 AHK 程序在不同系统权限下最大限度的保持单实例运行。

调试示例
====
``` autohotkey
#Requires AutoHotkey v2.0
#SingleInstance Off
#Include "UniqueInstance.ahk"

uiResult := UniqueInstance.Ensure(Map(
    "preferRunAsAdmin", true,
    "allowCoexist", true,
    "showReport", true,
    "closeWaitSeconds", 2,
    "restartArg", "/restart",
    "compiledFallbackByName", true
))

if uiResult["showReport"] && uiResult["report"] {
    MsgBox(uiResult["report"], "单实例处理报告")
}

if !uiResult["ok"] {
    ExitApp()
}

Persistent()
^Esc:: ExitApp()
```
