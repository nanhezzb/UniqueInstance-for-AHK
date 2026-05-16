# UniqueInstance-for-AHK
AutoHotkey v2 程序权限提升、单实例运行管理库，让 AHK 程序在不同系统权限下最大限度的保持单实例运行。
* 解决：#SingleInstance 不能有条件地执行，配置 Force 普通权限实例无法关闭高权限实例。

使用说明
====
最简单调用：
```
uiResult := UniqueInstance.Ensure()
```
带配置调用：
```
    uiResult := UniqueInstance.Ensure(Map(
    "preferRunAsAdmin", true,
    "allowCoexist", true,
    "showReport", true,
    "closeWaitSeconds", 2,
    "restartArg", "/restart",
    "compiledFallbackByName", true
    ))
```
常见处理：
```
if uiResult["showReport"] && uiResult["report"] {
    MsgBox(uiResult["report"], "单实例处理报告")
}

if !uiResult["ok"] {
    ExitApp()
}
```
参数说明：
```
preferRunAsAdmin       : 是否在当前不是管理员时主动提权，默认 true。
allowCoexist           : 关闭旧实例失败时是否允许共存，默认 true。
showReport             : 是否建议调用方显示报告，默认 false。
closeWaitSeconds       : 每种关闭方式后等待旧实例退出的秒数，默认 2。
restartArg             : 提权重启标记参数，默认 /restart。
compiledFallbackByName : 编译版路径不可读时是否按同名 EXE 回退匹配，默认 true。
```
调试示例
====
```
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
