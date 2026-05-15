#Requires AutoHotkey v2.0

; ======================================================================================================================
; UniqueInstance.ahk
;
; AutoHotkey v2 程序权限提升、单实例运行管理库，让 AHK 程序在不同系统权限下最大限度的保持单实例运行。
;
; ----------------------------------------------------------------------------------------------------------------------
; 核心能力
; ----------------------------------------------------------------------------------------------------------------------
; 1. 新实例启动后，优先尝试管理员方式重启自身。
; 2. 如果提权成功，由管理员权限的新实例负责关闭旧实例。
; 3. 如果提权失败，继续以当前权限尝试关闭旧实例。
; 4. 支持编译 EXE 与未编译 AHK。
; 5. 不依赖窗口标题识别实例。
; 6. 编译版优先通过 EXE 路径精确匹配。
; 7. 编译版在旧实例 EXE 路径不可读时，可按同名 EXE 回退匹配。
; 8. 未编译版通过旧实例命令行中的 .ahk 脚本路径匹配。
; 9. 排除当前新实例 PID，避免误杀自身。
; 10. 关闭旧实例时依次尝试：
;       - WM_COMMAND 65307
;       - WM_COMMAND 65405
;       - WM_CLOSE
;       - ProcessClose
; 11. 关闭失败时，可配置是否允许新旧实例共存。
; 12. 返回结构化结果，并生成统一格式的报告文本。
;
; ----------------------------------------------------------------------------------------------------------------------
; 重要设计原则
; ----------------------------------------------------------------------------------------------------------------------
; 旧实例信息字段只展示“实际从旧实例读取到的信息”。
;
; 匹配逻辑与实例信息分离：
; - processPath / scriptPath / commandLine 表示实际读取到的事实。
; - matchBasis / matchLevel / note 表示为什么判定它是旧实例。
;
; ----------------------------------------------------------------------------------------------------------------------
; 管理员提权逻辑
; ----------------------------------------------------------------------------------------------------------------------
; 本库现在采用更准确的顺序：
;
; 1. 先判断当前是否已经是管理员。
;    如果已经是管理员，无论 preferRunAsAdmin 是 true 还是 false，都直接认为无需提权。
;
; 2. 如果当前不是管理员，再判断 preferRunAsAdmin。
;    - preferRunAsAdmin=true  ：主动尝试 RunAs 提权。
;    - preferRunAsAdmin=false ：不弹 UAC，直接以当前权限继续单实例处理。
;
; ----------------------------------------------------------------------------------------------------------------------
; 官方提权风格
; ----------------------------------------------------------------------------------------------------------------------
; 编译版：
;     Run '*RunAs "' A_ScriptFullPath '" /restart'
;
; 未编译版：
;     Run '*RunAs "' A_AhkPath '" /restart "' A_ScriptFullPath '"'
;
; 注意：
; 官方风格下，未编译脚本的 /restart 位于脚本路径前面，不一定进入 A_Args。
; 因此本库的 _HasArg() 会同时检测：
;     1. A_Args
;     2. Kernel32.dll\GetCommandLineW 返回的完整命令行
;
; ----------------------------------------------------------------------------------------------------------------------
; 推荐主脚本写法
; ----------------------------------------------------------------------------------------------------------------------
; #Requires AutoHotkey v2.0
; #SingleInstance Off
; #Include "UniqueInstance.ahk"
;
; uiResult := UniqueInstance.Ensure(Map(
;     "preferRunAsAdmin", true,
;     "allowCoexist", true,
;     "showReport", false,
;     "closeWaitSeconds", 2,
;     "restartArg", "/restart",
;     "compiledFallbackByName", true
; ))
;
; if uiResult["showReport"] && uiResult["report"] {
;     MsgBox(uiResult["report"], "单实例处理报告")
; }
;
; if !uiResult["ok"] {
;     ExitApp()
; }
;
; Persistent()
; ^Esc::ExitApp()
;
; ----------------------------------------------------------------------------------------------------------------------
; 最简调用示例
; ----------------------------------------------------------------------------------------------------------------------
; #Requires AutoHotkey v2.0
; #SingleInstance Off
; #Include "UniqueInstance.ahk"
;
; uiResult := UniqueInstance.Ensure()
;
; ; 你的主程序从这里开始
;
; ======================================================================================================================

class UniqueInstance {
    ; ==================================================================================================================
    ; 对外主入口
    ; ==================================================================================================================
    static Ensure(__uiOptions := unset) {
        __uiCfg := UniqueInstance._NormalizeOptions(IsSet(__uiOptions) ? __uiOptions : Map())
        __uiResult := UniqueInstance._CreateResultSkeleton(__uiCfg)

        ; --------------------------------------------------------------------------------------------------------------
        ; 1. 优先尝试提权
        ; --------------------------------------------------------------------------------------------------------------
        __uiElevateInfo := UniqueInstance._TryRestartAsAdmin(__uiCfg)
        __uiResult["elevation"] := __uiElevateInfo["message"]

        ; 如果成功拉起管理员实例，当前普通权限实例通常会在 _TryRestartAsAdmin() 内 ExitApp()。
        ; 这里保留兜底返回逻辑。
        if __uiElevateInfo["launched"] {
            __uiResult["elevatedRestartLaunched"] := true
            __uiResult["ok"] := true
            __uiResult["report"] := UniqueInstance._BuildEarlyExitReport(__uiResult)
            return __uiResult
        }

        ; --------------------------------------------------------------------------------------------------------------
        ; 2. 放行部分窗口消息
        ;
        ; 这只影响“当前实例将来被低权限新实例关闭”的场景。
        ; 它不能反向修改已经运行的旧实例。
        ; --------------------------------------------------------------------------------------------------------------
        try DllCall("User32.dll\ChangeWindowMessageFilter", "UInt", 0x0111, "UInt", 1) ; WM_COMMAND
        try DllCall("User32.dll\ChangeWindowMessageFilter", "UInt", 0x0010, "UInt", 1) ; WM_CLOSE

        ; --------------------------------------------------------------------------------------------------------------
        ; 3. 枚举候选旧实例
        ; --------------------------------------------------------------------------------------------------------------
        __uiWmi := ComObjGet("winmgmts:")
        __uiQuery := UniqueInstance._BuildQuery(__uiResult["isCompiled"])

        for __uiProc in __uiWmi.ExecQuery(__uiQuery) {
            try {
                __uiPid := Integer(__uiProc.ProcessId)

                ; 排除当前新实例自己。
                if (__uiPid = __uiResult["currentInfo"]["pid"])
                    continue

                ; 判断是否为同一个实例。
                ; 返回对象中包含匹配级别、匹配依据、说明。
                __uiMatchInfo := UniqueInstance._GetInstanceMatchInfo(__uiProc, __uiResult["isCompiled"], __uiCfg)

                if !__uiMatchInfo["matched"]
                    continue

                __uiResult["foundOld"] := true

                ; 构造旧实例信息。
                ; 注意：旧实例字段只填写实际读到的信息，不再用当前实例路径回填。
                __uiOldInfo := UniqueInstance._CreateOldProcessInfo(__uiProc, __uiResult["isCompiled"], __uiMatchInfo)

                ; 尝试关闭旧实例。
                __uiCloseResult := UniqueInstance._CloseOldInstance(__uiPid, __uiCfg["closeWaitSeconds"])

                __uiOldInfo["closeSuccess"] := __uiCloseResult["Success"]
                __uiOldInfo["closeMethod"] := __uiCloseResult["Method"]
                __uiOldInfo["closeMessage"] := __uiCloseResult["Message"]

                if !__uiOldInfo["closeSuccess"]
                    __uiResult["closeFailedCount"] += 1

                __uiResult["oldInstances"].Push(__uiOldInfo)
            } catch as __uiErr {
                __uiErrInfo := UniqueInstance._CreateEmptyProcessInfo("旧实例")
                __uiErrInfo["closeSuccess"] := false
                __uiErrInfo["closeMethod"] := "处理异常"
                __uiErrInfo["closeMessage"] := __uiErr.Message
                __uiResult["oldInstances"].Push(__uiErrInfo)
                __uiResult["closeFailedCount"] += 1
            }
        }

        ; --------------------------------------------------------------------------------------------------------------
        ; 4. 根据 allowCoexist 决定整体 ok 状态
        ; --------------------------------------------------------------------------------------------------------------
        if (__uiResult["closeFailedCount"] > 0) && !__uiCfg["allowCoexist"]
            __uiResult["ok"] := false

        __uiResult["report"] := UniqueInstance._BuildReport(__uiResult, __uiCfg)
        return __uiResult
    }

    ; ==================================================================================================================
    ; 对外帮助文本
    ; ==================================================================================================================
    static GetUsageText() {
        __uiText := ""
        __uiText .= "UniqueInstance 使用说明`n"
        __uiText .= "========================================`n"
        __uiText .= "最简单调用：`n"
        __uiText .= 'uiResult := UniqueInstance.Ensure()`n`n'

        __uiText .= "带配置调用：`n"
        __uiText .= 'uiResult := UniqueInstance.Ensure(Map(`n'
        __uiText .= '    "preferRunAsAdmin", true,`n'
        __uiText .= '    "allowCoexist", true,`n'
        __uiText .= '    "showReport", true,`n'
        __uiText .= '    "closeWaitSeconds", 2,`n'
        __uiText .= '    "restartArg", "/restart",`n'
        __uiText .= '    "compiledFallbackByName", true`n'
        __uiText .= '))`n`n'

        __uiText .= "常见处理：`n"
        __uiText .= 'if uiResult["showReport"] && uiResult["report"] {`n'
        __uiText .= '    MsgBox(uiResult["report"], "单实例处理报告")`n'
        __uiText .= '}`n`n'
        __uiText .= 'if !uiResult["ok"] {`n'
        __uiText .= '    ExitApp()`n'
        __uiText .= '}`n`n'

        __uiText .= "参数说明：`n"
        __uiText .= "preferRunAsAdmin       : 是否在当前不是管理员时主动提权，默认 true。`n"
        __uiText .= "allowCoexist           : 关闭旧实例失败时是否允许共存，默认 true。`n"
        __uiText .= "showReport             : 是否建议调用方显示报告，默认 false。`n"
        __uiText .= "closeWaitSeconds       : 每种关闭方式后等待旧实例退出的秒数，默认 2。`n"
        __uiText .= "restartArg             : 提权重启标记参数，默认 /restart。`n"
        __uiText .= "compiledFallbackByName : 编译版路径不可读时是否按同名 EXE 回退匹配，默认 true。`n"

        return __uiText
    }

    static GetResultHelpText() {
        __uiText := ""
        __uiText .= "UniqueInstance.Ensure() 返回结果字段说明`n"
        __uiText .= "========================================`n"
        __uiText .= 'result["ok"] : 整体是否成功。若不允许共存且旧实例关闭失败，则为 false。`n'
        __uiText .= 'result["elevatedRestartLaunched"] : 是否成功拉起管理员实例。通常成功后当前实例会直接退出。`n'
        __uiText .= 'result["isCompiled"] : 当前是否编译版 EXE。`n'
        __uiText .= 'result["isAdmin"] : 当前是否管理员权限。`n'
        __uiText .= 'result["allowCoexist"] : 当前配置是否允许共存。`n'
        __uiText .= 'result["showReport"] : 当前配置是否建议显示报告。`n'
        __uiText .= 'result["elevation"] : 提权处理结果文本。`n'
        __uiText .= 'result["foundOld"] : 是否发现旧实例。`n'
        __uiText .= 'result["closeFailedCount"] : 关闭失败的旧实例数量。`n'
        __uiText .= 'result["currentInfo"] : 当前新实例信息 Map。`n'
        __uiText .= 'result["oldInstances"] : 旧实例信息数组。`n'
        __uiText .= 'result["report"] : 生成好的报告文本。`n`n'

        __uiText .= "currentInfo / oldInstances[*] 信息字段：`n"
        __uiText .= '  ["role"]         : 当前新实例 / 旧实例`n'
        __uiText .= '  ["pid"]          : PID`n'
        __uiText .= '  ["privilege"]    : 权限`n'
        __uiText .= '  ["runMode"]      : 编译 EXE / 未编译 AHK`n'
        __uiText .= '  ["name"]         : 进程名`n'
        __uiText .= '  ["processPath"]  : 进程路径，旧实例读不到则为空`n'
        __uiText .= '  ["scriptPath"]   : 脚本路径，旧实例读不到则为空`n'
        __uiText .= '  ["commandLine"]  : 命令行，旧实例读不到则为空`n'
        __uiText .= '  ["args"]         : 启动参数文本`n'
        __uiText .= '  ["matchLevel"]   : 匹配级别 current / exact / fallback / none`n'
        __uiText .= '  ["matchBasis"]   : 匹配依据`n'
        __uiText .= '  ["note"]         : 附加说明`n`n'

        __uiText .= "oldInstances[*] 额外关闭字段：`n"
        __uiText .= '  ["closeSuccess"] : 是否关闭成功`n'
        __uiText .= '  ["closeMethod"]  : 关闭方式`n'
        __uiText .= '  ["closeMessage"] : 关闭说明`n'

        return __uiText
    }

    ; ==================================================================================================================
    ; 配置处理
    ; ==================================================================================================================
    static _NormalizeOptions(__uiOptions) {
        __uiCfg := Map()
        __uiCfg["preferRunAsAdmin"] := UniqueInstance._GetOpt(__uiOptions, "preferRunAsAdmin", true)
        __uiCfg["allowCoexist"] := UniqueInstance._GetOpt(__uiOptions, "allowCoexist", true)
        __uiCfg["showReport"] := UniqueInstance._GetOpt(__uiOptions, "showReport", false)
        __uiCfg["closeWaitSeconds"] := UniqueInstance._GetOpt(__uiOptions, "closeWaitSeconds", 2)
        __uiCfg["restartArg"] := UniqueInstance._GetOpt(__uiOptions, "restartArg", "/restart")
        __uiCfg["compiledFallbackByName"] := UniqueInstance._GetOpt(__uiOptions, "compiledFallbackByName", true)
        return __uiCfg
    }

    static _GetOpt(__uiOptions, __uiKey, __uiDefaultValue) {
        try {
            if __uiOptions.Has(__uiKey)
                return __uiOptions[__uiKey]
        }

        return __uiDefaultValue
    }

    static _CreateResultSkeleton(__uiCfg) {
        __uiCurrentInfo := UniqueInstance._CreateCurrentProcessInfo()

        __uiResult := Map()
        __uiResult["ok"] := true
        __uiResult["elevatedRestartLaunched"] := false
        __uiResult["isCompiled"] := A_IsCompiled
        __uiResult["isAdmin"] := A_IsAdmin
        __uiResult["allowCoexist"] := __uiCfg["allowCoexist"]
        __uiResult["showReport"] := __uiCfg["showReport"]
        __uiResult["elevation"] := ""
        __uiResult["foundOld"] := false
        __uiResult["closeFailedCount"] := 0
        __uiResult["currentInfo"] := __uiCurrentInfo
        __uiResult["oldInstances"] := []
        __uiResult["report"] := ""

        ; 兼容旧版调用习惯，保留快捷字段。
        __uiResult["currentPid"] := __uiCurrentInfo["pid"]
        __uiResult["currentPath"] := __uiCurrentInfo["scriptPath"]

        return __uiResult
    }

    ; ==================================================================================================================
    ; 当前实例 / 旧实例信息构造
    ; ==================================================================================================================
    static _CreateEmptyProcessInfo(__uiRole := "") {
        __uiInfo := Map()
        __uiInfo["role"] := __uiRole
        __uiInfo["pid"] := 0
        __uiInfo["privilege"] := "未知"
        __uiInfo["runMode"] := ""
        __uiInfo["name"] := ""
        __uiInfo["processPath"] := ""
        __uiInfo["scriptPath"] := ""
        __uiInfo["commandLine"] := ""
        __uiInfo["args"] := ""
        __uiInfo["matchBasis"] := ""
        __uiInfo["matchLevel"] := "none"
        __uiInfo["note"] := ""
        __uiInfo["closeSuccess"] := ""
        __uiInfo["closeMethod"] := ""
        __uiInfo["closeMessage"] := ""
        return __uiInfo
    }

    static _CreateCurrentProcessInfo() {
        __uiPid := DllCall("Kernel32.dll\GetCurrentProcessId", "UInt")
        __uiInfo := UniqueInstance._CreateEmptyProcessInfo("当前新实例")

        __uiInfo["pid"] := __uiPid
        __uiInfo["privilege"] := A_IsAdmin ? "管理员" : "普通权限"
        __uiInfo["runMode"] := A_IsCompiled ? "编译 EXE" : "未编译 AHK"
        __uiInfo["scriptPath"] := A_ScriptFullPath
        __uiInfo["commandLine"] := UniqueInstance._GetCurrentCommandLine()
        __uiInfo["args"] := UniqueInstance._ArgsToText()
        __uiInfo["matchBasis"] := "当前正在启动的新实例"
        __uiInfo["matchLevel"] := "current"
        __uiInfo["note"] := "此实例会尝试关闭同一路径的旧实例。"

        try {
            __uiInfo["name"] := ProcessGetName(__uiPid)
        } catch {
            __uiInfo["name"] := A_ScriptName
        }

        try {
            __uiInfo["processPath"] := ProcessGetPath(__uiPid)
        } catch {
            __uiInfo["processPath"] := A_IsCompiled ? A_ScriptFullPath : A_AhkPath
        }

        return __uiInfo
    }

    static _CreateOldProcessInfo(__uiProc, __uiIsCompiled, __uiMatchInfo) {
        __uiInfo := UniqueInstance._CreateEmptyProcessInfo("旧实例")
        __uiPid := Integer(__uiProc.ProcessId)

        __uiInfo["pid"] := __uiPid
        __uiInfo["privilege"] := UniqueInstance._GetProcessElevationText(__uiPid)
        __uiInfo["runMode"] := __uiIsCompiled ? "编译 EXE" : "未编译 AHK"
        __uiInfo["args"] := "未知"
        __uiInfo["matchBasis"] := __uiMatchInfo["basis"]
        __uiInfo["matchLevel"] := __uiMatchInfo["level"]
        __uiInfo["note"] := __uiMatchInfo["note"]

        ; 下面这些字段只填“从旧实例实际读取到的值”。
        ; 读不到就保持为空，不使用当前实例路径回填。
        try __uiInfo["name"] := __uiProc.Name
        try __uiInfo["processPath"] := __uiProc.ExecutablePath
        try __uiInfo["commandLine"] := __uiProc.CommandLine

        if __uiIsCompiled {
            ; 编译版中，脚本本体就是 EXE。
            ; 只有旧实例 ExecutablePath 实际可读时，才把它视作旧实例脚本路径。
            ; 如果读不到，则 scriptPath 保持为空。
            __uiInfo["scriptPath"] := __uiInfo["processPath"]
        } else {
            ; 未编译版中，进程路径是 AutoHotkey.exe。
            ; 如果没有专门从旧实例命令行解析出脚本路径，则保持为空。
            ; 命令行本身会原样显示在 commandLine 字段中。
            __uiInfo["scriptPath"] := ""
        }

        return __uiInfo
    }

    ; ==================================================================================================================
    ; 提权逻辑：按官方文档风格
    ; ==================================================================================================================
    static _TryRestartAsAdmin(__uiCfg) {
        __uiInfo := Map()
        __uiInfo["launched"] := false
        __uiInfo["message"] := ""

        ; --------------------------------------------------------------------------------------------------------------
        ; 1. 先判断当前是否已经是管理员
        ;
        ; 这一步应该优先于 preferRunAsAdmin。
        ; preferRunAsAdmin 控制的是“当前不是管理员时，是否主动申请管理员权限”。
        ; 但如果当前已经是管理员，就不需要申请。
        ; --------------------------------------------------------------------------------------------------------------
        if A_IsAdmin {
            if UniqueInstance._HasArg(__uiCfg["restartArg"]) {
                __uiInfo["message"] := "当前实例已是管理员权限，并检测到重启标记；可能是提权重启后的实例。"
            } else {
                __uiInfo["message"] := "当前实例已经是管理员权限，无需提权。"
            }

            return __uiInfo
        }

        ; --------------------------------------------------------------------------------------------------------------
        ; 2. 当前不是管理员，再判断是否允许主动提权
        ;
        ; preferRunAsAdmin=false 的含义是：
        ; 当前不是管理员时，不主动弹 UAC，直接按当前权限继续运行。
        ; --------------------------------------------------------------------------------------------------------------
        if !__uiCfg["preferRunAsAdmin"] {
            __uiInfo["message"] := "当前实例不是管理员权限，且未启用主动提权；将以当前权限继续运行。"
            return __uiInfo
        }

        ; --------------------------------------------------------------------------------------------------------------
        ; 3. 如果已经带有重启标记，但仍然不是管理员，
        ;    说明提权失败、被用户拒绝或被系统策略限制。
        ; --------------------------------------------------------------------------------------------------------------
        if UniqueInstance._HasArg(__uiCfg["restartArg"]) {
            __uiInfo["message"] := "检测到重启参数，但当前仍不是管理员权限；提权失败，继续以受限模式运行。"
            return __uiInfo
        }

        ; --------------------------------------------------------------------------------------------------------------
        ; 4. 当前不是管理员，且允许主动提权，则按官方方式 RunAs
        ; --------------------------------------------------------------------------------------------------------------
        try {
            if A_IsCompiled {
                ; 官方风格：编译版直接以当前 EXE 追加重启标记。
                Run('*RunAs "' A_ScriptFullPath '" ' __uiCfg["restartArg"])
            } else {
                ; 官方风格：重启标记放在 A_AhkPath 后、脚本路径前。
                ; AutoHotkey v2 实际命令行可能显示为：
                ; AutoHotkey64.exe /restart /script "xxx.ahk"
                Run('*RunAs "' A_AhkPath '" ' __uiCfg["restartArg"] ' "' A_ScriptFullPath '"')
            }

            __uiInfo["launched"] := true
            __uiInfo["message"] := "已成功拉起管理员实例，当前普通权限实例即将退出。"

            ; 成功拉起管理员实例后，当前普通权限实例必须退出，
            ; 否则普通实例和管理员实例会同时继续执行。
            ExitApp()
        } catch as __uiErr {
            __uiInfo["message"] := "管理员权限获取失败，继续以受限模式运行。错误: " __uiErr.Message
            return __uiInfo
        }
    }

    ; ==================================================================================================================
    ; 候选旧实例查询与匹配
    ; ==================================================================================================================
    static _BuildQuery(__uiIsCompiled) {
        if __uiIsCompiled {
            return "Select ProcessId, Name, ExecutablePath, CommandLine from Win32_Process Where Name = '"
            . UniqueInstance._EscapeWmiString(A_ScriptName) "'"
        }

        return "Select ProcessId, Name, ExecutablePath, CommandLine from Win32_Process Where Name Like 'AutoHotkey%.exe'"
    }

    static _GetInstanceMatchInfo(__uiProc, __uiIsCompiled, __uiCfg) {
        __uiInfo := Map()
        __uiInfo["matched"] := false
        __uiInfo["basis"] := ""
        __uiInfo["level"] := "none"
        __uiInfo["note"] := ""

        try {
            if __uiIsCompiled {
                __uiName := ""
                __uiExePath := ""

                try __uiName := __uiProc.Name
                try __uiExePath := __uiProc.ExecutablePath

                ; ------------------------------------------------------------------------------------------------------
                ; 编译版第一优先级：路径精确匹配
                ; ------------------------------------------------------------------------------------------------------
                if (__uiExePath && UniqueInstance._PathEquals(__uiExePath, A_ScriptFullPath)) {
                    __uiInfo["matched"] := true
                    __uiInfo["basis"] := "旧实例 EXE 路径与当前程序路径一致"
                    __uiInfo["level"] := "exact"
                    __uiInfo["note"] := "已通过旧实例 ExecutablePath 精确识别。"
                    return __uiInfo
                }

                ; 路径可读但不一致，直接排除。
                if (__uiExePath && !UniqueInstance._PathEquals(__uiExePath, A_ScriptFullPath)) {
                    __uiInfo["matched"] := false
                    __uiInfo["basis"] := "旧实例 EXE 路径与当前程序路径不一致"
                    __uiInfo["level"] := "none"
                    __uiInfo["note"] := "路径可读但不匹配，已排除。"
                    return __uiInfo
                }

                ; ------------------------------------------------------------------------------------------------------
                ; 编译版第二优先级：路径不可读时按同名 EXE 回退匹配
                ;
                ; 典型场景：
                ; 普通权限新实例无法读取管理员权限旧实例的 ExecutablePath。
                ; 此时为了避免误报“未发现旧实例”，可以按同名 EXE 回退识别。
                ; ------------------------------------------------------------------------------------------------------
                if (__uiCfg["compiledFallbackByName"] && !__uiExePath && __uiName && StrLower(__uiName) = StrLower(A_ScriptName)) {
                    __uiInfo["matched"] := true
                    __uiInfo["basis"] := "无法读取旧实例 EXE 路径，按同名 EXE 回退匹配"
                    __uiInfo["level"] := "fallback"
                    __uiInfo["note"] := "这是降级匹配。旧实例部分信息可能因权限限制不可读。"
                    return __uiInfo
                }

                return __uiInfo
            } else {
                ; ------------------------------------------------------------------------------------------------------
                ; 未编译版：
                ; 进程本体是 AutoHotkey.exe，不能只比较 ExecutablePath。
                ; 必须通过 CommandLine 判断是否包含当前 .ahk 路径。
                ; ------------------------------------------------------------------------------------------------------
                __uiCmd := ""

                try __uiCmd := __uiProc.CommandLine

                if !__uiCmd {
                    __uiInfo["matched"] := false
                    __uiInfo["basis"] := "无法读取旧实例命令行"
                    __uiInfo["level"] := "none"
                    __uiInfo["note"] := "未编译模式下需要依赖旧实例 CommandLine 识别脚本路径。"
                    return __uiInfo
                }

                if UniqueInstance._CommandLineContainsScript(__uiCmd, A_ScriptFullPath) {
                    __uiInfo["matched"] := true
                    __uiInfo["basis"] := "旧实例命令行中包含当前 .ahk 脚本路径"
                    __uiInfo["level"] := "exact"
                    __uiInfo["note"] := "已通过旧实例 CommandLine 精确识别。脚本路径字段不回填，仅命令行字段展示实际读取内容。"
                    return __uiInfo
                }

                return __uiInfo
            }
        } catch as __uiErr {
            __uiInfo["matched"] := false
            __uiInfo["basis"] := "匹配判断异常"
            __uiInfo["level"] := "none"
            __uiInfo["note"] := __uiErr.Message
            return __uiInfo
        }
    }

    ; ==================================================================================================================
    ; 关闭旧实例
    ; ==================================================================================================================
    static _CloseOldInstance(__uiPid, __uiCloseWaitSeconds := 2) {
        __uiOldDetectHiddenWindows := A_DetectHiddenWindows
        DetectHiddenWindows(True)

        try {
            if !ProcessExist(__uiPid) {
                return Map(
                    "Success", true,
                    "Method", "无需关闭",
                    "Message", "旧实例在处理前已经退出。"
                )
            }

            __uiHwnd := 0

            try __uiHwnd := WinExist("ahk_class AutoHotkey ahk_pid " __uiPid)

            if __uiHwnd {
                try {
                    PostMessage(0x0111, 65307, 0, , "ahk_id " __uiHwnd)
                    Sleep(100)

                    if UniqueInstance._WaitProcessGone(__uiPid, __uiCloseWaitSeconds) {
                        return Map(
                            "Success", true,
                            "Method", "WM_COMMAND 65307",
                            "Message", "通过托盘 Exit 命令成功让旧实例退出。"
                        )
                    }
                }

                if ProcessExist(__uiPid) {
                    try {
                        PostMessage(0x0111, 65405, 0, , "ahk_id " __uiHwnd)
                        Sleep(100)

                        if UniqueInstance._WaitProcessGone(__uiPid, __uiCloseWaitSeconds) {
                            return Map(
                                "Success", true,
                                "Method", "WM_COMMAND 65405",
                                "Message", "通过 File Exit 命令成功让旧实例退出。"
                            )
                        }
                    }
                }

                if ProcessExist(__uiPid) {
                    try {
                        PostMessage(0x0010, 0, 0, , "ahk_id " __uiHwnd)
                        Sleep(100)

                        if UniqueInstance._WaitProcessGone(__uiPid, __uiCloseWaitSeconds) {
                            return Map(
                                "Success", true,
                                "Method", "WM_CLOSE",
                                "Message", "通过 WM_CLOSE 成功让旧实例退出。"
                            )
                        }
                    }
                }
            }

            if ProcessExist(__uiPid) {
                try {
                    ProcessClose(__uiPid)
                    Sleep(100)

                    if UniqueInstance._WaitProcessGone(__uiPid, __uiCloseWaitSeconds) {
                        return Map(
                            "Success", true,
                            "Method", "ProcessClose",
                            "Message", "消息退出失败，已通过 ProcessClose 强制结束旧实例。"
                        )
                    }
                } catch as __uiErr {
                    return Map(
                        "Success", false,
                        "Method", "ProcessClose 失败",
                        "Message", "尝试强制结束旧实例失败。可能是权限不足、进程受保护或系统阻止。错误: " __uiErr.Message
                    )
                }
            }

            if !ProcessExist(__uiPid) {
                return Map(
                    "Success", true,
                    "Method", "最终检测",
                    "Message", "旧实例在最终检测时已经退出。"
                )
            }

            return Map(
                "Success", false,
                "Method", "全部方式失败",
                "Message", "旧实例仍在运行。可能原因：权限不足、旧进程卡死、系统阻止结束，或消息被权限隔离拦截。"
            )
        } finally {
            DetectHiddenWindows(__uiOldDetectHiddenWindows)
        }
    }

    static _WaitProcessGone(__uiPid, __uiTimeoutSec := 2) {
        __uiEndTime := A_TickCount + Round(__uiTimeoutSec * 1000)

        while (A_TickCount < __uiEndTime) {
            if !ProcessExist(__uiPid)
                return true

            Sleep(50)
        }

        return !ProcessExist(__uiPid)
    }

    ; ==================================================================================================================
    ; 权限检测
    ; ==================================================================================================================
    static _GetProcessElevationText(__uiPid) {
        static PROCESS_QUERY_LIMITED_INFORMATION := 0x1000
        static TOKEN_QUERY := 0x0008
        static TokenElevation := 20

        __uiHProcess := 0
        __uiHToken := 0

        try {
            __uiHProcess := DllCall(
                "Kernel32.dll\OpenProcess",
                "UInt", PROCESS_QUERY_LIMITED_INFORMATION,
                "Int", false,
                "UInt", __uiPid,
                "Ptr"
            )

            if !__uiHProcess
                return "未知（无法打开进程）"

            __uiOk := DllCall(
                "Advapi32.dll\OpenProcessToken",
                "Ptr", __uiHProcess,
                "UInt", TOKEN_QUERY,
                "Ptr*", &__uiHToken
            )

            if !__uiOk || !__uiHToken
                return "未知（无法打开进程 Token）"

            __uiElevation := Buffer(4, 0)
            __uiReturnLength := 0

            __uiOk := DllCall(
                "Advapi32.dll\GetTokenInformation",
                "Ptr", __uiHToken,
                "Int", TokenElevation,
                "Ptr", __uiElevation,
                "UInt", __uiElevation.Size,
                "UInt*", &__uiReturnLength
            )

            if !__uiOk
                return "未知（无法读取 TokenElevation）"

            __uiIsElevated := NumGet(__uiElevation, 0, "UInt")
            return __uiIsElevated ? "管理员" : "普通权限"
        } catch as __uiErr {
            return "未知（检测异常: " __uiErr.Message "）"
        } finally {
            if __uiHToken
                DllCall("Kernel32.dll\CloseHandle", "Ptr", __uiHToken)

            if __uiHProcess
                DllCall("Kernel32.dll\CloseHandle", "Ptr", __uiHProcess)
        }
    }

    ; ==================================================================================================================
    ; 参数 / 命令行 / 字符串工具
    ; ==================================================================================================================
    static _HasArg(__uiTargetArg) {
        ; 第一层：检查 A_Args。
        ; 编译版或参数位于脚本路径之后时，通常能在 A_Args 中看到。
        for __uiArg in A_Args {
            if (StrLower(__uiArg) = StrLower(__uiTargetArg))
                return true
        }

        ; 第二层：检查完整命令行。
        ; 官方未编译提权写法中，/restart 位于脚本路径前面，不一定进入 A_Args。
        try {
            __uiFullCommandLine := UniqueInstance._GetCurrentCommandLine()

            if UniqueInstance._CommandLineHasToken(__uiFullCommandLine, __uiTargetArg)
                return true
        }

        return false
    }

    static _CommandLineHasToken(__uiCommandLine, __uiToken) {
        __uiCmd := StrLower(__uiCommandLine)
        __uiTok := StrLower(__uiToken)
        __uiPos := 1
        __uiTokLen := StrLen(__uiTok)
        __uiCmdLen := StrLen(__uiCmd)

        while (__uiPos := InStr(__uiCmd, __uiTok, false, __uiPos)) {
            __uiBefore := (__uiPos = 1) ? " " : SubStr(__uiCmd, __uiPos - 1, 1)
            __uiAfterPos := __uiPos + __uiTokLen
            __uiAfter := (__uiAfterPos > __uiCmdLen) ? " " : SubStr(__uiCmd, __uiAfterPos, 1)

            if UniqueInstance._IsCommandSeparator(__uiBefore) && UniqueInstance._IsCommandSeparator(__uiAfter)
                return true

            __uiPos += __uiTokLen
        }

        return false
    }

    static _IsCommandSeparator(__uiChar) {
        return (
            __uiChar = " "
            || __uiChar = "`t"
            || __uiChar = "`r"
            || __uiChar = "`n"
            || __uiChar = '"'
        )
    }

    static _ArgsToText() {
        if A_Args.Length = 0
            return "无"

        __uiText := ""

        for __uiIndex, __uiArg in A_Args {
            if (__uiIndex > 1)
                __uiText .= " "

            __uiText .= __uiArg
        }

        return __uiText
    }

    static _GetCurrentCommandLine() {
        try {
            return DllCall("Kernel32.dll\GetCommandLineW", "Str")
        } catch {
            return ""
        }
    }

    static _CommandLineContainsScript(__uiCommandLine, __uiScriptPath) {
        __uiCmd := StrLower(UniqueInstance._NormalizePath(__uiCommandLine))
        __uiTarget := StrLower(UniqueInstance._NormalizePath(__uiScriptPath))

        return InStr(__uiCmd, '"' __uiTarget '"') || InStr(__uiCmd, __uiTarget)
    }

    static _PathEquals(__uiPath1, __uiPath2) {
        return UniqueInstance._NormalizePath(__uiPath1) = UniqueInstance._NormalizePath(__uiPath2)
    }

    static _NormalizePath(__uiPath) {
        __uiPath := Trim(__uiPath)
        __uiPath := StrReplace(__uiPath, "/", "\")
        return StrLower(__uiPath)
    }

    static _EscapeWmiString(__uiStr) {
        return StrReplace(__uiStr, "'", "''")
    }

    ; ==================================================================================================================
    ; 报告生成
    ; ==================================================================================================================
    static _BuildEarlyExitReport(__uiResult) {
        __uiText := ""
        __uiText .= "单实例处理报告`n"
        __uiText .= "========================================`n"
        __uiText .= UniqueInstance._BuildProcessInfoText(__uiResult["currentInfo"])
        __uiText .= "----------------------------------------`n"
        __uiText .= "提权结果: " __uiResult["elevation"] "`n"
        __uiText .= "处理结果: 已拉起管理员实例，当前普通权限实例退出。"
        return __uiText
    }

    static _BuildReport(__uiResult, __uiCfg) {
        __uiReport := ""

        __uiReport .= "单实例处理报告`n"
        __uiReport .= "========================================`n"

        __uiReport .= UniqueInstance._BuildProcessInfoText(__uiResult["currentInfo"])
        __uiReport .= "----------------------------------------`n"

        __uiReport .= "配置与状态`n"
        __uiReport .= "----------------------------------------`n"
        __uiReport .= "提权结果: " __uiResult["elevation"] "`n"
        __uiReport .= "允许共存: " (__uiCfg["allowCoexist"] ? "是" : "否") "`n"
        __uiReport .= "建议显示报告: " (__uiCfg["showReport"] ? "是" : "否") "`n"
        __uiReport .= "编译版同名回退匹配: " (__uiCfg["compiledFallbackByName"] ? "启用" : "禁用") "`n"
        __uiReport .= "发现旧实例: " (__uiResult["foundOld"] ? "是" : "否") "`n"
        __uiReport .= "关闭失败数量: " __uiResult["closeFailedCount"] "`n"
        __uiReport .= "整体结果 ok: " (__uiResult["ok"] ? "true" : "false") "`n"
        __uiReport .= "----------------------------------------`n"

        if !__uiResult["foundOld"] {
            __uiReport .= "未发现需要关闭的旧实例。`n"
            __uiReport .= "当前新实例将正常运行。"
            return __uiReport
        }

        for __uiItem in __uiResult["oldInstances"] {
            __uiReport .= UniqueInstance._BuildProcessInfoText(__uiItem)

            __uiReport .= "----------------------------------------`n"
            __uiReport .= "关闭结果`n"
            __uiReport .= "----------------------------------------`n"
            __uiReport .= "关闭方式: " __uiItem["closeMethod"] "`n"
            __uiReport .= "是否成功: " (__uiItem["closeSuccess"] ? "是" : "否") "`n"
            __uiReport .= "说明: " __uiItem["closeMessage"] "`n"

            if !__uiItem["closeSuccess"] && __uiCfg["allowCoexist"]
                __uiReport .= "处理策略: 旧实例无法关闭，当前新实例不会退出，将允许共存。`n"

            __uiReport .= "----------------------------------------`n"
        }

        if (__uiResult["closeFailedCount"] > 0) {
            if __uiCfg["allowCoexist"] {
                __uiReport .= "旧实例处理完成，但存在未关闭实例；当前新实例将继续运行，并与其共存。"
            } else {
                __uiReport .= "旧实例处理失败，且不允许共存；调用方应根据 result[`"ok`"] 决定是否退出。"
            }
        } else {
            __uiReport .= "旧实例已全部处理完成。当前新实例将继续运行。"
        }

        return __uiReport
    }

    static _BuildProcessInfoText(__uiInfo) {
        __uiText := ""

        __uiText .= __uiInfo["role"] "`n"
        __uiText .= "----------------------------------------`n"
        __uiText .= "PID: " __uiInfo["pid"] "`n"
        __uiText .= "权限: " __uiInfo["privilege"] "`n"
        __uiText .= "运行模式: " __uiInfo["runMode"] "`n"
        __uiText .= "进程名: " __uiInfo["name"] "`n"
        __uiText .= "进程路径: " __uiInfo["processPath"] "`n"
        __uiText .= "脚本路径: " __uiInfo["scriptPath"] "`n"
        __uiText .= "命令行: " __uiInfo["commandLine"] "`n"
        __uiText .= "启动参数(A_Args): " __uiInfo["args"] "`n"
        __uiText .= "匹配级别: " __uiInfo["matchLevel"] "`n"
        __uiText .= "匹配依据: " __uiInfo["matchBasis"] "`n"
        __uiText .= "说明: " __uiInfo["note"] "`n"

        return __uiText
    }
}
