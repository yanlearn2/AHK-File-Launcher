#SingleInstance Force
#Requires AutoHotkey v2.0

; === 全局变量 ===
configFile := A_ScriptDir "\files.csv"
hotkeys := Map()

; === 主程序入口 ===
LoadConfiguration()

; === 配置加载函数 ===
LoadConfiguration() {
    global hotkeys, configFile
    
    ; 清空数据
    hotkeys := Map()
    
    ; 检查配置文件
    if !FileExist(configFile) {
        return
    }
    
    try {
        fileContent := FileRead(configFile, "UTF-8")
        lines := StrSplit(fileContent, "`n", "`r")
        
        for line in lines {
            line := Trim(line)
            if (line = "" || SubStr(line, 1, 1) = "#")
                continue

            parts := StrSplit(line, ",")
            if (parts.Length < 3)
                continue
            
            name := Trim(parts[1])
            path := Trim(parts[2])
            hotkeyStr := Trim(parts[3])
            
            ; 验证路径
            pathValid := false
            for p in StrSplit(path, ";") {
                p := Trim(p)
                if (p != "" && (FileExist(p) || p = "notepad.exe" || InStr(p, ".exe"))) {
                    pathValid := true
                    break
                }
            }
            
            ; 注册快捷键
            if (hotkeyStr != "" && pathValid) {
                try {
                    RegisterHotkey(hotkeyStr, path, name)
                } catch Error as e {
                    ; 静默处理错误，不显示提醒
                }
            }
        }
    } catch Error as e {
        ; 静默处理错误
    }
}

; === 快捷键转换函数 ===
ConvertToAHKHotkey(hotkeyStr) {
    ; 将快捷键字符串转换为AutoHotkey格式
    ahkHotkey := hotkeyStr
    ahkHotkey := StrReplace(ahkHotkey, "Ctrl+", "^")
    ahkHotkey := StrReplace(ahkHotkey, "Alt+", "!")
    ahkHotkey := StrReplace(ahkHotkey, "Shift+", "+")
    ahkHotkey := StrReplace(ahkHotkey, "Win+", "#")
    
    ; 只对字母转换为小写，保留数字和其他字符
    result := ""
    chars := StrSplit(ahkHotkey)
    for i, char in chars {
        charCode := Ord(char)
        if (charCode >= 65 && charCode <= 90) {  ; A-Z
            result .= Chr(charCode + 32)  ; 转换为小写
        } else {
            result .= char
        }
    }
    
    return result
}

; === 快捷键注册函数 ===
RegisterHotkey(hotkeyStr, filePath, fileName) {
    ; 将快捷键字符串转换为AutoHotkey格式
    ahkHotkey := ConvertToAHKHotkey(hotkeyStr)
    
    ; 注册热键
    try {
        Hotkey(ahkHotkey, ((path, name) => (*) => OpenPath(path, name))(filePath, fileName))
    } catch Error as e {
        ; 静默处理错误
    }
}

; === 文件打开函数 ===
OpenPath(filePath, fileName) {
    paths := StrSplit(filePath, ";")
    
    for p in paths {
        p := Trim(p)
        if (p = "")
            continue
            
        try {
            if (InStr(p, ".exe") && !FileExist(p)) {
                ; 如果是exe文件但路径不存在，尝试直接运行（可能在PATH中）
                Run(p)
            } else {
                Run(p)
            }
            return  ; 成功打开一个就退出
        } catch Error as e {
            ; 静默处理错误，继续尝试下一个路径
        }
    }
}

; === 重新加载配置快捷键 ===
^+r:: {
    LoadConfiguration()
    return
}