#SingleInstance Force
#Requires AutoHotkey v2.0

; === 全局变量 ===
configFile := A_ScriptDir "\files.csv"
hotkeys := Map()

; === 主程序入口 ===
LoadConfiguration()
SetupTrayMenu()

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
    successCount := 0
    
    for p in paths {
        p := Trim(p)
        if (p = "")
            continue
            
        try {
            if (InStr(p, ".exe") && !FileExist(p)) {
                ; 如果是exe文件但路径不存在，尝试直接运行（可能在PATH中）
                Run(p)
                successCount++
            } else {
                Run(p)
                successCount++
            }
            ; 添加小延迟，避免同时启动太多程序造成系统负担
            Sleep(100)
        } catch Error as e {
            ; 静默处理错误，继续尝试下一个路径
        }
    }
    
    ; 可选：显示成功打开的程序数量（调试用）
    ; TrayTip("已启动 " . successCount . " 个程序", fileName)
}

; === 显示菜单功能 ===
ShowAllItemsMenu() {
    global configFile
    
    if !FileExist(configFile) {
        MsgBox("配置文件不存在: " . configFile)
        return
    }
    
    ; 创建菜单
    itemMenu := Menu()
    itemCount := 0
    
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
            
            if (pathValid) {
                ; 添加菜单项，显示名称和快捷键
                menuText := name . " (" . hotkeyStr . ")"
                itemMenu.Add(menuText, ((filePath, fileName) => (*) => OpenPath(filePath, fileName))(path, name))
                itemCount++
            }
        }
        
        if (itemCount > 0) {
            ; 添加分隔线和重新加载选项
            itemMenu.Add()  ; 分隔线
            itemMenu.Add("重新加载配置 (Ctrl+Shift+R)", (*) => LoadConfiguration())
            
            ; 显示菜单
            itemMenu.Show()
        } else {
            MsgBox("没有找到有效的配置项")
        }
        
    } catch Error as e {
        MsgBox("读取配置文件失败: " . e.Message)
    }
}

; === 快捷键绑定 ===
; 显示菜单
^!f::ShowAllItemsMenu()

; 重新加载配置
^+r:: {
    LoadConfiguration()
    return
}

; === 托盘菜单设置 ===
SetupTrayMenu() {
    ; 设置托盘图标提示
    A_TrayMenu.Delete() ; 清除默认菜单项
    
    ; 添加主要功能菜单项
    A_TrayMenu.Add("显示所有项目 (Ctrl+Alt+F)", (*) => ShowAllItemsMenu())
    A_TrayMenu.Add() ; 分隔线
    
    ; 添加配置管理
    A_TrayMenu.Add("重新加载配置 (Ctrl+Shift+R)", (*) => ReloadConfig())
    A_TrayMenu.Add("编辑配置文件", (*) => EditConfigFile())
    A_TrayMenu.Add() ; 分隔线
    
    ; 添加帮助和退出
    A_TrayMenu.Add("关于", (*) => ShowAbout())
    A_TrayMenu.Add("退出", (*) => ExitApp())
    
    ; 设置默认双击动作
    A_TrayMenu.Default := "显示所有项目 (Ctrl+Alt+F)"
    
    ; 设置托盘提示
    A_IconTip := "File Launcher - 右键显示菜单"
}

; === 托盘菜单辅助函数 ===
ReloadConfig() {
    LoadConfiguration()
    TrayTip("配置已重新加载", "File Launcher")
}

EditConfigFile() {
    global configFile
    try {
        Run("notepad.exe " . configFile)
    } catch Error as e {
        MsgBox("无法打开配置文件: " . e.Message)
    }
}

ShowAbout() {
    MsgBox("AutoHotkey File Launcher v2.0`n`n快捷键:`n- Ctrl+Alt+F: 显示菜单`n- Ctrl+Shift+R: 重新加载配置`n`n右键托盘图标可访问此菜单", "关于 File Launcher")
}