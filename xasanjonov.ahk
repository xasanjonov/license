; === Автоматизация ввода СТИР ===
#Persistent
#SingleInstance Force
#NoEnv
SetTitleMatchMode, 2

; === Защита скрипта паролем ===
IniFile := "C:\ProgramData\activation_device.dat"
if !FileExist(IniFile) {
    InputBox, UserPassword, Активация, Введите пароль для активации скрипта:, hide
    if (UserPassword != "xasanjanov2002X") {
        MsgBox, Неверный пароль! Скрипт закрывается.
        ExitApp
    }
    FileAppend, ACTIVATED, %IniFile%
}

CoordMode, Pixel, Screen
CoordMode, Mouse, Screen
OnExit, CloseExcel

; === Глобальные переменные ===
global ExcelFilePath := ""
global ProgressFile := ""
global CurrentRow := 1
global Pause := false
global XL := ""
global Workbook := ""
global Sheet := ""
global StirFieldX := 1469
global StirFieldY := 438
global OkButtonImage := "C:\\Users\\User\\Desktop\\script\\button.png"
global LogFile := "C:\\Users\\User\\Desktop\\script\\log.txt"

FileDelete, %LogFile%

Esc::
    Pause := true
    SaveProgress()
return

^9::
    Pause := false
    if (ExcelFilePath = "") {
        FileSelectFile, SelectedFilePath, 3,, Выберите Excel файл базы СТИР, Excel Files (*.xlsx; *.xls)
        if (SelectedFilePath = "") {
            return
        }
        ExcelFilePath := SelectedFilePath
        StringReplace, ProgressSafe, SelectedFilePath, \\, _, All
        ProgressFile := "C:\\Users\\User\\Desktop\\script\\progress_" . ProgressSafe . ".txt"
    
        LoadProgress()
    }
    ExcelProcessing()
return

ExcelProcessing() {
    global
    if !XL {
        XL := ComObjCreate("Excel.Application")
        XL.Visible := false
        Workbook := XL.Workbooks.Open(ExcelFilePath)
        Sheet := Workbook.Sheets(1)
    }

    Loop {
        if (Pause) {
            Log("PAUSED")
            SaveProgress()
            return
        }

        STIR := Sheet.Cells(CurrentRow, 1).Value
        if (STIR = "" || STIR = "NULL") {
            Log("SKIPPED", CurrentRow)
            CurrentRow++
            continue
        }

        MouseClick, Left, %StirFieldX%, %StirFieldY%
        
        Send ^a
        
        Send {Del}
        
        SendInput %STIR%
        Sleep, 150

        if (!WaitForOkButton(20000)) {
            if (CheckForSuccessStir()) {
                SoundPlay, C:\Users\User\Desktop\script\x.wav ; Успешный звук
                Log("PAUSED", CurrentRow, "СТИР успешный (ЖАМИ > 0)")
            } else {
                SoundPlay, C:\Users\User\Desktop\script\x.wav ; Звук ошибки
                                Log("PAUSED", CurrentRow, "Кнопка 'Ок' не появилась и СТИР неуспешен")
            }
            Pause := true
            SaveProgress()
            return
        }

        Log("OK", CurrentRow, STIR)
        CurrentRow++

        if (Sheet.Cells(CurrentRow, 1).Value = "" && CurrentRow > 1) {
            SaveProgress()
            FileDelete, %ProgressFile%
            CurrentRow := 1
        } else {
            SaveProgress()
        }
    }

    XL.Quit()
}

WaitForOkButton(timeoutMs) {
    global OkButtonImage
    totalTime := 0

    Loop {
        ImageSearch, x, y, 0, 0, A_ScreenWidth, A_ScreenHeight, %OkButtonImage%
        if (ErrorLevel = 0) {
            MouseClick, Left, %x%, %y%
            return true
        }
        
        totalTime += 200
        if (totalTime >= timeoutMs)
            break
    }
    return false
}

CheckForSuccessStir() {
    capturePath := "D:\\Capture2Text\\Capture2Text.exe"
    if !FileExist(capturePath)
        return false
    cmd := capturePath . " -screen-rect 0,0," . A_ScreenWidth . "," . A_ScreenHeight
    RunWait, %cmd%, , Hide
    if (ErrorLevel) {
        Log("ERROR", "", "Capture2Text execution failed")
        return false
    }
    ClipWait, 1
    text := Clipboard
    return RegExMatch(text, "ЖАМИ:\s*\\K[1-9]\\d*")
}

LoadProgress() {
    global ProgressFile, CurrentRow
    if (FileExist(ProgressFile)) {
        FileRead, CurrentRow, %ProgressFile%
        if (CurrentRow = "" || CurrentRow < 1)
            CurrentRow := 1
    } else {
        CurrentRow := 1
    }
}

SaveProgress() {
    global ProgressFile, CurrentRow
    FileDelete, %ProgressFile%
    FileAppend, %CurrentRow%, %ProgressFile%
}

Log(type, row := "", msg := "") {
    global LogFile
    FormatTime, time,, yyyy-MM-dd HH:mm:ss
    FileAppend, [%time%] [%type%] ROW: %row% MSG: %msg%`n, %LogFile%
}

CloseExcel:
    SaveProgress()
    if IsObject(XL) {
        try XL.Quit()
        XL := ""
    }
ExitApp