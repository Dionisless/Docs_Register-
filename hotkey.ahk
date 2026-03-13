; ============================================================
; Горячая клавиша для запуска LanDocs Registrator
; ============================================================
; Требует AutoHotkey v1.x  (https://www.autohotkey.com/)
;
; Установка:
;   1. Скачайте и установите AutoHotkey.
;   2. Положите этот файл и LanDocs_Registrator.exe рядом.
;   3. Дважды щёлкните hotkey.ahk — значок появится в трее.
;   4. Опционально: добавьте hotkey.ahk в автозагрузку Windows.
;
; Горячая клавиша по умолчанию: Ctrl+Shift+R
;   Измените ^+r на нужную комбинацию:
;     ^  = Ctrl
;     +  = Shift
;     !  = Alt
;     #  = Win
; ============================================================

#NoEnv
#SingleInstance Force
SendMode Input

; ---- Путь к исполняемому файлу ----
; Если .exe лежит рядом с этим .ahk — оставьте как есть.
; Иначе укажите полный путь, например:
;   ExePath := "C:\Tools\LanDocs_Registrator.exe"
ExePath := A_ScriptDir . "\LanDocs_Registrator.exe"

; ---- Горячая клавиша ----
^+r::
    if !FileExist(ExePath) {
        MsgBox, 16, Ошибка, Файл не найден:`n%ExePath%
        return
    }
    Run, %ExePath%
return
