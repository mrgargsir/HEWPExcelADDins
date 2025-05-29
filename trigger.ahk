; === Configure your button coordinates here ===
itemX := 500
itemY := 120
chooseX := 735
chooseY := 120
uploadX := 870
uploadY := 120
copyX := 1020
copyY := 120
doneX := 1150
doneY := 120
firstItemX := 735   ; X of first search result
firstItemY := 650   ; Y of first search result

; === Upload process with item selection ===
UploadExcel(filePath) {
    global itemX, itemY, chooseX, chooseY, uploadX, uploadY, copyX, copyY, doneX, doneY, firstItemX, firstItemY

    ; Step 1: Copy item number from Excel
    WinActivate, ahk_class XLMAIN
    Sleep, 300
    Send, ^c
    ClipWait, 1
    itemNumber := Clipboard

    ; Step 2: Switch to Chrome
    SetTitleMatchMode, 2
    if WinExist("Haryana Engineering Works Portal - Google Chrome") {
        WinActivate
        WinWaitActive, Haryana Engineering Works Portal - Google Chrome
    } else {
        MsgBox, 48, ERROR, Website is not open.`nPlease open it manually and then try again.
        return
    }

    CoordMode, Mouse, Screen  ; Use screen coordinates
    Sleep, 300

    ; Step 3: Click "item"
    Click, %itemX%, %itemY%
    Sleep, 1000

    ; Step 4: Paste item number (without Enter)
    Send, ^v
    Sleep, 800

    ; Step 5: Click first matching result
    Click, %firstItemX%, %firstItemY%
    Sleep, 3000

    ; Step 6: Click "Choose Excel"
    Click, %chooseX%, %chooseY%
    Sleep, 1500

    ; Step 7: File selection
    SendInput, ^a
    Sleep, 100
    SendInput, %filePath%
    Sleep, 300
    Send, {Enter}
    Sleep, 1000

    ; Step 8: Upload
    Click, %uploadX%, %uploadY%
    Sleep, 7000

    ; Step 9: Copy Excel Data
    Click, %copyX%, %copyY%
    Sleep, 2500

    ; Step 10: Confirm alert
    Send, {Enter}
    Sleep, 800

    ; Step 11: Done
    Click, %doneX%, %doneY%
    Sleep, 800
}

; === Hotkeys to upload different Excel files ===
^+1::UploadExcel("C:\MRGARGSIR\Length_Breadth_Depth.xlsx")
^+2::UploadExcel("C:\MRGARGSIR\Length_Breadth.xlsx")
^+3::UploadExcel("C:\MRGARGSIR\Length.xlsx")
^+4::UploadExcel("C:\MRGARGSIR\Quantity.xlsx")
