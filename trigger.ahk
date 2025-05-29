; === Button coordinates ===
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
firstItemX := 735
firstItemY := 650

; === Upload process with item selection ===
UploadExcel(filePath) {
    global itemX, itemY, chooseX, chooseY, uploadX, uploadY, copyX, copyY, doneX, doneY, firstItemX, firstItemY

    ; Step 1: Copy item number from Excel
    WinActivate, ahk_class XLMAIN
    Sleep, 300
	Send, {F2}     ; Enters edit mode (e.g., in Excel cell)
	Sleep, 200     ; Wait briefly for edit mode to activate
	Send, ^a       ; Select all text in the field
	Sleep, 200
    	Send, ^c
	Sleep, 200
	Send, {Esc}
    
    itemNumber := Trim(Clipboard)
    Send, {Esc}
    Sleep, 500

   

    ; Step 2: Switch to Chrome
    SetTitleMatchMode, 2
    if WinExist("Haryana Engineering Works Portal") {
        WinActivate
        WinWaitActive
    } else {
        MsgBox, 48, ERROR, Website is not open.`nPlease open it manually and then try again.
        return
    }

    CoordMode, Mouse, Screen
    Sleep, 300

    ; Step 3: Click "item"
    Click, %itemX%, %itemY%
    Sleep, 1500

    ; Step 4: Paste item number (without Ctrl+V)
    Send, ^v
    Sleep, 800

    Send, {Enter}
    Sleep, 3000

    ; Step 5: Click "Choose Excel"
    Click, %chooseX%, %chooseY%
    Sleep, 1500

    ; Step 6: File selection
    if !FileExist(filePath) {
        MsgBox, 48, ERROR, File not found:`n%filePath%
        return
    }
    SendInput, ^a
    Sleep, 100
    SendInput, %filePath%
    Sleep, 1000
    Send, {Enter}
    Sleep, 1000

    ; Step 7: Click "Upload"
    Click, %uploadX%, %uploadY%
    Sleep, 7000

    ; Step 8: Click "Copy Excel Data"
    Click, %copyX%, %copyY%
    Sleep, 2500

    ; Step 9: Confirm alert
    Send, {Enter}
    Sleep, 1000

    ; Step 10: Click "Done"
    Click, %doneX%, %doneY%
    Sleep, 800

    ; Step 11: Log activity
    FileAppend, [%A_Now%] Uploaded: %filePath% with item: %itemNumber%`n, upload_log.txt
}

; === Hotkeys to upload different Excel files ===
^+1::UploadExcel("C:\MRGARGSIR\Length_Breadth_Depth.xlsx")
^+2::UploadExcel("C:\MRGARGSIR\Length_Breadth.xlsx")
^+3::UploadExcel("C:\MRGARGSIR\Length.xlsx")
^+4::UploadExcel("C:\MRGARGSIR\Quantity.xlsx")
