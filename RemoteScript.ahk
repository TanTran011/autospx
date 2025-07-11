Global Pause := False  ; Biến để kiểm soát tạm dừng
Global NotepadOpen := False  ; Kiểm tra xem Notepad đã mở chưa

Gui, Add, Text, x10 y10, Dữ liệu đầu vào:
Gui, Add, Edit, x10 y30 w400 h300 vInputBox ReadOnly,
Gui, Add, Text, x420 y10, Kết quả:
Gui, Add, Edit, x420 y30 w400 h300 vOutputBox ReadOnly,
Gui, Add, Button, x10 y340 gLoadExcelData, Tải dữ liệu từ Excel
Gui, Add, Button, x150 y340 gStartScanning, Bắt đầu quét
Gui, Show, w850 h400, Auto Assign Delivery - SPX Express
Return

LoadExcelData:
FileSelectFile, FilePath, 3, , Chọn file Excel, *.xlsx
If (FilePath)
{
    xl := ComObjCreate("Excel.Application")
    xl.Workbooks.Open(FilePath)
    sheet := xl.ActiveSheet
    
    Data := ""
    pivotData := {}  ; Tạo object để pivot dữ liệu
    Row := 2  ; Bắt đầu từ hàng thứ 2 để bỏ qua tiêu đề
    While (sheet.Cells(Row, 39).Value != "")  ; Cột AM là cột 39
    {
        key := sheet.Cells(Row, 39).Value  ; Giá trị cột AM
        value := sheet.Cells(Row, 1).Value  ; Giá trị cột A (đối chiếu)
        
        If (!pivotData.HasKey(key))
            pivotData[key] := []
        pivotData[key].Push(value)  ; Lưu giá trị cột A theo nhóm của cột AM
        Row++
    }
    xl.Quit()
    
    For key, values in pivotData  ; Duyệt qua từng nhóm dữ liệu trùng nhau
    {
        Data .= key "`n"  ; Cột AM làm tiêu đề nhóm
        For index, val in values
            Data .= "  " val "`n"  ; Mỗi mã đơn trong cột A xuống dòng
        Data .= "`n"
    }
    GuiControl,, InputBox, %Data%
}
Return

F7::  ; Bấm F7 để chạy tự động
Gosub, StartScanning
Return

F8::  ; Bấm F8 để tạm dừng hoặc tiếp tục
Pause := !Pause
Return

StartScanning:
GuiControlGet, InputData,, InputBox
CurrentGroup := ""  ; Biến lưu nhóm hiện tại

Loop, Parse, InputData, `n, `r
{
    If (A_LoopField && Trim(A_LoopField) != "")
    {
        If (SubStr(A_LoopField, 1, 2) != "  ")  ; Nhận diện nhóm mới (cột AM)
        {
            If (CurrentGroup != "")  ; Nếu không phải lần đầu thì lưu dữ liệu
            {
                AppendHighlightedTextToNotepad(CurrentGroup)  ; Ghi vào Notepad
            }

            CurrentGroup := A_LoopField  ; Cập nhật nhóm hiện tại
            Sleep, 1000  ; Chờ trước khi di chuyển chuột
            MoveMouseSequence()
        }
        Else  ; Chỉ nhập dữ liệu từ cột A (dòng thụt vào)
        {
            While (Pause)  ; Tạm dừng nếu biến Pause = True
                Sleep, 100
            
            GuiControl,, OutputBox, %A_LoopField%  ; Hiển thị dữ liệu đang nhập
            SendInput, %A_LoopField%  ; Nhập từng dòng
            Sleep, 100
            SendInput, {Enter}  ; Mô phỏng scanner nhấn Enter
            Sleep, 500  ; Điều chỉnh tốc độ nhập
        }
    }
}

If (CurrentGroup != "")  ; Ghi chú dữ liệu của nhóm cuối cùng
{
    AppendHighlightedTextToNotepad(CurrentGroup)
}
Return


MoveMouseSequence() {
    Positions := [[109, 861], [109, 891], [408, 340], [1102, 702], [408, 349]]
    For index, pos in Positions {
        MouseMove, % pos[1], % pos[2]
        Sleep, 100  ; Chờ 300ms để ổn định vị trí chuột
        Click
        Sleep, 1000  ; Chờ 1 giây trước khi tiếp tục
    }
}

AppendHighlightedTextToNotepad(GroupName) {
    Sleep, 500

    ; 🔹 Di chuyển chuột đến tọa độ X1870 Y203 để click 1 lần
    MouseMove, 1870, 203
    Sleep, 300
    Click
    Sleep, 500  ; Chờ 0.5 giây để hệ thống xử lý
    
    ; 🔹 Tiếp tục di chuyển đến vị trí cần double-click
    MouseMove, 380, 201
    Sleep, 300
    Click, 2  ; Double click để bôi đen văn bản
    Sleep, 300
    Send, ^c  ; Sao chép nội dung bôi đen
    Sleep, 500
    ClipWait, 2  ; Chờ clipboard cập nhật nội dung

    HighlightedText := Clipboard
    If (HighlightedText != "")
    {
        If (!NotepadOpen) {
            Run, notepad.exe  ; Chỉ tạo Notepad sau nhóm đầu tiên
            Sleep, 1000  ; Chờ Notepad mở
            NotepadOpen := True
        }

        WinActivate, ahk_exe notepad.exe  ; Chuyển đến Notepad đang mở
        Sleep, 500
        Send, %GroupName% %HighlightedText% `n  ; Ghi nội dung bôi đen vào Notepad
    }
}


Return

GuiClose:
ExitApp
