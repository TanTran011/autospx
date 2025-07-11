#Persistent
#NoEnv
SetBatchLines, -1
Paused := False

; === Kiểm tra quyền truy cập ===
if !CheckAuthorization() {
    MsgBox, 16, Cảnh báo, Bạn không có quyền sử dụng phần mềm này!
    ExitApp
}

; === Giao diện chính ===
Gui, Add, Edit, vUserInput w300 R10 Multi gUpdateLineCount  ; Ô nhập dữ liệu, có sự kiện gõ
Gui, Add, Text, vLineCount w150, Số dòng: 0                 ; Hiển thị số dòng (đã tăng chiều rộng)
Gui, Add, Edit, vSpeed w50 Number, 500                      ; Ô nhập tốc độ (ms)
Gui, Add, Button, gClearOutput, Xóa Output                 ; Nút xóa output
Gui, Add, Edit, vOutput w300 R10 Multi ReadOnly            ; Ô hiển thị kết quả
Gui, Show, , SPX Express
return

; === Cập nhật số dòng mỗi khi gõ vào ===
UpdateLineCount:
GuiControlGet, UserInput,, UserInput
LineCount := 0
Loop, Parse, UserInput, `n, `r
{
    if (Trim(A_LoopField) != "")
        LineCount++
}
GuiControl,, LineCount, Số dòng: %LineCount%
return

; === Nhấn F7 để bắt đầu nhập từng dòng ===
F7::
Gui, Submit, NoHide

OutputText := ""  ; Xóa nội dung cũ
Loop, Parse, UserInput, `n, `r
{
    if (A_LoopField = "")  ; Bỏ qua dòng trống
        Continue
    while Paused  ; Tạm dừng nếu cần
        Sleep, 50
    SendInput, %A_LoopField%
    Send, {Enter}
    OutputText .= A_LoopField . "`n"
    GuiControl,, Output, %OutputText%
    Sleep, %Speed%
}
return

; === Nhấn F8 để tạm dừng hoặc tiếp tục ===
F8:: 
Paused := !Paused
return

ClearOutput:
GuiControl,, Output,
GuiControl,, LineCount, Số dòng: 0
return

GuiClose:
ExitApp

; === Hàm kiểm tra quyền truy cập ===
CheckAuthorization() {
    AuthorizedUsers := GetAuthorizedUsers()
    EnvGet, CurrentUser, USERNAME
    MsgBox, User Name: %CurrentUser%
    
    if (!IsObject(AuthorizedUsers) || AuthorizedUsers.Length() = 0) {
        MsgBox, Lỗi! Không tải được danh sách user hoặc danh sách rỗng.
        return False
    }
    
    for index, user in AuthorizedUsers {
        if (Trim(user) = Trim(CurrentUser)) {
            MsgBox, Bạn có quyền truy cập!
            return True
        }
    }
    
    MsgBox, Vui lòng liên hệ Tân để được hỗ trợ!
    return False
}

; === Hàm lấy danh sách user từ Google Sheets ===
GetAuthorizedUsers() {
    URL := "https://script.google.com/macros/s/AKfycby0Z-MPEwaL7NAW4ctnIKBg7g5QRo18XaZbEHs_Ym4o9lPhaysegDJS4DL5Y9RTRYhg5w/exec"
    WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")

    try {
        WebRequest.Open("GET", URL, False)
        WebRequest.Send()
        Users := WebRequest.ResponseText
        if (Users = "" || InStr(Users, "error")) {
            MsgBox, Lỗi! Không nhận được dữ liệu hợp lệ.
            return []
        }

        ; Xử lý chuỗi danh sách user
        Users := RegExReplace(Users, "\[|\]|\r|\n|""", "")
        Users := RegExReplace(Users, ",\s+", ",")
        Users := Trim(Users)

        UserList := StrSplit(Users, ",")
        if (!IsObject(UserList) || UserList.Length() = 0) {
            MsgBox, Danh sách user rỗng hoặc lỗi!
            return []
        }

        return UserList
    } catch {
        return []
    }
}
