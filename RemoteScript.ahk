#Persistent
#SingleInstance force
CoordMode, Mouse, Screen
SetTitleMatchMode, 2

auto := false
exit_now := false

; === GUI ===
Gui, Add, Text, vStatusText w350 h30 center, 🍑 Trạng thái: Nhấn F7 để bắt đầu hái đào
Gui, Show,, 🍑 Auto Kéo Đào Về Giỏ (By Tan)
return

; === F7: Bật/tắt ===
F7::
auto := !auto
exit_now := false
if (auto)
{
    GuiControl,, StatusText, 🟢 Đang kéo... Nhấn F9 để dừng
    SetTimer, AutoDrag, 10
}
else
{
    GuiControl,, StatusText, ⏹ Đã dừng
    SetTimer, AutoDrag, Off
}
return

; === F9: Dừng khẩn cấp ===
F9::
exit_now := true
auto := false
SetTimer, AutoDrag, Off
GuiControl,, StatusText, ⛔ Đã dừng tất cả hành động!
return

; === Hàm kéo tự động ===
AutoDrag:
if (exit_now || !auto)
    return

; VÙNG CÂY ĐÀO
startX := 340
startY := 145
endX := 1133
endY := 678
step := 100

; VỊ TRÍ GIỎ
basketX := 1470
basketY := 958

; MÀU CỦA QUẢ ĐÀO
targetColor := 0xF6A072

Loop % ((endX - startX) // step + 1) {
    x := startX + (A_Index - 1) * step
    Loop % ((endY - startY) // step + 1) {
        y := startY + (A_Index - 1) * step

        if (exit_now || !auto) {
            SetTimer, AutoDrag, Off
            return
        }

        ; Kéo siêu nhanh không delay
        MouseMove, x, y, 0
        DllCall("mouse_event", "UInt", 0x02, "UInt", 0, "UInt", 0, "UInt", 0, "UPtr", 0) ; Chuột trái xuống
        MouseMove, basketX, basketY, 0
        DllCall("mouse_event", "UInt", 0x04, "UInt", 0, "UInt", 0, "UInt", 0, "UPtr", 0) ; Chuột trái nhả
    }
}

; === Kiểm tra còn quả đào ===
PixelSearch, px, py, 340, 145, 1133, 678, targetColor, 5, Fast RGB
if (ErrorLevel != 0) {
    GuiControl,, StatusText, ⏳ Hết quả đào... Đang bấm E và chờ...

    Sleep, 500
    WinActivate, ahk_exe FiveM_b2802_GTAProcess.exe
    Sleep, 200

    ; Bấm phím E
    DllCall("keybd_event", "UInt", 0x45, "UInt", 0x12, "UInt", 0, "UPtr", 0)
    Sleep, 30
    DllCall("keybd_event", "UInt", 0x45, "UInt", 0x12, "UInt", 2, "UPtr", 0)

    GuiControl,, StatusText, 🟢 Tiếp tục kéo...
}
return

; === Đóng GUI thì thoát ===
GuiClose:
ExitApp
