/*修飾キー
^ = Ctrl
+ = shIft
! = Alt
# = win
VK1D = 無変換
VK1C = 変換
VKF2 = かな
sc027 = ;
sc028 = :
sc033 = ,
*/

#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%
SetKeyDelay, 200
SetTitleMatchMode,2

startup:
    MsgBox,, % "PI-NET半自動入力のショートカットキー"
            ,% "F1キー `t`t 詳細説明表示 `r"
            .  "無変換 + f `t 図番検索/員数入力 `r"
            .  "無変換 + q `t 終了"
    Return

F1::Run, % "https://github.com/ikuma-hiroyuki/pinet_input"
vk1d & q::
    MsgBox,% "終了します"
    ExitApp

; pi-netへの入力手順-------------------------------------------------
; 1. 月末在庫に数を入力
; 2. 右上の「検索/更新」をクリック
; 3. テーブル最左列のチェックボックスをオンにする
; 4. 右下の「在庫報告(画面単位)」をクリックする
; 5. 次のページに移動して1~4の手順を繰り返す

; エクセル側で必要な初期設定
; 検索モードを完全一致にする
vk1d & f::
    if WinActive("ahk_exe EXCEL.EXE") {
        Excel2PiNet()
    } else if WinActive("PI-NETシステム"){
        PiNet2Excel()
    }   
    Return              

PiNet2Excel(){
    Clipboard := GetSelectionString()
    WinActivate, % "ahk_exe EXCEL.EXE"
    if WinActive("検索と置換") {
        send,{esc}
        sleep,500
        WinActivate, % "ahk_exe EXCEL.EXE"
    }
    SendEvent, ^f
    SendEvent, ^v
    send,{Enter}
}

Excel2PiNet(){
    send,^f
    send,{esc 2}
    send,{tab 8}
    send,!hh
    sleep,200
    SendEvent,{up 3}
    send,{Enter}
    send,^c
    sleep,100
    WinActivate, % "PI-NETシステム"
    send,{tab 4}
    SendEvent,^v
}

GetSelectionString(){
    Clipboard := ""
    Send,^c
    ClipWait, 0.2
    selectionStr := Clipboard
    selectionStr := StrReplace(selectionStr, A_Space,"")
    Return selectionStr
}