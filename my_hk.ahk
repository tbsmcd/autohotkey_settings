/*
title: Window Titles https://ahkscript.github.io/ja/docs/v2/misc/WinTitle.htm
cmd: アプリを起動するためのコマンド
*/

switch_app_status(title, cmd:="") {
	if WinExist(title) {
		if WinActive() {
			WinMoveBottom
			if WinExist( , , title) {
				WinActivate ; `title` 以外の見つかったウインドウがアクティブになる
			}
		} else {
			WinMoveTop
			WinActivate
		}
	} else {
        cmd := String(cmd)
        if (cmd != "") {
		    Run cmd
        } else {
            MsgBox "このアプリケーションは起動していません"
        }
	}
}


/*
===================================
ショートカット
===================================
*/
; mac 風にスクリーンショット
^+4::Send '{PrintScreen}'
/*
 Esc で IMEオフ + Esc 送出
 Google IME の設定で F12 に「IME無効化」を割り当てている
 $ は Esc の無限ループ防止のため
*/
$Esc::Send '{Esc}{F12}'

/*
===================================
ホットキー
アプリをトグルでアクティブ/インアクティブ
アプリが起動していない場合は起動
===================================
*/

/*
Ctrl + Enter Windows Terminal
*/
^Enter::
{
	switch_app_status("ahk_exe WindowsTerminal.exe", "wt -p Ubuntu")
}


/*
Ctrl + \ エクスプローラー
*/
^\::
{
	switch_app_status("ahk_class CabinetWClass", "explorer")
}


/*
Ctrl + 0 VSCode
*/
^0::
{
	switch_app_status("ahk_exe Code.exe", "Code")
}


/*
===================================
GUI
アプリをトグルでアクティブ/インアクティブ
アプリが起動していない場合は起動
起動コマンドが指定されていない場合は
アプリ未起動時は button が disabled になる
===================================
*/
a_outlook(*) {
    switch_app_status("ahk_exe outlook.exe")
    WinClose("Select an app")
}
a_teams(*) {
    switch_app_status("ahk_exe ms-teams.exe")
    WinClose("Select an app")
}
a_edge(*) {
    switch_app_status("ahk_exe msedge.exe")
    WinClose("Select an app")
}
a_chrome(*) {
    switch_app_status("ahk_exe chrome.exe")
    WinClose("Select an app")
}
a_code(*) {
    switch_app_status("ahk_exe Code.exe", "code")
    WinClose("Select an app")
}
exit_gui(*) {
    If WinExist("Select an app") {
		WinClose
    }
}

/*
Alt 連打で GUI を起動
*/
Alt::
{
    if (A_ThisHotkey == A_PriorHotkey && A_TimeSincePriorHotkey <= 400) {
        If WinExist("Select an app") {
            WinActivate
            return
        }
        MyGui := Gui(, "Select an app")
        MyGui.SetFont("s14")
        apps := [
            Map("app_name", "Outlook", "func", a_outlook, "win_title", "ahk_exe outlook.exe"),
            Map("app_name", "Teams", "func", a_teams, "win_title", "ahk_exe ms-teams.exe"),
            Map("app_name", "Edge", "func", a_edge, "win_title", "ahk_exe msedge.exe"),
            Map("app_name", "Chrome", "func", a_chrome, "win_title", "ahk_exe chrome.exe"),
            Map("app_name", "VSCode", "func", a_code, "win_title", "ahk_exe Code.exe")
        ]
        for (v in apps) {
            opt_disabled := ""
            name_disabled := ""
            if not WinExist(v["win_title"]) {
                opt_disabled := "disabled "
                name_disabled := " (not started)"
            }
            MyGui.Add("Button", opt_disabled "W250", v["app_name"] name_disabled).OnEvent("click", v["func"])
        }
        MyGui.OnEvent("Escape", exit_gui)
        MyGui.show()
    }
}

/*
===================================
ローマ字を再変換
右シフト 2連打で発火 Interval <= 400ms
選択中のテキストが `半角英数-~のみ1文字以上` の場合だけ有効
===================================
*/

RShift::{
    if (A_ThisHotkey == A_PriorHotkey && A_TimeSincePriorHotkey <= 400) {
        clip_data := ClipboardAll()
        A_Clipboard := ""
        Send "^c"
        ClipWait 0.5
        copied := String(A_Clipboard)
        if (RegExMatch(copied, "^[0-9a-zA-Z\-\~]+$")) {
            Send "{F13}" copied
        }
        A_Clipboard := clip_data
        return
    }
}

/*
===================================
どのアプリからもググる or URLを開く
右 Shift + g で発火
※ メインブラウザが Edge, Chrome も併用しているのが前提
- Chrome 上で操作した場合 → Chrome で開く
- それ以外で操作した場合
    - Edge が開いている場合は Edge で開く
    - 開いていない場合はメッセージを表示
===================================
*/

RShift & g::{
    clip_data := ClipboardAll()
    A_Clipboard := ""
    Send "^c"
    ClipWait 1
    copied := String(A_Clipboard)
    if (copied != "") {
        if not WinActive("ahk_exe chrome.exe") {
            if WinExist("ahk_exe msedge.exe") {
                WinActivate
            } else {
                MsgBox "Edge が起動していません"
                A_Clipboard := clip_data
                return
            }
        }
        Send "^t"
        Send copied
        Send "{Enter}"
    }
    A_Clipboard := clip_data
}
