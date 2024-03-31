/*
===================================
共通関数
===================================
*/

/*
アプリケーションのウインドウの状態を変更する関数
- アプリケーションが起動している場合、以下を toggle で
    - アプリケーションのウインドウをアクティブにする
    - アプリケーションのウインドウを非アクティブにし背後に隠す
- アプリケーションが起動していない場合は cmd で起動しアクティブに
    - cmd が指定されない場合はメッセージを表示

title: Window Titles https://ahkscript.github.io/ja/docs/v2/misc/WinTitle.htm
cmd: アプリを起動するためのコマンド
*/

SwitchAppStatus(title, cmd:="") {
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
貼り付け完了を待つ
貼り付けた後に全選択・コピーして貼り付けたテキストが含まれるか確認する
含まれていたら OK 
前後のクリップボードの内容を保全する
- copied: 貼り付けるテキスト
- max_loop: ループ回数の上限
*/
PasteAndWait(text, max_loop := 500) {
    SendInput "^v"
    tmp_clip := ClipboardAll()
    i := 0
    Loop {
        SendInput "^a" "^c"
        sleep 50
        i += 1
        if (InStr(String(A_Clipboard), text) || i >= max_loop) {
            break
        }
    }
    A_Clipboard := tmp_clip
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
 Edge, Chrome がアクティブの場合は F12 を送信しない
 (Dev Tools が開くため)
 $ は Esc の無限ループ防止のため
*/
$Esc::
{
    if (WinActive("ahk_exe chrome.exe") or WinActive("ahk_exe msedge.exe")) {
        Send '{Esc}'
    } else {
        Send '{Esc}{F12}'
    }
}
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
	SwitchAppStatus("ahk_exe WindowsTerminal.exe", "wt -p Ubuntu")
}


/*
Ctrl + \ エクスプローラー
*/
^\::
{
	SwitchAppStatus("ahk_class CabinetWClass", "explorer")
}


/*
Ctrl + 0 VSCode
*/
^0::
{
	SwitchAppStatus("ahk_exe Code.exe", "Code")
}


/*
===================================
GUI
アプリを listview で選択するランチャ
数字でも選べる
===================================
*/

ExitGui(*) {
    If WinExist("Launcher") {
		WinClose
    }
}

/*
右 Shift + L でランチャ起動
*/
Rshift & l::{
    If WinExist("Launcher") {
        WinActivate
        return
    }
    apps := [
        Map("app_name", "Outlook", "win_title", "ahk_exe outlook.exe"),
        Map("app_name", "Teams", "win_title", "ahk_exe ms-teams.exe"),
        Map("app_name", "Edge", "win_title", "ahk_exe msedge.exe"),
        Map("app_name", "Chrome", "win_title", "ahk_exe chrome.exe"),
        Map("app_name", "VSCode", "win_title", "ahk_exe Code.exe"),
    ]
    MyGui := Gui(, "Launcher")
    LV := MyGui.Add("ListView", "-Multi R" apps.Length +1, ["Id", "Name"])
    for (k,v in apps) {
        opt_disabled := ""
        name_disabled := ""
        if not WinExist(v["win_title"]) {
            opt_disabled := "disabled "
            name_disabled := " (Not Running)"
        }
        LV.Add(, k, v["app_name"] name_disabled)
    }
    LV.Add(, "q", "Quit")
    LV.ModifyCol ; Auto-size each column to fit its contents.

    ; Enter 打鍵を検知するために隠しボタンを配置
    MyGui.Add("Button", "Hidden Default", "OK").OnEvent("Click", LV_Enter)
    LV_Enter(*) {
        if MyGui.FocusedCtrl != LV
            return
        id := LV.GetNext(0, "Forcused")
        if apps.Has(id) {
            SwitchAppStatus(apps[id]["win_title"])
            WinClose("Launcher")
        } else if (id = apps.Length + 1) {
            WinClose("Launcher")
        }
    }

    MyGui.OnEvent("Escape", ExitGui) ; ウインドウを閉じる
    SetTimer(ExitGui, 10000)
    MyGui.Show
}

/*
===================================
ローマ字を再変換
右シフト 2連打で発火 Interval <= 400ms
選択中のテキストが `半角英数-~のみ1文字以上(末尾の空白だけ許容)` の場合だけ有効
    末尾の空白は単語を打鍵→変換の流れで打ってしまうから
===================================
*/

RShift::{
    if (A_ThisHotkey == A_PriorHotkey && A_TimeSincePriorHotkey <= 400) {
        clip_data := ClipboardAll()
        A_Clipboard := ""
        Send "^c"
        ClipWait 0.5
        copied := String(A_Clipboard)
        if (RegExMatch(copied, "^[0-9a-zA-Z\-\~]+\s*$")) {
            copied := StrReplace(copied, " ")
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

検索文字列の入力には Send ではなく SendInput を使う
> また、送信中にキーボードやマウスを操作した場合、バッファリングが行われるため、
> ユーザーのキー入力と送信中のキー入力が混在することを防ぐことができます。
https://ahkscript.github.io/ja/docs/v2/lib/Send.htm
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
        SendInput "^t" 
        PasteAndWait(copied)
        SendInput "{Enter}"
    }
    A_Clipboard := clip_data
}

/*
===================================
Edge でメールを開く
URL の入力には Send ではなく SendInput を使う
===================================
*/

RShift & m::{
    if WinExist("ahk_exe msedge.exe") {
        WinActivate
        clip_data := ClipboardAll()
        A_Clipboard := ""
        A_Clipboard := "https://outlook.office.com/mail/"
        ClipWait 1
        SendInput "^t" 
        PasteAndWait(A_Clipboard)
        SendInput "{Enter}"
        sleep 500
        A_Clipboard := clip_data
    } else {
        MsgBox "Edge が起動していません"
        return
    }
}

/*
===================================
Edge カレンダーを開く
URL の入力には Send ではなく SendInput を使う
===================================
*/

RShift & c::{
    if WinExist("ahk_exe msedge.exe") {
        WinActivate
        clip_data := ClipboardAll()
        A_Clipboard := ""
        A_Clipboard := "https://outlook.office.com/calendar/view/week"
        ClipWait 1
        SendInput "^t" 
        PasteAndWait(A_Clipboard)
        SendInput "{Enter}"
        sleep 500
        A_Clipboard := clip_data
    } else {
        MsgBox "Edge が起動していません"
        return
    }
}

/*
===================================
どのアプリからも Edge の新タブを開く
右 Shift + t で発火
※ メインブラウザが Edge, Chrome も併用しているのが前提
- Edge が開いている場合は Edge で開く
- 開いていない場合はメッセージを表示
===================================
*/

RShift & t::{
    if not WinActive("ahk_exe chrome.exe") {
        if WinExist("ahk_exe msedge.exe") {
            WinActivate
            Send "^t"
        } else {
            MsgBox "Edge が起動していません"
        }
    }
}