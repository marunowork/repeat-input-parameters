Dim INPUT_PARAM
Dim MAXCOUNT
Set WshShell=Wscript.CreateObject("Wscript.Shell")

'メインルーチン
INPUT_PARAM = InputBox("入力する値は")
MAXCOUNT = InputBox("入力を繰り返す回数")

If IsNumeric(INPUT_PARAM) = true AND IsNumeric(MAXCOUNT) = true Then
	WScript.Echo "準備はよろしいですか?"
	'5秒待機
	WScript.Sleep(5000)

	'キー送信を繰り返す
	For s = 1 To MAXCOUNT Step 1
		'キー送信を実行する
	    call sendkeyunit(INPUT_PARAM)
		WScript.Sleep(200)
	Next

	WScript.Echo "完了!"
Else
	WScript.Echo "失敗!"
End If
Set WshShell = Nothing

'キー送信を実行する。
Sub sendkeyunit( InputParam )

	'編集前のグリッド入力値をクリップボードに格納
	WshShell.SendKeys("{ENTER}")
	WshShell.SendKeys("^(c)")
	WScript.Sleep(100)

	WshShell.SendKeys("{ENTER}")
	WshShell.SendKeys("{ENTER}")
	WshShell.SendKeys(CStr(InputParam))

	'クリップボードの値を判定する
	If IsNumeric(GetClipboardText) = true AND GetClipboardText < 0 Then
		'負数の場合
		WshShell.SendKeys("{DOWN}")
	Else
		'正数の場合
		WshShell.SendKeys("{ENTER}")
	End If
End Sub

'クリップボードの値を取得する。
Function GetClipboardText()
    Dim objHTML
    Set objHTML = CreateObject("htmlfile")
    GetClipboardText = Trim(objHTML.ParentWindow.ClipboardData.GetData("text"))
End Function