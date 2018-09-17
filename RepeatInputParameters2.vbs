Dim BEF
Dim AFT
Dim ADJ
Dim CNTMAX
Dim CntSuccess
Set WshShell=Wscript.CreateObject("Wscript.Shell")

BEF = InputBox("変更前の数値")
AFT = InputBox("比較したい数値")
PRM = GetAdjValue(BEF,AFT)
'WScript.Echo CStr( PRM )
CNTMAX = InputBox("入力回数")

If IsNumeric(PRM) = true AND IsNumeric(CNTMAX) = true Then
	CntSuccess = 0
	WScript.Echo "補正値" & CStr(PRM) & "で補正を行います。"
	WScript.Sleep(5000)

	For s = 1 To CNTMAX Step 1
	    If sendkeyunit(PRM) Then
			CntSuccess = CntSuccess + 1
		End If
		WScript.Sleep(200)
	Next
	WScript.Echo CStr(CntSuccess) & "/" & CStr(CNTMAX) & "回変更に成功しました！"
Else
	WScript.Echo "失敗しました!"
End If
Set WshShell = Nothing

Function sendkeyunit( adj_param )
	Dim aft
	sendkeyunit = False

	WshShell.SendKeys("{ENTER}")
	WshShell.SendKeys("^(c)")
	WScript.Sleep(100)

	If IsNumeric(GetClipboardText) = true Then
		aft = CCur(GetClipboardText) + CCur(adj_param)
		WshShell.SendKeys(CStr(aft))
		WshShell.SendKeys("{ENTER}")
		sendkeyunit = true
	End If
End Function

Function GetClipboardText()
    Dim objHTML
    Set objHTML = CreateObject("htmlfile")
    GetClipboardText = Trim(objHTML.ParentWindow.ClipboardData.GetData("text"))
End Function

Function GetAdjValue(GetValue,NormValue)
	GetAdjValue = GetValue
	If IsNumeric(GetValue) = true AND  IsNumeric(NormValue) = true Then
		GetAdjValue =  CCur(NormValue) - CCur(GetValue)
	End If
End Function