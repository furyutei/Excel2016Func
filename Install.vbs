' [VBScript で Excel にアドインを自動でインストール/アンインストールする方法: ある SE のつぶやき](http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html)

On Error Resume Next

Dim installPath 
Dim addInName 
Dim addInFileName 
Dim objExcel 
Dim objAddin

'アドイン情報を設定 
addInName = "疑似Excel2016追加関数" 
addInFileName = "Excel2016Func.xlam"

IF MsgBox("「" & addInName & "」アドインをインストールしますか？", vbYesNo + vbQuestion) = vbNo Then 
  WScript.Quit 
End IF

Set objWshShell = CreateObject("WScript.Shell") 
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'インストール先パスの作成 
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName] 
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'ファイルコピー(上書き) 
objFileSys.CopyFile  addInFileName ,installPath , True

Set objWshShell = Nothing 
Set objFileSys = Nothing

'Excel インスタンス化 
Set objExcel = CreateObject("Excel.Application") 
objExcel.Workbooks.Add

'アドイン登録 
Set objAddin = objExcel.AddIns.Add(installPath, True) 
objAddin.Installed = True

'Excel 終了 
objExcel.Quit

Set objAddin = Nothing 
Set objExcel = Nothing

IF Err.Number = 0 THEN 
   MsgBox "アドインのインストールが終了しました。", vbInformation 
ELSE 
   MsgBox "エラーが発生しました。" & vbCrLF & "実行環境を確認してください。", vbExclamation 
End IF
