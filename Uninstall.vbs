' [VBScript �� Excel �ɃA�h�C���������ŃC���X�g�[��/�A���C���X�g�[��������@: ���� SE �̂Ԃ₫](http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html)

On Error Resume Next

Dim installPath 
Dim addInName 
Dim addInFileName 
Dim objExcel 
Dim objAddin

'�A�h�C������ݒ� 
addInName = "�^��Excel2016�ǉ��֐�" 
addInFileName = "Excel2016Func.xlam"

IF MsgBox("�u" & addInName & "�v �A�h�C�����A���C���X�g�[�����܂����H", vbYesNo + vbQuestion) = vbNo Then 
  WScript.Quit 
End IF

'Excel �C���X�^���X�� 
Set objExcel = CreateObject("Excel.Application") 
objExcel.Workbooks.Add

'�A�h�C���o�^���� 
For i = 1 To objExcel.Addins.Count 
  Set objAddin = objExcel.Addins.item(i) 
  If objAddin.Name = addInFileName Then 
    objAddin.Installed = False 
  End If 
Next

'Excel �I�� 
objExcel.Quit

Set objAddin = Nothing 
Set objExcel = Nothing

Set objWshShell = CreateObject("WScript.Shell") 
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'�C���X�g�[����p�X�̍쐬 
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName] 
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'�t�@�C���폜 
If objFileSys.FileExists(installPath) = True Then 
  objFileSys.DeleteFile installPath , True 
Else 
  MsgBox "�A�h�C���t�@�C�������݂��܂���B", vbExclamation 
End If

Set objWshShell = Nothing 
Set objFileSys = Nothing

IF Err.Number = 0 THEN 
   MsgBox "�A�h�C���̃A���C���X�g�[�����I�����܂����B", vbInformation 
ELSE 
   MsgBox "�G���[���������܂����B" & vbCrLF & "���s�����m�F���Ă��������B", vbExclamation 
End IF