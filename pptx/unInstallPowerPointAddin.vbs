' �萔��`
'' �Ώۂ̃A�h�C���t�@�C��
Const FILLE_NAME = "OhinaWork.ppam"
Const ADDIN_NAME = "OhinaWork"

' ���C������
Call installAddin

' �A�h�C�����C���X�g�[��
Sub installAddin()

    Dim objPowerPoint
    Dim objExcel
    Dim objFileSys
    Dim objAddin
    Dim objAddins    
    Dim addinFolderPath
    Dim currentFolderPath
    Dim addinFilePath
    Dim addinFilePathOld
    Dim currentFilePath
    Dim count
    
   ' �Ǘ��Ҍ����擾
'    Set obj = Wscript.CreateObject("Shell.Application")
'    if Wscript.Arguments.Count = 0 then
'        obj.ShellExecute "wscript.exe", WScript.ScriptFullName & " runas", "", "runas", 1
'        Wscript.Quit
'    end if
    
    ' Office Object��`
    Set objExcel = CreateObject("Excel.Application")
    Set objPowerPoint = CreateObject("PowerPoint.Application")
    Set objFileSys = CreateObject("Scripting.FileSystemObject")
    Set objAddins = objPowerPoint.Addins
    
    ' �A�h�C���t�H���_�[�擾
    addinFolderPath = objExcel.Application.UserLibraryPath

    ' �J�����g�t�H���_�[�擾
    currentFolderPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
    imageFolderPath = currentFolderPath & "\img"

    ' �A�h�C���t�@�C���p�X���擾
    addinFilePath = objFileSys.BuildPath(addinFolderPath, FILLE_NAME)
    
    '�A�h�C�����o�^�ς݂̏ꍇ�͉��������{
    For count = 1 To objAddins.Count
      Set objAddin = objAddins.item(count)
      
      If objAddin.Name = ADDIN_NAME Then
        objAddin.AutoLoad = False
        objPowerPoint.Addins.Remove ADDIN_NAME
      End If
      
    Next
    
    ' �p���[�|�C���g���~
    objPowerPoint.Quit
    
    Set objPowerPoint = Nothing
    Set objExcel = Nothing
    Set objFileSys = Nothing
    Set objAddins = Nothing
    Set objAddin = Nothing

    MsgBox "�A�h�C����" & addinFilePath & " ����A���C���X�g�[�����܂����B"
End Sub
