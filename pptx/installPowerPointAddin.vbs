' �萔��`
'' �Ώۂ̃A�h�C���t�@�C��
Const FILLE_NAME = "OhinaWork.ppam"
Const ADDIN_NAME = "OhinaWork"
Const AZURE_ICON = "azure-icons"
Const AWS_ICON = "aws-icons"

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
    imageAddinFolderPath = addinFolderPath & "\img"    
    azureAddinFolderPath = imageAddinFolderPath & "\" &AZURE_ICON
    awsAddinFolderPath = imageAddinFolderPath & "\" & AWS_ICON
        
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
    
    ' �J�����g�t�@�C���p�X���擾
    currentFilePath = objFileSys.BuildPath(currentFolderPath, FILLE_NAME)
    
    ' �Ώۂ̃A�h�C���t�@�C�����A�h�C���t�H���_�Ɉړ�
    objFileSys.CopyFile currentFilePath, addinFilePath ,True
    
    ' �t�H���_�폜
    If objFileSys.FolderExists(azureAddinFolderPath) Then
        objFileSys.DeleteFolder azureAddinFolderPath, True
    End If
    If objFileSys.FolderExists(awsAddinFolderPath) Then
        objFileSys.DeleteFolder awsAddinFolderPath, True
    End If

 
    ' �摜�t�H���_���A�h�C���t�H���_�ɃR�s�[
    objFileSys.CopyFolder imageFolderPath, addinFolderPath, True
 
     ' Addin�I�u�W�F�N�g��ݒ�
     Set objAddin = objPowerPoint.AddIns.Add(addinFilePath)
       
    ' �p���[�|�C���g���\���ŋN��
    objPowerPoint.Presentations.Add(False)

    ' �N�����Ɏ����I�ɓǂ�
    objAddin.AutoLoad = True
    
    ' �p���[�|�C���g���~
    objPowerPoint.Quit
    
    Set objPowerPoint = Nothing
    Set objExcel = Nothing
    Set objFileSys = Nothing
    Set objAddins = Nothing
    Set objAddin = Nothing

    MsgBox "�A�h�C����" & addinFilePath & " �ɃC���X�g�[�����܂����B"
End Sub
