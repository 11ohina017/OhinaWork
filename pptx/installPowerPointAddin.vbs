' �萔��`
'' �Ώۂ̃A�h�C���t�@�C��
Const FILLE_NAME="OhinaWork.ppam"

' ���C������
Call installAddin

' �A�h�C�����C���X�g�[��
Sub installAddin()
    Dim objPowerPoint
    Dim objExcel
    Dim addinFolderPath
    Dim currentFolderPath
    Dim addinFilePath
    Dim currentFilePath
    Dim objFileSys
    Dim objAddin

    ' Office Object��`
    Set objExcel = CreateObject("Excel.Application")
    Set objPowerPoint = CreateObject("PowerPoint.Application")
    Set objFileSys = CreateObject("Scripting.FileSystemObject")

    ' �A�h�C���t�H���_�[�擾
    addinFolderPath = objExcel.Application.UserLibraryPath

    ' �J�����g�t�H���_�[�擾
    currentFolderPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")

    ' �A�h�C���t�@�C���p�X���擾
    addinFilePath   = objFileSys.BuildPath(addinFolderPath, FILLE_NAME)

    ' �J�����g�t�@�C���p�X���擾
    currentFilePath   = objFileSys.BuildPath(currentFolderPath, FILLE_NAME)

    ' �Ώۂ̃A�h�C���t�@�C�����A�h�C���t�H���_�Ɉړ�
    objFileSys.CopyFile currentFilePath, addinFilePath

    ' �p���[�|�C���g���\���ŋN��
    objPowerPoint.Presentations.Add(False)

    ' �A�h�C����ǉ�
    Set objAddin = objPowerPoint.AddIns.Add(addinFilePath)

    ' �N�����Ɏ����I�ɓǂ�
    objAddin.AutoLoad = True

    ' �p���[�|�C���g���~
    objPowerPoint.Quit
    Set objPowerPoint = Nothing
    Set objExcel = Nothing
    Set objFileSys = Nothing

    MsgBox "�A�h�C����" & addinFilePath & " �ɃC���X�g�[�����܂����B"
End Sub
