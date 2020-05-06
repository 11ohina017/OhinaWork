' 定数定義
'' 対象のアドインファイル
Const FILLE_NAME = "OhinaWork.ppam"
Const ADDIN_NAME = "OhinaWork"

' メイン処理
Call installAddin

' アドインをインストール
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
    
   ' 管理者権限取得
'    Set obj = Wscript.CreateObject("Shell.Application")
'    if Wscript.Arguments.Count = 0 then
'        obj.ShellExecute "wscript.exe", WScript.ScriptFullName & " runas", "", "runas", 1
'        Wscript.Quit
'    end if
    
    ' Office Object定義
    Set objExcel = CreateObject("Excel.Application")
    Set objPowerPoint = CreateObject("PowerPoint.Application")
    Set objFileSys = CreateObject("Scripting.FileSystemObject")
    Set objAddins = objPowerPoint.Addins
    
    ' アドインフォルダー取得
    addinFolderPath = objExcel.Application.UserLibraryPath

    ' カレントフォルダー取得
    currentFolderPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
    imageFolderPath = currentFolderPath & "\img"

    ' アドインファイルパスを取得
    addinFilePath = objFileSys.BuildPath(addinFolderPath, FILLE_NAME)
    
    'アドインが登録済みの場合は解除を実施
    For count = 1 To objAddins.Count
      Set objAddin = objAddins.item(count)
      
      If objAddin.Name = ADDIN_NAME Then
        objAddin.AutoLoad = False
        objPowerPoint.Addins.Remove ADDIN_NAME
      End If
      
    Next
    
    ' パワーポイントを停止
    objPowerPoint.Quit
    
    Set objPowerPoint = Nothing
    Set objExcel = Nothing
    Set objFileSys = Nothing
    Set objAddins = Nothing
    Set objAddin = Nothing

    MsgBox "アドインを" & addinFilePath & " からアンインストールしました。"
End Sub
