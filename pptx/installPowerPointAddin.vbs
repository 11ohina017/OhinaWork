' 定数定義
'' 対象のアドインファイル
Const FILLE_NAME = "OhinaWork.ppam"
Const ADDIN_NAME = "OhinaWork"
Const AZURE_ICON = "azure-icons"
Const AWS_ICON = "aws-icons"

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
    imageAddinFolderPath = addinFolderPath & "\img"    
    azureAddinFolderPath = imageAddinFolderPath & "\" &AZURE_ICON
    awsAddinFolderPath = imageAddinFolderPath & "\" & AWS_ICON
        
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
    
    ' カレントファイルパスを取得
    currentFilePath = objFileSys.BuildPath(currentFolderPath, FILLE_NAME)
    
    ' 対象のアドインファイルをアドインフォルダに移動
    objFileSys.CopyFile currentFilePath, addinFilePath ,True
    
    ' フォルダ削除
    If objFileSys.FolderExists(azureAddinFolderPath) Then
        objFileSys.DeleteFolder azureAddinFolderPath, True
    End If
    If objFileSys.FolderExists(awsAddinFolderPath) Then
        objFileSys.DeleteFolder awsAddinFolderPath, True
    End If

 
    ' 画像フォルダをアドインフォルダにコピー
    objFileSys.CopyFolder imageFolderPath, addinFolderPath, True
 
     ' Addinオブジェクトを設定
     Set objAddin = objPowerPoint.AddIns.Add(addinFilePath)
       
    ' パワーポイントを非表示で起動
    objPowerPoint.Presentations.Add(False)

    ' 起動時に自動的に読む
    objAddin.AutoLoad = True
    
    ' パワーポイントを停止
    objPowerPoint.Quit
    
    Set objPowerPoint = Nothing
    Set objExcel = Nothing
    Set objFileSys = Nothing
    Set objAddins = Nothing
    Set objAddin = Nothing

    MsgBox "アドインを" & addinFilePath & " にインストールしました。"
End Sub
