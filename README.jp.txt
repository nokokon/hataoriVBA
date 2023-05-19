----------------------------------------------
hataori VBA

MIT License
Copyright 2023 Fukasawa Takashi
----------------------------------------------

1. 説明
  VBA (Microsoft Office 2019+) からWebブラウザーを操作します。

2. 使用方法
  (1). 以下のファイルをVBEにインポートします。
    

  (2). 以下を参照設定します。
    Microsoft VBScript Regular Expressions 5.5
    Microsoft Scripting Runtime

  (3). hataoriSettingモジュールの以下を編集します。
      hataoriをインストールしたフォルダーのパスを入力してください。
    '// hataori install directory
    Public Function GetInstallDirectoryPath() As String
        GetInstallDirectoryPath = "C:\nokoko\hataori"
    End Function

3. コード例
  Sub Example()
      Dim browser As New hataoriBrowser: browser.SetBrowserType = hvBrowserTypeChrome
      Dim html As hataoriElement: Set html = browser.Document.QuerySelector("html")
      Dim body As hataoriElements: Set body = html.QuerySelectorAll("body")
      Debug.Print body(0).InnerHTML
  End Sub

EOF