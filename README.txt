----------------------------------------------
hataori VBA

MIT License
Copyright 2023 Fukasawa Takashi
----------------------------------------------

1.Description
  This program operates a web browser from VBA (Microsoft Office 2019+).

2. How to use
  (1). Import the following files into VBE.
    

  (2). Set the following references.
    Microsoft VBScript Regular Expressions 5.5
    Microsoft Scripting Runtime

  (3) Edit the following in the .hataoriSetting module.
      Enter the path to the folder where you installed hataori.
    '// hataori install directory
    Public Function GetInstallDirectoryPath() As String
        GetInstallDirectoryPath = "C:\nokoko\hataori"
    End Function

3. Code Example
  Sub Example()
      Dim browser As New hataoriBrowser: browser.SetBrowserType = hvBrowserTypeChrome
      Dim html As hataoriElement: Set html = browser.Document.QuerySelector("html")
      Dim body As hataoriElements: Set body = html.QuerySelectorAll("body")
      Debug.Print body(0).InnerHTML
  End Sub

EOF