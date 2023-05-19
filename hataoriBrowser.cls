VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "hataoriBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************
' hataoriBrowser Ver 2023-05-18-01
'
' License: MIT License (http://www.opensource.org/licenses/mit-license.php)
'  (c) 2023 Fukasawa Takashi
'
' Note: Please enable the following in the reference settings.
'       * Microsoft Scripting Runtime
'***************************************************************

Option Explicit

Private pPath As String
Private pBrowserType As hvBrowserType
Private pBrowserPath As String
Private pBrowserCaption As String
Private pReqFile As String
Private pResFile As String

'// Set browser type
Public Property Let SetBrowserType(ByVal browser_type As hvBrowserType)
    pBrowserType = browser_type
    pBrowserPath = hataoriSetting.GethataoriBrowserSetting("browser_path", browser_type)
    pBrowserCaption = hataoriSetting.GethataoriBrowserSetting("browser_caption", browser_type)
    pReqFile = hataoriSetting.GethataoriBrowserSetting("req_file", browser_type)
    pResFile = hataoriSetting.GethataoriBrowserSetting("res_file", browser_type)
End Property

'// Browser
Public Function OpenBrowser( _
    ByVal url_string As String, _
    Optional ByVal timeout_seconds As Long = 10 _
) As Boolean
    OpenBrowser = False
    If Len(url_string) = 0 Then Exit Function
    OpenBrowser = hataori.OpenBrowserApplication(pReqFile, pResFile, pBrowserPath, url_string, timeout_seconds)
End Function

Public Function CloseBrowser() As Boolean
    CloseBrowser = hataori.ConnectBrowser(hvBrowserExit, pReqFile, pResFile)
End Function

Public Function BrowserWindowMode(ByVal window_mode As hvWindowMode) As Boolean
    Dim windowMode As hvMethodName
    If window_mode = 0 Then windowMode = hvNormalWindow
    If window_mode = 1 Then windowMode = hvMaximizedWindow
    If window_mode = 2 Then windowMode = hvFullscreenWindow
    BrowserWindowMode = hataori.ConnectBrowser(windowMode, pReqFile, pResFile)
End Function

'// Wait tab, host
Public Function WaitHostExist(Optional ByVal timeout_seconds As Long = 10) As Boolean
    WaitHostExist = hataori.WaitExistHost(pReqFile, pResFile, timeout_seconds)
End Function

Public Function WaitTab( _
    Optional ByVal wait_interaction As Boolean = True, _
    Optional ByVal timeout_seconds As Long = 10 _
) As Boolean
    WaitTab = hataori.WaitCompleteTab(pReqFile, pResFile, wait_interaction, timeout_seconds)
End Function

'// Control tab
Public Sub NewTab(ByVal url_string As String)
    If Not hataori.ConnectBrowser(hvNewTab, pReqFile, pResFile, url_string) Then _
         Err.Raise vbObjectError + 12, "hataoriBrowser", hataoriSetting.GetErrorString(12)
End Sub

Public Sub CloseTab()
    If Not hataori.ConnectBrowser(hvCloseTab, pReqFile, pResFile) Then _
         Err.Raise vbObjectError + 12, "hataoriBrowser", hataoriSetting.GetErrorString(12)
End Sub

Public Sub Refresh()
    If Not hataori.ConnectBrowser(hvReloadTab, pReqFile, pResFile) Then _
         Err.Raise vbObjectError + 12, "hataoriBrowser", hataoriSetting.GetErrorString(12)
End Sub

Public Sub ActiveTab()
    ActiveTab = hataori.SelectForegroundTab(pReqFile, pResFile, pBrowserCaption)
End Sub

'// Get tab information
Public Function GetTabInformation() As Dictionary
    Set GetTabInformation = hataori.ConnectBrowser(hvThisGetTab, pReqFile, pResFile, "this")
End Function

Public Function GetTabsInformation() As Collection
    Set GetTabsInformation = hataori.ConnectBrowser(hvGetTabs, pReqFile, pResFile)
End Function

Public Function GetTabId() As Double
    GetTabId = -1
    Dim tabInfo As Dictionary: Set tabInfo = hataori.ConnectBrowser(hvThisGetTab, pReqFile, pResFile, "this")
    If tabInfo.Exists("id") Then GetTabId = tabInfo("id")
End Function

Public Property Let SetUrl(ByVal url_string As String)
    Dim resp As Boolean: resp = hataori.ConnectBrowser(hvJumpTab, pReqFile, pResFile, url_string)
    If TypeName(resp) = "Boolean" Then If resp = False Then Err.Raise vbObjectError + 12, "hataoriBrowser", hataoriSetting.GetErrorString(12)
End Property

'// Selected Tab
Public Function SelectTabByTitle( _
    ByVal title_string As String, _
    Optional ByVal timeout_seconds As Long = 10 _
) As Boolean
    SelectTabByTitle = False
    If Len(title_string) = 0 Then Exit Function
    SelectTabByTitle = selectTab("title", title_string, timeout_seconds)
End Function

Public Function SelectTabByURL( _
    ByVal url_string As String, _
    Optional ByVal timeout_seconds As Long = 10 _
) As Boolean
    SelectTabByURL = False
    If (Len(url_string) = 0) Or ((Left(url_string, 7) <> "http://") And (Left(url_string, 8) <> "https://")) Then Exit Function
    SelectTabByURL = selectTab("url", url_string, timeout_seconds)
End Function

Public Function SelectTabByID( _
    ByVal id_number As Double, _
    Optional ByVal timeout_seconds As Long = 10 _
) As Boolean
    SelectTabByID = False
    If id_number < 0 Then Exit Function
    SelectTabByID = selectTab("id", id_number, timeout_seconds)
End Function

'// Send key, text
Public Function SendKey(ByVal key_string As String) As Boolean
    SendKey = hataori.SendKeyStringToBrowser(pReqFile, pResFile, key_string, pBrowserCaption)
End Function

Public Function SendText(ByVal unicode_string As String) As Boolean
    SendText = hataori.SendUnicodeStringToBrowser(pReqFile, pResFile, unicode_string, pBrowserCaption, True)
End Function

'// Control dialog
Public Function DialogOK() As Boolean
    DialogOK = hataori.SendKeyStringToBrowser(pReqFile, pResFile, "enter", pBrowserCaption)
End Function

Public Function DialogCancel() As Boolean
    DialogCancel = hataori.SendKeyStringToBrowser(pReqFile, pResFile, "esc", pBrowserCaption)
End Function

Public Function DialogInput(ByVal unicode_string As String) As Boolean
    DialogCancel = hataori.SendUnicodeStringToBrowser(pReqFile, pResFile, unicode_string, pBrowserCaption)
End Function

'// Get hataoriElement
Public Function Document() As hataoriElement
    Set Document = New hataoriElement
    Dim pathDict As New Dictionary
    pathDict.Add "path", pPath
    pathDict.Add "name", Empty
    pathDict.Add "id", Empty
    pathDict.Add "tag_name", Empty
    Document.init pathDict, pBrowserType
End Function

'// Get hataoriPage
Public Function Page() As hataoriPage
    Set Page = New hataoriPage
    Dim pathDict As New Dictionary
    pathDict.Add "path", pPath
    pathDict.Add "name", Empty
    pathDict.Add "id", Empty
    pathDict.Add "tag_name", Empty
    Page.init pathDict, pBrowserType
End Function

'// Private function
' USage: select_type: "id", "title", "url"
Private Function selectTab( _
    ByVal select_type As String, _
    ByVal option_value, _
    Optional ByVal timeout_seconds As Long = 10 _
) As Boolean
    selectTab = False
    Dim startTime As Double: startTime = Timer
    Dim ret As Boolean
    Do
        Select Case select_type
            Case "id"
                ret = hataori.HostMessaging(pReqFile, pResFile, "browser", "select_tab", "id", CDbl(option_value))
            Case "title"
                ret = hataori.HostMessaging(pReqFile, pResFile, "browser", "select_tab", "title", CStr(option_value))
            Case "url"
                ret = hataori.HostMessaging(pReqFile, pResFile, "browser", "select_tab", "url", CStr(option_value))
            Case Else
                Exit Function
        End Select
        If TypeName(ret) <> "Boolean" Then: Exit Function
        If ret Then
            selectTab = hataori.SelectForegroundTab(pReqFile, pResFile, pBrowserCaption)
            Exit Do
        End If
        SecondsSleep 0.1
        DoEvents
    Loop While Timer - startTime <= timeout_seconds
End Function

Private Sub Class_Initialize()
    pPath = "/"
    pBrowserType = 99999
    pBrowserPath = vbNullString
    pBrowserCaption = vbNullString
    pReqFile = vbNullString
    pResFile = vbNullString
End Sub
