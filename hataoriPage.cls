VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "hataoriPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************
' hataoriPage Ver 2023-05-18-01
'
' License: MIT License (http://www.opensource.org/licenses/mit-license.php)
'  (c) 2023 Fukasawa Takashi
'
' Note: Please enable the following in the reference settings.
'       * Microsoft Scripting Runtime
'***************************************************************

Option Explicit

Private pPath
Private pName
Private pId
Private pTagName
Private pBrowserType As hvBrowserType
Private pBrowserPath As String
Private pBrowserCaption As String
Private pReqFile As String
Private pResFile As String

Public Property Let SetBrowserType(ByVal browser_type As hvBrowserType)
    pPath = "/"
    pName = Empty
    pId = Empty
    pTagName = Empty
    pBrowserType = browser_type
    pBrowserPath = hataoriSetting.GethataoriBrowserSetting("browser_path", browser_type)
    pBrowserCaption = hataoriSetting.GethataoriBrowserSetting("browser_caption", browser_type)
    pReqFile = hataoriSetting.GethataoriBrowserSetting("req_file", browser_type)
    pResFile = hataoriSetting.GethataoriBrowserSetting("res_file", browser_type)
End Property

Public Function SetAxonPath(ByVal axon_path As String, Optional browser_type As hvBrowserType = -1) As Boolean
    SetAxonPath = False

    If browser_type > -1 Then pBrowserType = browser_type

    Dim path As Dictionary: Set path = hataori.ConnectBrowser(hvPathToElement, pReqFile, pResFile, axon_path)
    If path Is Nothing Then Exit Function

    pBrowserPath = hataoriSetting.GethataoriBrowserSetting("browser_path", pBrowserType)
    pBrowserCaption = hataoriSetting.GethataoriBrowserSetting("browser_caption", pBrowserType)
    pReqFile = hataoriSetting.GethataoriBrowserSetting("req_file", pBrowserType)
    pResFile = hataoriSetting.GethataoriBrowserSetting("res_file", pBrowserType)
    pPath = path("path")
    pName = path("name")
    pId = path("id")
    pTagName = path("tag_name")

    SetAxonPath = True
End Function

Public Function GetAxonPath() As String
    GetAxonPath = pPath
End Function

Public Function Document() As hataoriElement
    Set Document = New hataoriElement
    Dim pathDict As New Dictionary
    pathDict.Add "path", pPath
    pathDict.Add "name", Empty
    pathDict.Add "id", Empty
    pathDict.Add "tag_name", Empty
    Document.init pathDict, pBrowserType
End Function

Public Function Forms() As hataoriElements
    Dim paths As Collection: Set paths = hataori.ConnectBrowser(hvForms, pReqFile, pResFile, pPath)
    Set Forms = New hataoriElements: Forms.init paths, pBrowserType
End Function

Public Function Frames() As hataoriElements
    Dim paths As Collection: Set paths = hataori.ConnectBrowser(hvCssSelector, pReqFile, pResFile, pPath, "frame")
    Set Frames = New hataoriElements: Frames.init paths, pBrowserType
End Function

Public Function Head() As hataoriElement
    Dim path As Dictionary: Set path = hataori.ConnectBrowser(hvHead, pReqFile, pResFile, pPath)
    If Not path Is Nothing Then Set Head = New hataoriElement: Head.init path, pBrowserType
End Function

Public Function body() As hataoriElement
    Dim path As Dictionary: Set path = hataori.ConnectBrowser(hvBody, pReqFile, pResFile, pPath)
    If Not path Is Nothing Then Set body = New hataoriElement: body.init path, pBrowserType
End Function

Public Function ActiveElement() As hataoriElement
    Dim path As Dictionary: Set path = hataori.ConnectBrowser(hvCssSelector, pReqFile, pResFile, pPath)
    If Not path Is Nothing Then Set ActiveElement = New hataoriElement: ParentElement.init path, pBrowserType
End Function

Public Function Title() As String
    Title = hataori.ConnectBrowser(hvGetTitle, pReqFile, pResFile, pPath)
End Function

Public Function Url() As String
    Url = hataori.ConnectBrowser(hvGetUrl, pReqFile, pResFile, pPath)
End Function

Public Sub ScrollTo(ByVal horizontal_coordinate As Long, ByVal vertical_coordinate As Long)
    If Not hataori.ConnectBrowser(hvScrollTo, pReqFile, pResFile, pPath, horizontal_coordinate, vertical_coordinate) Then _
            Err.Raise vbObjectError + 12, "hataoriPage", hataoriSetting.GetErrorString(12)
End Sub

Public Sub ScrollBy(ByVal horizontal_coordinate As Long, ByVal vertical_coordinate As Long)
    If Not hataori.ConnectBrowser(hvScrollBy, pReqFile, pResFile, pPath, horizontal_coordinate, vertical_coordinate) Then _
            Err.Raise vbObjectError + 12, "hataoriPage", hataoriSetting.GetErrorString(12)
End Sub

Public Sub PrintPage()
    If Not hataori.ConnectBrowser(hvPrintout, pReqFile, pResFile, pPath) Then _
         Err.Raise vbObjectError + 12, "hataoriPage", hataoriSetting.GetErrorString(12)
End Sub

Public Sub GoBack()
    If Not hataori.ConnectBrowser(hvHistoryBack, pReqFile, pResFile, pPath) Then _
         Err.Raise vbObjectError + 12, "hataoriPage", hataoriSetting.GetErrorString(12)
End Sub

Public Sub GoForward()
    If Not hataori.ConnectBrowser(hvHistoryForward, pReqFile, pResFile, pPath) Then _
         Err.Raise vbObjectError + 12, "hataoriPage", hataoriSetting.GetErrorString(12)
End Sub

Public Sub init(element_path As Dictionary, ByVal browser_type As hvBrowserType)
    pPath = element_path("path")
    pName = element_path("name")
    pId = element_path("id")
    pTagName = element_path("tag_name")
    pBrowserType = browser_type
    pBrowserPath = hataoriSetting.GethataoriBrowserSetting("browser_path", browser_type)
    pBrowserCaption = hataoriSetting.GethataoriBrowserSetting("browser_caption", browser_type)
    pReqFile = hataoriSetting.GethataoriBrowserSetting("req_file", browser_type)
    pResFile = hataoriSetting.GethataoriBrowserSetting("res_file", browser_type)
End Sub

Private Sub Class_Initialize()
    pPath = Empty
    pName = Empty
    pId = Empty
    pTagName = Empty
    pBrowserType = 99999
    pBrowserPath = vbNullString
    pBrowserCaption = vbNullString
    pReqFile = vbNullString
    pResFile = vbNullString
End Sub


