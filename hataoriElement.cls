VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "hataoriElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************
' hataoriElement Ver 2023-05-18-01
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

Public Function SetAxonPath(ByVal axon_path As String) As Boolean
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

'// Elements
Public Function QuerySelectorAll(ByVal css_string As String) As hataoriElements
    Dim paths As Collection: Set paths = hataori.ConnectBrowser(hvCssSelector, pReqFile, pResFile, pPath, css_string)
    Set QuerySelectorAll = New hataoriElements: QuerySelectorAll.init paths, pBrowserType
End Function

Public Function GetElementsByName(ByVal name_string As String) As hataoriElements
    Dim paths As Collection: Set paths = hataori.ConnectBrowser(hvCssSelector, pReqFile, pResFile, pPath, "[name=""" & name_string & """]")
    Set GetElementsByName = New hataoriElements: GetElementsByName.init paths, pBrowserType
End Function

Public Function GetElementsByClassName(ByVal class_name_string As String) As hataoriElements
    Dim paths As Collection: Set paths = hataori.ConnectBrowser(hvCssSelector, pReqFile, pResFile, pPath, "." & class_name_string)
    Set GetElementsByClassName = New hataoriElements: GetElementsByClassName.init paths, pBrowserType
End Function

Public Function GetElementsByTagName(ByVal tag_name_string As String) As hataoriElements
    Dim paths As Collection: Set paths = hataori.ConnectBrowser(hvCssSelector, pReqFile, pResFile, pPath, tag_name_string)
    Set GetElementsByTagName = New hataoriElements: GetElementsByTagName.init paths, pBrowserType
End Function

Public Function Children() As hataoriElements
    Dim paths As Collection: Set paths = hataori.ConnectBrowser(hvChildren, pReqFile, pResFile, pPath)
    Set Children = New hataoriElements: Children.init paths, pBrowserType
End Function

Public Function Brothers() As hataoriElements
    Dim paths As Collection: Set paths = hataori.ConnectBrowser(hvBros, pReqFile, pResFile, pPath)
    Set Brothers = New hataoriElements: Brothers.init paths, pBrowserType
End Function

'// Element
Public Function QuerySelector(ByVal css_string As String) As hataoriElement
    Set QuerySelector = Nothing
    Dim paths As Collection: Set paths = hataori.ConnectBrowser(hvCssSelector, pReqFile, pResFile, pPath, css_string)
    If paths.Count > 0 Then Set QuerySelector = New hataoriElement: QuerySelector.init paths(1), pBrowserType
End Function

Public Function GetElementById(ByVal id_string As String) As hataoriElement
    Set GetElementById = Nothing
    Dim paths As Collection: Set paths = hataori.ConnectBrowser(hvCssSelector, pReqFile, pResFile, pPath, "#" & id_string)
    If paths.Count > 0 Then Set GetElementById = New hataoriElement: GetElementById.init paths(1), pBrowserType
End Function

Public Function FirstChild() As hataoriElement
    Dim path As Dictionary: Set path = hataori.ConnectBrowser(hvFirstChild, pReqFile, pResFile, pPath)
    If Not path Is Nothing Then Set FirstChild = New hataoriElement: FirstChild.init path, pBrowserType
End Function

Public Function LastChild() As hataoriElement
    Dim path As Dictionary: Set path = hataori.ConnectBrowser(hvLastChild, pReqFile, pResFile, pPath)
    If Not path Is Nothing Then Set LastChild = New hataoriElement: LastChild.init path, pBrowserType
End Function

Public Function PreviousElementSibling() As hataoriElement
    Dim path As Dictionary: Set path = hataori.ConnectBrowser(hvPrevElement, pReqFile, pResFile, pPath)
    If Not path Is Nothing Then Set PreviousElementSibling = New hataoriElement: PreviousElementSibling.init path, pBrowserType
End Function

Public Function NextElementSibling() As hataoriElement
    Dim path As Dictionary: Set path = hataori.ConnectBrowser(hvNextElement, pReqFile, pResFile, pPath)
    If Not path Is Nothing Then Set NextElementSibling = New hataoriElement: NextElementSibling.init path, pBrowserType
End Function

'// Document
Public Function GetRootNode() As hataoriElement
    Dim path As Dictionary: Set path = hataori.ConnectBrowser(hvRoot, pReqFile, pResFile, pPath)
    If Not path Is Nothing Then Set GetRootNode = New hataoriElement: Root.init path, pBrowserType
End Function

Public Function GetCurrentRootNode() As hataoriElement
    Dim path As Dictionary: Set path = hataori.ConnectBrowser(hvCurrentRoot, pReqFile, pResFile, pPath)
    If Not path Is Nothing Then Set GetCurrentRootNode = New hataoriElement: GetRootNode.init path, pBrowserType
End Function

Public Function Document() As hataoriElement
    Dim path As Dictionary: Set path = hataori.ConnectBrowser(hvInnerContents, pReqFile, pResFile, pPath)
    If Not path Is Nothing Then Set Document = New hataoriElement: Document.init path, pBrowserType
End Function

Public Function Page() As hataoriPage
    Dim path As Dictionary: Set path = hataori.ConnectBrowser(hvInnerContents, pReqFile, pResFile, pPath)
    If Not path Is Nothing Then Set Page = New hataoriPage: Page.init path, pBrowserType
End Function

Public Function ParentNode() As hataoriElement
    Dim path As Dictionary: Set path = hataori.ConnectBrowser(hvParentNode, pReqFile, pResFile, pPath)
    If Not path Is Nothing Then Set ParentNode = New hataoriElement: ParentNode.init path, pBrowserType
End Function

Public Function ParentElement() As hataoriElement
    Dim path As Dictionary: Set path = hataori.ConnectBrowser(hvParentElement, pReqFile, pResFile, pPath)
    If Not path Is Nothing Then Set ParentElement = New hataoriElement: ParentElement.init path, pBrowserType
End Function

'// Set value, Get value
Public Function TagName() As String
    TagName = pTagName
End Function

Public Function Id() As String
    Id = pId
End Function

Public Function Name() As String
    Name = pName
End Function

Public Property Let Value(ByVal value_string As String)
    If Not hataori.ConnectBrowser(hvSetValue, pReqFile, pResFile, pPath, value_string) Then _
        Err.Raise vbObjectError + 12, "hataoriElement", hataoriSetting.GetErrorString(12)
End Property

Public Property Get Value() As String
    Value = hataori.ConnectBrowser(hvGetValue, pReqFile, pResFile, pPath)
End Property

Public Property Let InputValue(ByVal value_string As String)
    If Not hataori.ConnectBrowser(hvSetInput, pReqFile, pResFile, pPath, value_string) Then _
        Err.Raise vbObjectError + 12, "hataoriElement", hataoriSetting.GetErrorString(12)
End Property

Public Function OuterText() As String
    OuterText = hataori.ConnectBrowser(hvGetOuterText, pReqFile, pResFile, pPath)
End Function

Public Function OuterHTML() As String
    OuterHTML = hataori.ConnectBrowser(hvGetOuterHtml, pReqFile, pResFile, pPath)
End Function

Public Function InnerText() As String
    InnerText = hataori.ConnectBrowser(hvGetText, pReqFile, pResFile, pPath)
End Function

Public Function InnerHTML() As String
    InnerHTML = hataori.ConnectBrowser(hvGetHtml, pReqFile, pResFile, pPath)
End Function

Public Function GetCssValue(ByVal css_property_name As String) As String
    GetCssValue = hataori.ConnectBrowser(hvGetCss, pReqFile, pResFile, pPath, css_property_name)
End Function

Public Function ClassList() As Collection
    Set ClassList = hataori.ConnectBrowser(hvGetClassList, pReqFile, pResFile, pPath)
End Function

Public Function GetAttribute(ByVal attribute_property_name As String) As String
    GetAttribute = hataori.ConnectBrowser(hvGetAttr, pReqFile, pResFile, pPath, attribute_property_name)
End Function

'// Focus
Public Property Let Focus(ByVal focus_boolean As Boolean)
    If focus_boolean Then
        If Not hataori.ConnectBrowser(hvFocus, pReqFile, pResFile, pPath) Then _
            Err.Raise vbObjectError + 12, "hataoriElement", hataoriSetting.GetErrorString(12)
    Else
        If Not hataori.ConnectBrowser(hvBlur, pReqFile, pResFile, pPath) Then _
            Err.Raise vbObjectError + 12, "hataoriElement", hataoriSetting.GetErrorString(12)
    End If
End Property

'Form
Public Property Let Selected(ByVal selected_boolean As Boolean)
    If Not hataori.ConnectBrowser(hvSetSelected, pReqFile, pResFile, pPath, selected_boolean) Then _
        Err.Raise vbObjectError + 12, "hataoriElement", hataoriSetting.GetErrorString(12)
End Property

Public Property Get Selected() As Boolean
    Dim ret: ret = hataori.ConnectBrowser(hvGetSelected, pReqFile, pResFile, pPath)
    If IsEmpty(ret) Then _
        Err.Raise vbObjectError + 2, "hataoriElement", hataoriSetting.GetErrorString(2)
    Selected = ret
End Property

Public Property Let Checked(ByVal checked_boolean As Boolean)
    If Not hataori.ConnectBrowser(hvSetChecked, pReqFile, pResFile, pPath, checked_boolean) Then _
        Err.Raise vbObjectError + 12, "hataoriElement", hataoriSetting.GetErrorString(12)
End Property

Public Property Get Checked() As Boolean
    Dim ret: ret = hataori.ConnectBrowser(hvGetChecked, pReqFile, pResFile, pPath)
    If IsEmpty(ret) Then _
        Err.Raise vbObjectError + 2, "hataoriElement", hataoriSetting.GetErrorString(2)
    Checked = ret
End Property

Public Function Submit() As Boolean
    Submit = hataori.ConnectBrowser(hvSubmit, pReqFile, pResFile, pPath)
End Function

'// Mouse event
Public Sub Click()
    If Not hataori.ConnectBrowser(hvClick, pReqFile, pResFile, pPath) Then _
            Err.Raise vbObjectError + 12, "hataoriElement", hataoriSetting.GetErrorString(12)
End Sub

Public Sub DblClick()
    If Not hataori.ConnectBrowser(hvDblClick, pReqFile, pResFile, pPath) Then _
            Err.Raise vbObjectError + 12, "hataoriElement", hataoriSetting.GetErrorString(12)
End Sub

Public Sub RightClick()
    If Not hataori.ConnectBrowser(hvRClick, pReqFile, pResFile, pPath) Then _
            Err.Raise vbObjectError + 12, "hataoriElement", hataoriSetting.GetErrorString(12)
End Sub

Public Sub MouseDown()
    If Not hataori.ConnectBrowser(hvMouseDown, pReqFile, pResFile, pPath) Then _
            Err.Raise vbObjectError + 12, "hataoriElement", hataoriSetting.GetErrorString(12)
End Sub

Public Sub MouseUp()
    If Not hataori.ConnectBrowser(hvMouseUp, pReqFile, pResFile, pPath) Then _
            Err.Raise vbObjectError + 12, "hataoriElement", hataoriSetting.GetErrorString(12)
End Sub

'// Scroll
Public Sub ScrollTo(ByVal horizontal_coordinate As Long, ByVal vertical_coordinate As Long)
    If Not hataori.ConnectBrowser(hvElementScrollTo, pReqFile, pResFile, pPath, horizontal_coordinate, vertical_coordinate) Then _
            Err.Raise vbObjectError + 12, "hataoriElement", hataoriSetting.GetErrorString(12)
End Sub

Public Sub ScrollBy(ByVal horizontal_coordinate As Long, ByVal vertical_coordinate As Long)
    If Not hataori.ConnectBrowser(hvElementScrollBy, pReqFile, pResFile, pPath, horizontal_coordinate, vertical_coordinate) Then _
            Err.Raise vbObjectError + 12, "hataoriElement", hataoriSetting.GetErrorString(12)
End Sub

Public Sub FocusAndScroll()
    If Not hataori.ConnectBrowser(hvActive, pReqFile, pResFile, pPath) Then _
        Err.Raise vbObjectError + 12, "hataoriElement", hataoriSetting.GetErrorString(12)
End Sub

'// Upload file
Public Function UploadFile(ByVal upload_file_path As String) As Boolean
    UploadFile = False

    If Not hataori.ExistsUploadFile(upload_file_path) Then Exit Function

    If Not hataori.ConnectBrowser(hvFocus, pReqFile, pResFile, pPath, True) Then Exit Function
    
    If Not hataori.SendKeyStringToBrowser(pReqFile, pResFile, "enter", pBrowserCaption) Then Exit Function

    hataori.SecondsSleep 0.1

    Dim uploadCaption As Dictionary: Set uploadCaption = GethataoriBrowserSetting("upload_caption", pBrowserType)
    Dim systemLanguage As String: systemLanguage = GethataoriSetting("language")

    If Not hataori.SelectWindowByName(pReqFile, pResFile, uploadCaption(systemLanguage)) Then Exit Function

    If Not hataori.HostMessaging(pReqFile, pResFile, "host", "action", "set_clip", upload_file_path) Then Exit Function
    If Not hataori.HostMessaging(pReqFile, pResFile, "host", "action", "send_key", "alt+n") Then Exit Function
    If Not hataori.HostMessaging(pReqFile, pResFile, "host", "action", "send_key", "ctrl+v") Then Exit Function
    If Not hataori.HostMessaging(pReqFile, pResFile, "host", "action", "send_key", "alt+o") Then Exit Function
    If Not hataori.HostMessaging(pReqFile, pResFile, "host", "action", "set_clip", "") Then Exit Function
    
    UploadFile = True
End Function

'// Exists element
Public Function WaitElementExists( _
    ByVal css_string As String, _
    Optional element_index As Long = 0, _
    Optional ByVal timeout_seconds As Long = 10 _
) As Boolean
    WaitElementExists = hataori.WaitExistsElement(pReqFile, pResFile, pPath, css_string, element_index, timeout_seconds)
End Function

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
