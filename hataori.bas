Attribute VB_Name = "hataori"
'***************************************************************
' hataori Ver 2023-05-18-01
'
' License: MIT License (http://www.opensource.org/licenses/mit-license.php)
'  (c) 2023 Fukasawa Takashi
'
' Library used:
'  The following libraries are used. Thanks.
'   VBA-JSON v2.3.1
'    License: MIT License (http://www.opensource.org/licenses/mit-license.php)
'    (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
'
' Note: Please enable the following in the reference settings.
'       * Microsoft Scripting Runtime
'       * Microsoft VBScript Regular Expressions 5.5
'***************************************************************

Option Explicit

' Public Enum lists
Enum hvBrowserType
    hvBrowserTypeChrome = 0
    hvBrowserTypeEdge = 1
    hvBrowserTypeNone = 255
End Enum

Enum hvWindowMode
    hvWindowModeNormal = 0
    hvWindowModeMaximized = 1
    hvWindowModeFullscreen = 2
End Enum

Enum hvMethodName
    hvPrintout = 21001
    hvHistoryBack = 21002
    hvHistoryForward = 21003
    hvClick = 21004
    hvDblClick = 21005
    hvRClick = 21006
    hvMouseDown = 21007
    hvMouseUp = 21008
    hvFocus = 21009
    hvActive = 21010
    hvBlur = 21011
    hvScrollTo = 21012
    hvScrollBy = 21013
    hvElementScrollTo = 21014
    hvElementScrollBy = 21015
    hvSetSelected = 21016
    hvSetChecked = 21017
    hvSubmit = 21018
    hvSetValue = 21019
    hvSetInput = 21020
    hvRoot = 21021
    hvCurrentRoot = 21022
    hvInnerContents = 21023
    hvParentNode = 21024
    hvParentElement = 21025
    hvPrevElement = 21026
    hvNextElement = 21027
    hvHead = 21028
    hvLastChild = 21029
    hvFirstChild = 21030
    hvCssSelector = 21031
    hvChildren = 21032
    hvBros = 21033
    hvBody = 21034
    hvForms = 21035
    hvGetCss = 21036
    hvGetHtml = 21037
    hvGetText = 21038
    hvGetOuterText = 21039
    hvGetOuterHtml = 21040
    hvGetClassList = 21041
    hvGetSelected = 21042
    hvGetChecked = 21043
    hvGetValue = 21044
    hvGetAttr = 21045
    hvGetTitle = 21046
    hvGetUrl = 21047
    hvDocumentStatus = 21048
    hvThisGetTab = 20049
    hvGetTabs = 20050
    hvNormalWindow = 20051
    hvMaximizedWindow = 20052
    hvFullscreenWindow = 20053
    hvNewTab = 20054
    hvReloadTab = 20055
    hvCloseTab = 20056
    hvJumpTab = 20057
    hvBrowserExit = 20058
    hvActiveElement = 21059
    hvPathToElement = 21060
End Enum

' https://learn.microsoft.com/en-us/windows/win32/api/synchapi/nf-synchapi-sleep
Private Declare PtrSafe Sub Sleep Lib "kernel32" ( _
    ByVal dwMilliseconds As LongPtr _
)

'// hataori main function
Public Function ConnectBrowser( _
    enum_method_name As hvMethodName, _
    req_file As String, _
    res_file As String, _
    ParamArray parameter_values() _
)
    If Len(req_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(res_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(req_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(res_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)

    Dim request: request = hataori_getMethodName(enum_method_name)
    If UBound(request) = -1 Then Err.Raise vbObjectError + 11, "hataori", hataoriSetting.GetErrorString(11)
    
    Dim group As String: group = request(0)
    Dim func As String: func = request(1)
    Dim mode As String: mode = request(2)
    Dim functionType As Long: functionType = request(3)
    
    Dim parameter_1: parameter_1 = Empty: If UBound(parameter_values) > -1 Then parameter_1 = parameter_values(0)
    Dim parameter_2: parameter_2 = Empty: If UBound(parameter_values) > 0 Then parameter_2 = parameter_values(1)
    Dim parameter_3: parameter_3 = Empty: If UBound(parameter_values) > 1 Then parameter_3 = parameter_values(2)
    
    Dim resp: resp = Array(hataori.HostMessaging(req_file, res_file, group, func, mode, parameter_1, parameter_2, parameter_3))
    
    Select Case functionType
        Case 2, 3, 6, 7
            Select Case TypeName(resp(0))
                Case "Null"
                    ConnectBrowser = Empty
                Case Else
                    ConnectBrowser = resp(0)
            End Select
        Case 1, 5
            Select Case TypeName(resp(0))
                Case "Collection"
                    Set ConnectBrowser = resp(0)
                Case "Null"
                    Set ConnectBrowser = New Collection
                Case Else
                    Err.Raise vbObjectError + 9, "hataori", hataoriSetting.GetErrorString(9)
            End Select
        Case 0
            Select Case TypeName(resp(0))
                Case "Dictionary"
                    Set ConnectBrowser = resp(0)
                Case "Null"
                    Set ConnectBrowser = Nothing
                Case Else
                    Err.Raise vbObjectError + 9, "hataori", hataoriSetting.GetErrorString(9)
            End Select
        Case 4
            Select Case TypeName(resp(0))
                Case "Dictionary"
                    Set ConnectBrowser = resp(0)
                Case "Null"
                    Set ConnectBrowser = New Dictionary
                Case Else
                    Err.Raise vbObjectError + 9, "hataori", hataoriSetting.GetErrorString(9)
            End Select
        Case Else
            Err.Raise vbObjectError + 10, "hataori", hataoriSetting.GetErrorString(10)
    End Select
End Function

Public Function WaitExistHost( _
    ByVal req_file As String, _
    ByVal res_file As String, _
    Optional ByVal timeout_seconds As Long = 10 _
) As Boolean
    If Len(req_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(res_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(req_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(res_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)

    WaitExistHost = False

    Dim ret
    Dim nowDateTime As Double: nowDateTime = hataori_getNowUnixEpoch()

    Do
        ret = hataori.HostMessaging(req_file, res_file, "host", "action", "exists", Empty, Empty, Empty, 0)

        If Not IsEmpty(ret) Then
            If TypeName(ret) <> "Boolean" Then Err.Raise vbObjectError + 2, "hataori", hataoriSetting.GetErrorString(2)
            WaitExistHost = ret
            Exit Do
        End If

        DoEvents
    Loop While hataori_getNowUnixEpoch() - nowDateTime <= timeout_seconds

    hataori.SecondsSleep 0.1
End Function

Public Function SelectForegroundTab( _
    ByVal req_file As String, _
    ByVal res_file As String, _
    ByVal browser_caption As String _
) As Boolean
    If Len(req_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(res_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(req_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(res_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(browser_caption) = 0 Then Err.Raise vbObjectError + 14, "hataori", hataoriSetting.GetErrorString(14)

    SelectForegroundTab = False

    Dim list As Dictionary: Set list = hataori.HostMessaging(req_file, res_file, "browser", "get_tabs", "this", "")
    If Not list.Exists("id") Then Exit Function
    Dim tabId As Double: tabId = list("id")
    Dim tabTitle As String: tabTitle = list("title")
    
    Dim leng As Long: leng = hataori.HostMessaging(req_file, res_file, "host", "window", "get_window_len", tabTitle & browser_caption)
    
    Dim windowPos As Long
    For windowPos = 0 To leng - 1
        If Not hataori.HostMessaging(req_file, res_file, "host", "window", "active_by_title", tabTitle & browser_caption, windowPos) Then Exit Function
        Set list = hataori.HostMessaging(req_file, res_file, "browser", "get_tabs", "this", "")
        If Not list.Exists("id") Then Exit Function
        If tabId = list("id") Then
            SelectForegroundTab = True
            Exit For
        End If
    Next windowPos
End Function

Public Function SelectWindowByName( _
    ByVal req_file As String, _
    ByVal res_file As String, _
    ByVal window_title As String _
) As Boolean
    If Len(req_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(res_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(req_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(res_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(window_title) = 0 Then Err.Raise vbObjectError + 14, "hataori", hataoriSetting.GetErrorString(14)

    SelectWindowByName = False

    Dim leng As Long: leng = hataori.HostMessaging(req_file, res_file, "host", "window", "get_window_len", window_title)
    If leng > 1 Then Exit Function
    
    If Not hataori.HostMessaging(req_file, res_file, "host", "window", "active_by_title", window_title, 0) Then Exit Function

    SelectWindowByName = True
End Function

Public Function WaitCompleteTab( _
    ByVal req_file As String, _
    ByVal res_file As String, _
    Optional ByVal wait_interaction As Boolean = True, _
    Optional ByVal timeout_seconds As Long = 10 _
) As Boolean
    If Len(req_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(res_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(req_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(res_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)

    WaitCompleteTab = False
    Dim startTime As Double: startTime = Timer
    Dim domStatusString As String: domStatusString = "interactive"
    If wait_interaction Then domStatusString = "complete"
    Do
        If GetTabStatusString(req_file, res_file) = "complete" Then
            If hataori.HostMessaging(req_file, res_file, "dom", "get_document_data", "get_status", "/") = domStatusString Then
                WaitCompleteTab = True
                Exit Do
            End If
        End If
        SecondsSleep 0.1
        DoEvents
    Loop While Timer - startTime <= timeout_seconds
End Function

Public Function WaitExistsElement( _
    ByVal req_file As String, _
    ByVal res_file As String, _
    ByVal axon_path As String, _
    ByVal css_string As String, _
    Optional ByVal element_index As Long = 0, _
    Optional ByVal timeout_seconds As Long = 10 _
) As Boolean
    If Len(req_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(res_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(req_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(res_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)

    WaitExistsElement = False
    
    Dim paths As Collection
    Dim start_time As Double: start_time = Timer
    Do
        Set paths = hataori.HostMessaging(req_file, res_file, "dom", "get_elements_selector", "css_selector", axon_path, css_string)
        If paths.Count > element_index Then
            WaitExistsElement = True
            Exit Do
        End If
        SecondsSleep 0.1
        DoEvents
    Loop While Timer - start_time <= timeout_seconds
End Function

Public Function ExistsUploadFile( _
    ByVal upload_file_path As String _
) As Boolean
    ExistsUploadFile = hataori_existsFilePath(upload_file_path)
End Function

Public Function SendKeyStringToBrowser( _
    ByVal req_file As String, _
    ByVal res_file As String, _
    ByVal key_string As String, _
    ByVal browser_caption As String _
) As Boolean
    If Len(req_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(res_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(req_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(res_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(browser_caption) = 0 Then Err.Raise vbObjectError + 14, "hataori", hataoriSetting.GetErrorString(14)
    
    SendKeyStringToBrowser = False
    If Not hataori.SelectForegroundTab(req_file, res_file, browser_caption) Then Exit Function
    If Not hataori.HostMessaging(req_file, res_file, "host", "action", "send_key", key_string) Then Exit Function
    SendKeyStringToBrowser = True
End Function

Public Function SendUnicodeStringToBrowser( _
    ByVal req_file As String, _
    ByVal res_file As String, _
    ByVal unicode_string As String, _
    ByVal browser_caption As String, _
    Optional ByVal send_input_mode As Boolean = False _
) As Boolean
    If Len(req_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(res_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(req_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(res_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(browser_caption) = 0 Then Err.Raise vbObjectError + 14, "hataori", hataoriSetting.GetErrorString(14)
    
    SendUnicodeStringToBrowser = False
    If Not hataori.SelectForegroundTab(req_file, res_file, browser_caption) Then Exit Function

    If send_input_mode Then
        If Not hataori.HostMessaging(req_file, res_file, "host", "action", "send_text", unicode_string) Then Exit Function
    Else
        If Not hataori.HostMessaging(req_file, res_file, "host", "action", "set_clip", unicode_string) Then Exit Function
        If Not hataori.HostMessaging(req_file, res_file, "host", "action", "send_key", "ctrl+v") Then Exit Function
        If Not hataori.HostMessaging(req_file, res_file, "host", "action", "set_clip", "") Then Exit Function
    End If

    SendUnicodeStringToBrowser = True
End Function

Public Function SendUnicodeStringToWindow( _
    ByVal req_file As String, _
    ByVal res_file As String, _
    ByVal unicode_string As String, _
    ByVal browser_caption As String _
) As Boolean
    If Len(req_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(res_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(req_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(res_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(browser_caption) = 0 Then Err.Raise vbObjectError + 14, "hataori", hataoriSetting.GetErrorString(14)
    
    SendUnicodeStringToWindow = False
    If Not hataori.HostMessaging(req_file, res_file, "host", "action", "send_text", unicode_string) Then Exit Function
    SendUnicodeStringToWindow = True
End Function

Public Function OpenBrowserApplication( _
    ByVal req_file As String, _
    ByVal res_file As String, _
    ByVal browser_path As String, _
    ByVal url_string As String, _
    Optional ByVal timeout_seconds As Long = 10 _
)
    If Len(req_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(res_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(browser_path) = 0 Then Err.Raise vbObjectError + 16, "hataori", hataoriSetting.GetErrorString(16)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(req_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(res_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsFilePath(browser_path) Then Err.Raise vbObjectError + 16, "hataori", hataoriSetting.GetErrorString(16)
    
    hataori_openApplication """" & browser_path & """ --new-window --kiosk-printing " & url_string
    OpenBrowserApplication = hataori.WaitExistHost(req_file, res_file, timeout_seconds)
End Function

Public Function HostMessaging( _
    ByVal req_file As String, _
    ByVal res_file As String, _
    ByVal group_name As String, _
    ByVal function_name As String, _
    Optional ByVal mode_string = Empty, _
    Optional ByVal parameter_1 = Empty, _
    Optional ByVal parameter_2 = Empty, _
    Optional ByVal parameter_3 = Empty, _
    Optional ByVal timout_seconds As Long = -1 _
)
    If Len(req_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Len(res_file) = 0 Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(req_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    If Not hataori_existsDirectoryPath(hataori_getDirectoryPath(res_file)) Then Err.Raise vbObjectError + 15, "hataori", hataoriSetting.GetErrorString(15)
    
    If hataori_existsFilePath(req_file) Then hataori_removeFile hataori_getShortFilePath(req_file)
    If hataori_existsFilePath(res_file) Then hataori_removeFile hataori_getShortFilePath(res_file)
    Dim modified As Date: modified = CDate("1970/01/01 00:00:00")
    
    ' ConvertToJson: VBA-JSON (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
    Dim messageString As String: _
        messageString = ConvertToJson(hataori_getMessagingDict(group_name, function_name, mode_string, parameter_1, parameter_2, parameter_3))

    If Not hataori_writeFile(req_file, messageString) Then Err.Raise vbObjectError + 4, "hataori", hataoriSetting.GetErrorString(4)

    Dim timeoutSeconds As Long: timeoutSeconds = timout_seconds
    If timout_seconds = -1 Then timeoutSeconds = hataoriSetting.GethataoriSetting("res_timeout_seconds")
    Dim nowDateTime As Double: nowDateTime = hataori_getNowUnixEpoch()
    
    Do While True
        If hataori_existsFilePath(res_file) Then If modified < hataori_getFileModified(res_file) Then Exit Do
        If hataori_getNowUnixEpoch() - nowDateTime > timeoutSeconds Then
            If timout_seconds = -1 Then Err.Raise vbObjectError + 5, "hataori", hataoriSetting.GetErrorString(5)
            HostMessaging = Empty: Exit Function
        End If
        Sleep 10
        DoEvents
    Loop

    Dim jsonText As String: jsonText = hataori_readFile(res_file)
    hataori_removeFile res_file
   
    ' ParseJson: VBA-JSON (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
    Dim jsonObject As Object: Set jsonObject = ParseJson(jsonText)

    If Not hataori_checkResponseDict(jsonObject) Then Err.Raise vbObjectError + 1, "hataori", hataoriSetting.GetErrorString(1)

    Select Case TypeName(jsonObject("ret"))
        Case "Collection", "Dictionary"
            Set HostMessaging = jsonObject("ret")
        Case "Integer", "Long", "Single", "Double", "Boolean", "String"
            HostMessaging = jsonObject("ret")
        Case "Null"
            HostMessaging = Null
        Case Else
            Err.Raise vbObjectError + 13, "hataori", hataoriSetting.GetErrorString(13)
    End Select
End Function

'// hataori comoon function
Public Sub SecondsSleep( _
    Optional ByVal seconds As Double = 0.5 _
)
    Sleep seconds * 1000
End Sub

'// hataori private main function
Private Function hataori_getMethodName(enum_method_name As hvMethodName)
    hataori_getMethodName = Array()

    ' Type
    Const hvElementObject = 0
    Const hvElementsObject = 1
    Const hvValueVariant = 2
    Const hvValueString = 3
    Const hvValueDictionary = 4
    Const hvValueCollection = 5
    Const hvValueBoolean = 6
    Const hvValueTrue = 7

    ' Group & Function & Mode & Return Type
    Select Case enum_method_name
        Case 21001
            hataori_getMethodName = Array("dom", "page_action", "printout", hvValueTrue)
        Case 21002
            hataori_getMethodName = Array("dom", "page_action", "history-back", hvValueTrue)
        Case 21003
            hataori_getMethodName = Array("dom", "page_action", "history-forward", hvValueTrue)
        Case 21004
            hataori_getMethodName = Array("dom", "element_action", "click", hvValueTrue)
        Case 21005
            hataori_getMethodName = Array("dom", "element_action", "dbl_click", hvValueTrue)
        Case 21006
            hataori_getMethodName = Array("dom", "element_action", "r_click", hvValueTrue)
        Case 21007
            hataori_getMethodName = Array("dom", "element_action", "mouse_down", hvValueTrue)
        Case 21008
            hataori_getMethodName = Array("dom", "element_action", "mouse_up", hvValueTrue)
        Case 21009
            hataori_getMethodName = Array("dom", "element_action", "focus", hvValueTrue)
        Case 21010
            hataori_getMethodName = Array("dom", "element_action", "active", hvValueTrue)
        Case 21011
            hataori_getMethodName = Array("dom", "element_action", "blur", hvValueTrue)
        Case 21012
            hataori_getMethodName = Array("dom", "scroll", "scroll_to", hvValueTrue)
        Case 21013
            hataori_getMethodName = Array("dom", "scroll", "scroll_by", hvValueTrue)
        Case 21014
            hataori_getMethodName = Array("dom", "scroll", "element_scroll_to", hvValueTrue)
        Case 21015
            hataori_getMethodName = Array("dom", "scroll", "element_scroll_by", hvValueTrue)
        Case 21016
            hataori_getMethodName = Array("dom", "set_element_data", "set_selected", hvValueBoolean)
        Case 21017
            hataori_getMethodName = Array("dom", "set_element_data", "set_checked", hvValueBoolean)
        Case 21018
            hataori_getMethodName = Array("dom", "set_element_data", "submit", hvValueBoolean)
        Case 21019
            hataori_getMethodName = Array("dom", "set_element_data", "set_value", hvValueBoolean)
        Case 21020
            hataori_getMethodName = Array("dom", "set_element_data", "set_input", hvValueBoolean)
        Case 21021
            hataori_getMethodName = Array("dom", "get_element", "root", hvElementObject)
        Case 21022
            hataori_getMethodName = Array("dom", "get_element", "current_root", hvElementObject)
        Case 21023
            hataori_getMethodName = Array("dom", "get_element", "inner_contents", hvElementObject)
        Case 21024
            hataori_getMethodName = Array("dom", "get_element", "parent_node", hvElementObject)
        Case 21025
            hataori_getMethodName = Array("dom", "get_element", "parent_element", hvElementObject)
        Case 21026
            hataori_getMethodName = Array("dom", "get_element", "prev_element", hvElementObject)
        Case 21027
            hataori_getMethodName = Array("dom", "get_element", "next_element", hvElementObject)
        Case 21028
            hataori_getMethodName = Array("dom", "get_element", "head", hvElementObject)
        Case 21029
            hataori_getMethodName = Array("dom", "get_element", "last_child", hvElementObject)
        Case 21030
            hataori_getMethodName = Array("dom", "get_element", "first_child", hvElementObject)
        Case 21031
            hataori_getMethodName = Array("dom", "get_elements_selector", "css_selector", hvElementsObject)
        Case 21032
            hataori_getMethodName = Array("dom", "get_elements", "children", hvElementsObject)
        Case 21033
            hataori_getMethodName = Array("dom", "get_elements", "bros", hvElementsObject)
        Case 21034
            hataori_getMethodName = Array("dom", "get_elements", "body", hvElementObject)
        Case 21035
            hataori_getMethodName = Array("dom", "get_elements", "forms", hvElementsObject)
        Case 21036
            hataori_getMethodName = Array("dom", "get_element_data", "get_css", hvValueVariant)
        Case 21037
            hataori_getMethodName = Array("dom", "get_element_data", "get_html", hvValueString)
        Case 21038
            hataori_getMethodName = Array("dom", "get_element_data", "get_text", hvValueString)
        Case 21039
            hataori_getMethodName = Array("dom", "get_element_data", "get_outer_text", hvValueString)
        Case 21040
            hataori_getMethodName = Array("dom", "get_element_data", "get_outer_html", hvValueString)
        Case 21041
            hataori_getMethodName = Array("dom", "get_element_data", "get_class_list", hvValueCollection)
        Case 21042
            hataori_getMethodName = Array("dom", "get_element_data", "get_selected", hvValueBoolean)
        Case 21043
            hataori_getMethodName = Array("dom", "get_element_data", "get_checked", hvValueBoolean)
        Case 21044
            hataori_getMethodName = Array("dom", "get_element_data", "get_value", hvValueString)
        Case 21045
            hataori_getMethodName = Array("dom", "get_element_data", "get_attr", hvValueString)
        Case 21046
            hataori_getMethodName = Array("dom", "get_document_data", "get_title", hvValueString)
        Case 21047
            hataori_getMethodName = Array("dom", "get_document_data", "get_url", hvValueString)
        Case 21048
            hataori_getMethodName = Array("dom", "get_document_data", "get_status", hvValueString)
        Case 20049
            hataori_getMethodName = Array("browser", "get_tabs", "this", hvValueDictionary)
        Case 20050
            hataori_getMethodName = Array("browser", "get_tabs", "", hvValueCollection)
        Case 20051
            hataori_getMethodName = Array("browser", "window_action", "normal", hvValueBoolean)
        Case 20052
            hataori_getMethodName = Array("browser", "window_action", "maximized", hvValueBoolean)
        Case 20053
            hataori_getMethodName = Array("browser", "window_action", "fullscreen", hvValueBoolean)
        Case 20054
            hataori_getMethodName = Array("browser", "tab_action", "new", hvValueBoolean)
        Case 20055
            hataori_getMethodName = Array("browser", "tab_action", "reload", hvValueBoolean)
        Case 20056
            hataori_getMethodName = Array("browser", "tab_action", "close", hvValueBoolean)
        Case 20057
            hataori_getMethodName = Array("browser", "tab_action", "jump", hvValueBoolean)
        Case 20058
            hataori_getMethodName = Array("browser", "browser_exit", "", hvValueBoolean)
        Case 21059
            hataori_getMethodName = Array("dom", "get_element", "active_element", hvElementObject)
        Case 21060
            hataori_getMethodName = Array("dom", "get_element", "path_to_element", hvElementObject)
    End Select
End Function

Private Function hataori_getMessagingDict( _
    ByVal group_name As String, _
    ByVal function_name As String, _
    Optional ByVal mode_string = Empty, _
    Optional ByVal parameter_1 = Empty, _
    Optional ByVal parameter_2 = Empty, _
    Optional ByVal parameter_3 = Empty _
) As Dictionary
    Set hataori_getMessagingDict = New Dictionary
    
    hataori_getMessagingDict.Add "gp", group_name
    hataori_getMessagingDict.Add "fc", function_name
   
    Dim paramDict As Dictionary: Set paramDict = New Dictionary
    Dim paramExists As Boolean: paramExists = False
    
    If Not IsEmpty(mode_string) Then paramDict.Add "v1", mode_string: paramExists = True
    If Not IsEmpty(parameter_1) Then paramDict.Add "v2", parameter_1: paramExists = True
    If Not IsEmpty(parameter_2) Then paramDict.Add "v3", parameter_2: paramExists = True
    If Not IsEmpty(parameter_3) Then paramDict.Add "v4", parameter_3: paramExists = True
    
    If paramExists Then hataori_getMessagingDict.Add "p", paramDict
End Function

Private Function hataori_checkResponseDict( _
    ByVal dict As Dictionary _
) As Boolean
    hataori_checkResponseDict = False
    If (Not dict.Exists("ret")) Or (Not dict.Exists("err")) Then Exit Function
    If Len(dict("err")) > 0 Then Err.Raise vbObjectError + 6, "hataori", hataoriSetting.GetErrorString(6, dict("err"))
    hataori_checkResponseDict = True
End Function

'// hataori private common function
Private Sub hataori_openApplication(application_path As String)
    CreateObject("WScript.Shell").Run application_path, 1, False
End Sub

Private Function hataori_getDirectoryPath( _
    ByVal file_path As String _
) As String
    hataori_getDirectoryPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(file_path)
End Function

Private Function hataori_existsDirectoryPath( _
    ByVal directory_path As String _
) As Boolean
    hataori_existsDirectoryPath = CreateObject("Scripting.FileSystemObject").FolderExists(directory_path)
End Function

Private Function hataori_existsFilePath( _
    ByVal file_path As String _
) As Boolean
    hataori_existsFilePath = CreateObject("Scripting.FileSystemObject").FileExists(file_path)
End Function

Private Sub hataori_removeFile( _
    ByVal file_path As String _
)
    If hataori_existsFilePath(file_path) Then CreateObject("Scripting.FileSystemObject").DeleteFile file_path
End Sub

Private Function hataori_getShortFilePath( _
    ByVal file_name As String _
) As String
    hataori_getShortFilePath = CreateObject("Scripting.FileSystemObject").getFile(file_name).ShortPath
End Function

Private Function hataori_getFileModified( _
    ByVal file_path As String _
) As Date
    hataori_getFileModified = CreateObject("Scripting.FileSystemObject").getFile(file_path).DateLastModified
End Function

Private Function hataori_writeFile( _
    ByVal file_path As String, _
    Optional ByVal write_data = Empty _
) As Boolean
    hataori_writeFile = False
    
    Dim fileNumber As Long: fileNumber = FreeFile
    
    Open file_path For Output Lock Read Write As #fileNumber
    Close #fileNumber

    If Not IsEmpty(write_data) Then
        Dim bytes() As Byte: bytes = write_data
        fileNumber = FreeFile
        Open file_path For Binary Access Read Write Lock Read Write As #fileNumber
            Put #fileNumber, 1, bytes
        Close #fileNumber
    End If
    
    hataori_writeFile = True
End Function

Private Function hataori_readFile( _
    ByVal file_path As String _
) As Byte()
    Dim fileNumber As Long: fileNumber = FreeFile
    Dim leng As Long
    Dim bytes() As Byte
    
    Do
        Sleep 10
        DoEvents
        
        Open file_path For Binary Access Read Write Lock Read Write As #fileNumber
            leng = LOF(fileNumber) - 1
            If leng > -1 Then
                ReDim bytes(leng)
                Get #fileNumber, 1, bytes
            End If
        Close #fileNumber
    Loop While leng = -1
    
    hataori_readFile = bytes
End Function

Private Function hataori_getNowUnixEpoch() As Double
    hataori_getNowUnixEpoch = DateDiff("s", "1970/1/1 9:00", Now)
End Function

