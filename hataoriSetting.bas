Attribute VB_Name = "hataoriSetting"
'***************************************************************
' hataori setting Ver 2023-05-18-01
'
' License: MIT License (http://www.opensource.org/licenses/mit-license.php)
'  (c) 2023 Fukasawa Takashi
'
' Note: Please enable the following in the reference settings.
'       * Microsoft Scripting Runtime
'***************************************************************

Option Explicit

'// hataori install directory
Public Function GetInstallDirectoryPath() As String
    GetInstallDirectoryPath = "C:\*****\hataori-main\hataori"
End Function

'// hataori setting function
Public Function GethataoriSetting(ByVal key_name As String)
    Dim settingDict As Dictionary: Set settingDict = New Dictionary
    
    '# Error message language.
    '  Note: Japanese: "jp", English: "en-us"
        settingDict.Add "language", "jp"

    '# Response file timeout seconds.
        settingDict.Add "res_timeout_seconds", 10


    '# Below is the processing code.
        If Not settingDict.Exists(key_name) Then Err.Raise vbObjectError + 3, "hataoriSetting", hataoriSetting.GetErrorString(3)

        GethataoriSetting = settingDict(key_name)
End Function

'// hataori browser setting function
Public Function GethataoriBrowserSetting(ByVal key_name As String, ByVal browser_type As hvBrowserType)
    Dim settingDict As Dictionary: Set settingDict = New Dictionary
    Dim hataoriInstallDirectory As String

    '########################################
    '# hataori Install Directory.
        hataoriInstallDirectory = hataoriSetting.GetInstallDirectoryPath


    '########################################
    '# Web Browser Setting.
        '* Google Chrome
            If browser_type = 0 Then
                settingDict.Add "req_file", hataoriInstallDirectory & "\file\req" ' Request file path.
                settingDict.Add "res_file", hataoriInstallDirectory & "\file\res" ' Response file path.
                
                settingDict.Add "browser_caption", " - Google Chrome" ' Window caption

                settingDict.Add "upload_caption", New Dictionary
                settingDict("upload_caption").Add "en-us", "Open"
                settingDict("upload_caption").Add "jp", "�J��"
                
                settingDict.Add "browser_path", "C:\Program Files\Google\Chrome\Application\chrome.exe" ' Administrator install
                'settingDict.Add "browser_path", Environ("LOCALAPPDATA") & "\Google\Chrome\Application\chrome.exe"' User install
        
        '* Microsoft Edge
            ElseIf browser_type = 1 Then
                settingDict.Add "req_file", hataoriInstallDirectory & "\..\req" ' Request file path.
                settingDict.Add "res_file", hataoriInstallDirectory & "\..\res" ' Response file path.
                
                settingDict.Add "browser_caption", " - Microsoft Edge" ' Window caption

                settingDict.Add "upload_caption", New Dictionary
                settingDict("upload_caption").Add "en-us", "File Upload"
                settingDict("upload_caption").Add "jp", "�t�@�C���̃A�b�v���[�h"
                
                settingDict.Add "browser_path", "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" ' Administrator(32bit) install
                'settingDict.Add "browser_path", "C:\Program Files\Microsoft\Edge\Application\msedge.exe" ' Administrator(64bit) install
                'settingDict.Add "browser_path", Environ("LOCALAPPDATA") & "\Microsoft\msedge.exe"' User install
        
        '* Etc
            Else
                Err.Raise vbObjectError + 8, "hataoriSetting", hataoriSetting.GetErrorString(8)
            End If


        If Not settingDict.Exists(key_name) Then Err.Raise vbObjectError + 3, "hataoriSetting", hataoriSetting.GetErrorString(3)

        Select Case TypeName(settingDict(key_name))
            Case "Dictionary", "Collection"
                Set GethataoriBrowserSetting = settingDict(key_name)
            Case Else
                GethataoriBrowserSetting = settingDict(key_name)
        End Select
End Function

'// hataori message language definition
Private Function hataoriErrorString(error_key) As String
    'Note: * Key of the error message for each language is the following structure.
    '          Key -> [Language]_[ErrorIndex]  Example: en_1, jp_1
    '      * You can specify 1 or more replacement strings.
    '          1 -> __{param1}__
    '          2 -> __{param2}__ ...
    
    '# Common
        Dim errorDict As New Dictionary: errorDict.Add "all_1024", "Error key not found."
    
    '# Japanese
        errorDict.Add "jp_1", "�߂�l�� 'ret' �L�[�܂��� 'err' �L�[��������܂���B"
        errorDict.Add "jp_2", "�l��Empty�A�܂���Null�ł��i���҂����l�� True �� False �ł��j�B"
        errorDict.Add "jp_3", "�L�[��������܂���B"
        errorDict.Add "jp_4", "���X�|���X�t�@�C���̏������݂Ɏ��s���܂����B"
        errorDict.Add "jp_5", "���X�|���X���^�C���A�E�g���܂����B"
        errorDict.Add "jp_6", "'err' �L�[���ɕ����񂪑��݂��܂��B" & vbCrLf & "Error:" & vbCrLf & "__{param1}__"
        errorDict.Add "jp_7", "�C���f�b�N�X�����݂��܂���B"
        errorDict.Add "jp_8", "�u���E�U�[�̎�ނ��s���ł��B"
        errorDict.Add "jp_9", "�߂�l���s���ł��B"
        errorDict.Add "jp_10", "�Ăяo��API������`�ł��B"
        errorDict.Add "jp_11", "���N�G�X�g���郁�\�b�h�̎�ނ�����`�ł��B"
        errorDict.Add "jp_12", "�߂�l�� True �ȊO�ł��B"
        errorDict.Add "jp_13", "�߂�l�̌^���s���ł��B"
        errorDict.Add "jp_14", "�u���E�U�[�L���v�V���������݂��܂���B"
        errorDict.Add "jp_15", "���N�G�X�g�E���X�|���X�t�@�C���̃t�H���_�[�����݂��܂���B"
        errorDict.Add "jp_16", "�u���E�U�[�̎��s�t�@�C�������݂��܂���B"
        errorDict.Add "jp_17", "�v�f��������܂���B"

    '# English
        errorDict.Add "en-us_1", "No 'ret' or 'err' key found in return value."
        errorDict.Add "en-us_2", "The value is Empty or Null (expected value is True or False)."
        errorDict.Add "en-us_3", "Key not found."
        errorDict.Add "en-us_4", "Failed to write response file."
        errorDict.Add "en-us_5", "Response timed out."
        errorDict.Add "en-us_6", "String exists in the 'err' key." & vbCrLf & "Error:" & vbCrLf & "__{param1}__"
        errorDict.Add "en-us_7", "Index does not exist."
        errorDict.Add "en-us_8", "Unknown browser type."
        errorDict.Add "en-us_9", "Invalid return value."
        errorDict.Add "en-us_10", "Calling API is undefined."
        errorDict.Add "en-us_11", "Type of method to request is undefined."
        errorDict.Add "en-us_12", "Return value is not True."
        errorDict.Add "en-us_13", "Return type is unknown."
        errorDict.Add "en-us_14", "Browser caption does not exist."
        errorDict.Add "en-us_15", "Request/response file folder does not exist."
        errorDict.Add "en-us_16", "Browser executable file does not exist."
        errorDict.Add "en-us_17", "Element not found."


    '# Below is the processing code.
        If Not errorDict.Exists(error_key) Then _
            Err.Raise vbObjectError + 1024, "hataoriSetting", errorDict("all_1024")

        hataoriErrorString = errorDict(error_key)
End Function

'# Below is the processing code.
    '// hataori get message
    Public Function GetErrorString(ByVal error_code As Long, ParamArray param_string()) As String
        Dim languageString: languageString = hataoriSetting.GethataoriSetting("language")
        Dim ret: ret = hataoriSetting.hataoriErrorString(languageString & "_" & CStr(error_code))
        Dim paramLeng As Long: paramLeng = UBound(param_string)
        Dim paramPos As Long
        For paramPos = 0 To paramLeng
            ret = Replace(ret, "__{param" & CStr(paramPos + 1) & "}__", param_string(paramLeng))
        Next paramPos
        
        GetErrorString = ret
    End Function

