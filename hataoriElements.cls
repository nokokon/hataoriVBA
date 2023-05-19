VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "hataoriElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************
' hataoriElements Ver 2023-05-18-01
'
' License: MIT License (http://www.opensource.org/licenses/mit-license.php)
'  (c) 2023 Fukasawa Takashi
'
' Note: Please enable the following in the reference settings.
'       * Microsoft Scripting Runtime
'***************************************************************

Option Explicit

Private pPaths As Collection
Private pElements As Collection
Private pBrowserType As hvBrowserType
Private pBrowserPath As String
Private pBrowserCaption As String
Private pReqFile As String
Private pResFile As String

Public Function GetItem( _
    ByVal index _
) As hataoriElement
Attribute GetItem.VB_UserMemId = 0
    Dim elem As hataoriElement: Set elem = New hataoriElement
    Select Case TypeName(index)
        Case "Integer", "Long", "Single", "Double"
            If (pPaths.Count < index + 1) Or (index + 1 <= 0) Then Err.Raise vbObjectError + 7, "hataoriElements", hataoriSetting.GetErrorString(7)
            elem.init pPaths(index + 1), pBrowserType
        Case "String"
            Dim pathsCount As Long
            Dim attrValue
            Dim pathDict As Dictionary
            For pathsCount = 1 To pPaths.Count
                Set pathDict = pPaths(pathsCount)
                attrValue = pathDict("name")
                If attrValue = index Then elem.init pPaths(pathsCount), pBrowserType: Exit For
            Next pathsCount
        Case Else
            Err.Raise vbObjectError + 7, "hataoriElements", hataoriSetting.GetErrorString(7)
    End Select
    Set GetItem = elem
End Function

Public Function Length() As Long
    Length = pPaths.Count
End Function

Public Function iter() As IUnknown
Attribute iter.VB_UserMemId = -4
    Set pElements = New Collection
    Dim path As Dictionary
    For Each path In pPaths
        pElements.Add New hataoriElement
        pElements(pElements.Count).init path, pBrowserType
    Next path
    Set iter = pElements.[_NewEnum]
End Function

Public Sub init(elements_paths As Collection, ByVal browser_type As hvBrowserType)
    Set pPaths = elements_paths
    pBrowserType = browser_type
    pBrowserPath = hataoriSetting.GethataoriBrowserSetting("browser_path", browser_type)
    pBrowserCaption = hataoriSetting.GethataoriBrowserSetting("browser_caption", browser_type)
    pReqFile = hataoriSetting.GethataoriBrowserSetting("req_file", browser_type)
    pResFile = hataoriSetting.GethataoriBrowserSetting("res_file", browser_type)
End Sub

Private Sub Class_Initialize()
    Set pPaths = New Collection
    pBrowserType = 99999
    pBrowserPath = vbNullString
    pBrowserCaption = vbNullString
    pReqFile = vbNullString
    pResFile = vbNullString
End Sub

