Attribute VB_Name = "basFile"
Option Explicit

'INIファイルからデータを取得
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" _
    Alias "GetPrivateProfileStringA" _
   (ByVal lpAppName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

'INIファイルにデータを保存
Private Declare Function WritePrivateProfileString Lib "kernel32.dll" _
    Alias "WritePrivateProfileStringA" _
   (ByVal lpAppName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long

Public Function IsOpenVbpFileDlg(ByVal vobjCmnDlg As Object _
                               , ByRef rstrFileNm As String _
                               , ByVal rstrDefFileNm As String) As Boolean
'vbp ファイルを開く ダイアログを表示

    Dim intRet                      As Integer

    On Error GoTo Exception

    IsOpenVbpFileDlg = False
    rstrFileNm = vbNullString

    With vobjCmnDlg
        .CancelError = True
        .Flags = cdlOFNCreatePrompt Or _
                 cdlOFNHideReadOnly Or _
                 cdlOFNNoReadOnlyReturn Or _
                 cdlOFNOverwritePrompt
        .Filter = "プロジェクト ファイル (*.vbp)|*.vbp"
        .FilterIndex = 1
        .InitDir = DIR_SEPARATE
        .FileName = rstrDefFileNm
    End With

    On Error Resume Next

    'ダイアログを表示
    vobjCmnDlg.ShowOpen

    If (Err.Number = cdlCancel) Then
        'キャンセルされたら処理しない
        Exit Function
    Else
        If (Err.Number <> 0) Then
            Call pclsMsg.ShowError(Err.Description)
            Exit Function
        End If
    End If

    vobjCmnDlg.Parent.Refresh

    On Error GoTo 0
    On Error GoTo Exception

    'ダイアログで選択したファイル名を返す
    rstrFileNm = vobjCmnDlg.FileName

    IsOpenVbpFileDlg = True
    Exit Function

Exception:
    Call pclsMsg.ShowError(Err.Description)
    On Error GoTo 0

End Function

Public Function IsSaveTxtFileDlg(ByVal vobjCmnDlg As Object _
                               , ByRef rstrFileNm As String _
                               , ByVal rstrDefFileNm As String) As Boolean
'Txt ファイルを保存 ダイアログを表示
    Dim intRet                      As Integer

    On Error GoTo Exception

    IsSaveTxtFileDlg = False
    rstrFileNm = vbNullString

    With vobjCmnDlg
        .CancelError = True
        .Flags = cdlOFNCreatePrompt Or _
                 cdlOFNHideReadOnly Or _
                 cdlOFNNoReadOnlyReturn Or _
                 cdlOFNOverwritePrompt
        .Filter = "テキスト ファイル (*.txt)|*.txt"
        .FilterIndex = 1
        .InitDir = DIR_SEPARATE
        .FileName = rstrDefFileNm
    End With

    On Error Resume Next

    'ダイアログを表示
    vobjCmnDlg.ShowSave

    If (Err.Number = cdlCancel) Then
        'キャンセルされたら処理しない
        Exit Function
    Else
        If (Err.Number <> 0) Then
            Call pclsMsg.ShowError(Err.Description)
            Exit Function
        End If
    End If

    vobjCmnDlg.Parent.Refresh

    On Error GoTo 0
    On Error GoTo Exception

    'ダイアログで選択したファイル名を返す
    rstrFileNm = vobjCmnDlg.FileName

    IsSaveTxtFileDlg = True
    Exit Function

Exception:
    Call pclsMsg.ShowError(Err.Description)
    On Error GoTo 0

End Function

Public Function StrGetFilePath(ByVal vstrFilePath As String) As String
'引数に指定したファイルのフルパスからパスだけを返す

    Dim lngCharLength               As Long
    Dim lngCnt                      As Long
    Dim strWork                     As String
    Dim lngWork                     As Long
    Dim strPathWithoutFileName      As String

    StrGetFilePath = vbNullString
    strPathWithoutFileName = vbNullString

    If InStr(1, vstrFilePath, DIR_SEPARATE) <= 0 Then
        StrGetFilePath = vbNullString
        Exit Function
    End If

    lngCharLength = Len(vstrFilePath)

    For lngCnt = 0 To lngCharLength - 1
        lngWork = lngCharLength - lngCnt
        strWork = Mid(vstrFilePath, lngWork, 1)
        If strWork = DIR_SEPARATE Then
            strPathWithoutFileName = Left(vstrFilePath, lngWork - 1)
            Exit For
        End If
    Next lngCnt

    StrGetFilePath = Trim(strPathWithoutFileName)
End Function

Public Function StrGetFileName(ByVal vstrFilePath As String) As String
'引数に指定したファイルのフルパスからファイル名だけを返す

    Dim lngCharLength               As Long
    Dim lngCnt                      As Long
    Dim strWork                     As String
    Dim lngWork                     As Long
    Dim strFileName                 As String

    StrGetFileName = vbNullString
    strFileName = vbNullString

    If InStr(1, vstrFilePath, DIR_SEPARATE) <= 0 Then
        StrGetFileName = Trim(vstrFilePath)
        Exit Function
    End If

    lngCharLength = Len(vstrFilePath)

    For lngCnt = 0 To lngCharLength - 1
        lngWork = lngCharLength - lngCnt
        strWork = Mid(vstrFilePath, lngWork, 1)
        If strWork = DIR_SEPARATE Then
            strFileName = Right(vstrFilePath, lngCnt)
            Exit For
        End If
    Next lngCnt

    StrGetFileName = Trim(strFileName)
End Function

Public Function StrIniData(ByVal vstrIniPath As String _
                         , ByVal vstrIniSec As String _
                         , ByVal vstrIniKey As String) As String
'INIファイルからデータを読み込む

    Dim strChar                     As String * 1024
    Dim lngLen                      As Long
    Dim strRet                      As String
    Dim lngRet                      As Long

    On Error Resume Next

    StrIniData = vbNullString

    strChar = vbNullString

    strRet = Dir(vstrIniPath)
    If (strRet = vbNullString) Then Exit Function

    lngLen = Len(strChar)

    lngRet = GetPrivateProfileString(vstrIniSec _
                                   , vstrIniKey _
                                   , vbNullString _
                                   , strChar _
                                   , lngLen _
                                   , vstrIniPath)

    StrIniData = Left(strChar, InStr(strChar, vbNullChar) - 1)

    On Error GoTo 0
End Function

Public Sub SetIniData(ByVal vstrIniPath As String _
                    , ByVal vstrIniSec As String _
                    , ByVal vstrIniKey As String _
                    , ByVal vstrData As String)
'INIファイルにデータを保存する

    Dim strChar                     As String
    Dim lngRet                      As Long

    On Error Resume Next

    lngRet = WritePrivateProfileString(vstrIniSec _
                                     , vstrIniKey _
                                     , vbNullString _
                                     , vstrIniPath)

    lngRet = WritePrivateProfileString(vstrIniSec _
                                     , vstrIniKey _
                                     , vstrData _
                                     , vstrIniPath)

    On Error GoTo 0
End Sub

