Attribute VB_Name = "basMain"
Option Explicit

'設定画面チェックボックス状態
Public pintChkXls                   As Integer
Public pintChkTxt                   As Integer
Public pintChkCmt                   As Integer
Public pintChkFunc                  As Integer
Public pintChkDir                   As Integer
Public pintChkTrim                  As Integer

'INIファイル関連
Public pstrIniPath                  As String               'INIファイルフルパス

'共通クラスオブジェクト
Public pclsMsg                      As New clsMsg           'メッセージ出力クラス
Public pclsLog                      As New clsLog           'ログ出力クラス

'自動実行パラメータ
Private Const AUTO_PARA             As String = "/a"

Sub Main()
'起動時処理

    Dim strValue                    As String
    Dim strCmd                      As String
    Dim intRet                      As Integer

    'INIファイルのパスを変数に格納
    pstrIniPath = App.Path & DIR_SEPARATE & INI_FILENM

    'INIファイルからデータを取得
    strValue = StrIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_XLS)
    If IsNumeric(strValue) Then
        pintChkXls = CInt(strValue)
    Else
        pintChkXls = 0
    End If
    strValue = StrIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_TXT)
    If IsNumeric(strValue) Then
        pintChkTxt = CInt(strValue)
    Else
        pintChkTxt = 0
    End If
    strValue = StrIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_CMT)
    If IsNumeric(strValue) Then
        pintChkCmt = CInt(strValue)
    Else
        pintChkCmt = 0
    End If
    strValue = StrIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_FUNC)
    If IsNumeric(strValue) Then
        pintChkFunc = CInt(strValue)
    Else
        pintChkFunc = 0
    End If
    strValue = StrIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_DIR)
    If IsNumeric(strValue) Then
        pintChkDir = CInt(strValue)
    Else
        pintChkDir = 0
    End If
    strValue = StrIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_TRIM)
    If IsNumeric(strValue) Then
        pintChkTrim = CInt(strValue)
    Else
        pintChkTrim = 0
    End If

    strCmd = Command()
    intRet = InStr(1, strCmd, AUTO_PARA, vbTextCompare)
    If (intRet < 1) Then
        frmMain.Show
    Else
        Call AutoExec(strCmd)
    End If

End Sub

Private Sub AutoExec(ByVal vstrCmd As String)
'自動実行開始

    Dim intPos                      As Integer
    Dim strVbpNm                    As String

    'パラメータよりvbpファイルのパスを取得
    intPos = InStr(1, vstrCmd, AUTO_PARA) + Len(AUTO_PARA)
    strVbpNm = Replace(Trim(Mid(vstrCmd, intPos)), """", vbNullString)

    Call IsExec(strVbpNm)

End Sub

Public Function IsExec(ByVal vstrVbpPath As String) As Boolean
'仕様書作成開始

    Dim strVbpNm                    As String
    Dim blnRet                      As Boolean
    Dim lngIdx                      As Long
    Dim strPath                     As String
    Dim strMsg                      As String
    Dim lngRowNum                   As Long

    IsExec = False

    'vbp ファイルの存在チェック
    strVbpNm = Dir(vstrVbpPath)
    If (strVbpNm = vbNullString) Then
        Call pclsMsg.ShowMessage("指定されたフォルダに vbp ファイルが見つかりません。")
        Exit Function
    End If

    '関数件数の初期化
    plngFuncCnt = -1

    'vbp ファイルの読み込み
    blnRet = IsReadVbpFile(vstrVbpPath)
    If (blnRet = False) Then Exit Function

    '進捗画面の表示
    Load frmProgress
    With frmProgress
        .Show
        .Refresh
    End With

    'frm ファイルの読み込み
    If (plngFrmCnt > 0) Then
        strMsg = "フォームを読み込んでいます…"
        Call frmProgress.InitProgress(plngFrmCnt, strMsg)

        For lngIdx = 0 To plngFrmCnt - 1
            With ptypVbFile.Form(lngIdx)
                strPath = StrGetFilePath(vstrVbpPath) & DIR_SEPARATE & .FormPath _
                        & DIR_SEPARATE & .FormName
                blnRet = IsReadOtherFile(strPath, lngRowNum)
                If (blnRet = False) Then
                    Unload frmProgress
                    Exit Function
                End If
                .RowNum = lngRowNum
            End With
            Call frmProgress.Progress(lngIdx)
        Next lngIdx
    End If

    'bas ファイルの読み込み
    If (plngBasCnt > 0) Then
        strMsg = "標準モジュールを読み込んでいます…"
        Call frmProgress.InitProgress(plngBasCnt, strMsg)

        For lngIdx = 0 To plngBasCnt - 1
            With ptypVbFile.Module(lngIdx)
                strPath = StrGetFilePath(vstrVbpPath) & DIR_SEPARATE & .ModulePath _
                        & DIR_SEPARATE & .ModuleName
                blnRet = IsReadOtherFile(strPath, lngRowNum)
                If (blnRet = False) Then
                    Unload frmProgress
                    Exit Function
                End If
                .RowNum = lngRowNum
            End With
            Call frmProgress.Progress(lngIdx)
        Next lngIdx
    End If

    'cls ファイルの読み込み
    If (plngClsCnt > 0) Then
        strMsg = "クラス モジュールを読み込んでいます…"
        Call frmProgress.InitProgress(plngClsCnt, strMsg)

        For lngIdx = 0 To plngClsCnt - 1
            With ptypVbFile.Class(lngIdx)
                strPath = StrGetFilePath(vstrVbpPath) & DIR_SEPARATE & .ClassPath _
                        & DIR_SEPARATE & .ClassName
                blnRet = IsReadOtherFile(strPath, lngRowNum)
                If (blnRet = False) Then
                    Unload frmProgress
                    Exit Function
                End If
                .RowNum = lngRowNum
            End With
            Call frmProgress.Progress(lngIdx)
        Next lngIdx
    End If

    '進捗画面を閉じる
    Unload frmProgress

    '構造体の内容をエクセルテンプレートに展開
    If (pintChkXls = 1) Then
        blnRet = IsExcelOut()
        If (blnRet = False) Then
            Exit Function
        End If
    End If

    '構造体の内容をテキストに展開
    If (pintChkTxt = 1) Then
        blnRet = IsTextOut()
        If (blnRet = False) Then
            Exit Function
        End If
    End If

    IsExec = True
End Function

Private Function IsReadVbpFile(ByVal vstrVbpPath As String) As Boolean
'vbp ファイルの読み込み

    Dim intFileNo                   As Integer
    Dim varBuf                      As Variant
    Dim lngRet                      As Long
    Dim strChar                     As String
    Dim lngPos                      As Long

    On Error GoTo Exception

    IsReadVbpFile = False

    plngObjCnt = 0
    plngFrmCnt = 0
    plngBasCnt = 0
    plngClsCnt = 0

    intFileNo = FreeFile

    'vbp ファイルを開く
    Open vstrVbpPath For Input Access Read As intFileNo

    Do
    '全行読み切るまで以下の処理を続ける
        If EOF(intFileNo) Then
            Exit Do
        End If

        '1行読み込む
        Line Input #intFileNo, varBuf

        With ptypVbFile
            'タイトル
            lngRet = InStr(1, varBuf, VBP_TITLE, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_TITLE) + 1)
                strChar = Right(varBuf, lngPos)
                .TITLE = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            'プロジェクト名
            lngRet = InStr(1, varBuf, VBP_NAME, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_NAME) + 1)
                strChar = Right(varBuf, lngPos)
                .Name = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            '実行ファイル名
            lngRet = InStr(1, varBuf, VBP_EXENAME32, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_EXENAME32) + 1)
                strChar = Right(varBuf, lngPos)
                .ExeName32 = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            'パラメータ
            lngRet = InStr(1, varBuf, VBP_COMMAND32, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_COMMAND32) + 1)
                strChar = Right(varBuf, lngPos)
                .Command32 = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            'ヘルプファイル名
            lngRet = InStr(1, varBuf, VBP_HELPFILE, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_HELPFILE) + 1)
                strChar = Right(varBuf, lngPos)
                .HELPFILE = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            'コメント
            lngRet = InStr(1, varBuf, VBP_VERSIONCOMMENTS, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_VERSIONCOMMENTS) + 1)
                strChar = Right(varBuf, lngPos)
                .VersionComments = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            '説明
            lngRet = InStr(1, varBuf, VBP_VERSIONFILEDESCRIPTION, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_VERSIONFILEDESCRIPTION) + 1)
                strChar = Right(varBuf, lngPos)
                .VersionFileDescription = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            '会社名
            lngRet = InStr(1, varBuf, VBP_VERSIONCOMPANYNAME, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_VERSIONCOMPANYNAME) + 1)
                strChar = Right(varBuf, lngPos)
                .VersionCompanyName = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            '著作権
            lngRet = InStr(1, varBuf, VBP_VERSIONLEGALCOPYRIGHT, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_VERSIONLEGALCOPYRIGHT) + 1)
                strChar = Right(varBuf, lngPos)
                .VersionLegalCopyright = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            'バージョン Major
            lngRet = InStr(1, varBuf, VBP_MAJORVER, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_MAJORVER) + 1)
                strChar = Right(varBuf, lngPos)
                .MajorVer = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            'バージョン Minor
            lngRet = InStr(1, varBuf, VBP_MINORVER, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_MINORVER) + 1)
                strChar = Right(varBuf, lngPos)
                .MinorVer = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            'バージョン Revision
            lngRet = InStr(1, varBuf, VBP_REVISIONVER, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_REVISIONVER) + 1)
                strChar = Right(varBuf, lngPos)
                .RevisionVer = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            'コンポーネント
            lngRet = InStr(1, varBuf, VBP_OBJECT, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_FORM) + 1)
                strChar = Right(varBuf, lngPos)
                strChar = Replace(strChar, DBL_QUOTATION, vbNullString)
                lngRet = InStr(1, strChar, SEMICOLON, vbTextCompare)
                lngPos = Len(strChar) - lngRet
                .Object = Right(strChar, lngPos)
            End If
            'フォーム
            lngRet = InStr(1, varBuf, VBP_FORM, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_FORM) + 1)
                strChar = Right(varBuf, lngPos)
                strChar = Replace(strChar, DBL_QUOTATION, vbNullString)
                lngRet = InStr(1, strChar, EQUALMARK, vbTextCompare)
                lngPos = Len(strChar) - lngRet
                strChar = Right(strChar, lngPos)
                lngRet = InStr(1, strChar, PEARENT_DIR & PEARENT_DIR, vbTextCompare)
                '上位ディレクトリにあるファイルを読み飛ばす設定なら、それを読み飛ばす
                If (pintChkDir <> 1) Or (lngRet = 0) Then
                    ReDim Preserve .Form(plngFrmCnt)
                    .Form(plngFrmCnt).FormName = StrGetFileName(strChar)
                    .Form(plngFrmCnt).FormPath = StrGetFilePath(strChar)
                    plngFrmCnt = plngFrmCnt + 1
                End If
            End If
            'モジュール
            lngRet = InStr(1, varBuf, VBP_MODULE, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_MODULE) + 1)
                strChar = Right(varBuf, lngPos)
                strChar = Replace(strChar, DBL_QUOTATION, vbNullString)
                lngRet = InStr(1, strChar, SEMICOLON, vbTextCompare)
                lngPos = Len(strChar) - lngRet
                strChar = Right(strChar, lngPos)
                lngRet = InStr(1, strChar, PEARENT_DIR & PEARENT_DIR, vbTextCompare)
                '上位ディレクトリにあるファイルを読み飛ばす設定なら、それを読み飛ばす
                If (pintChkDir <> 1) Or (lngRet = 0) Then
                    ReDim Preserve .Module(plngBasCnt)
                    .Module(plngBasCnt).ModuleName = StrGetFileName(strChar)
                    .Module(plngBasCnt).ModulePath = StrGetFilePath(strChar)
                    plngBasCnt = plngBasCnt + 1
                End If
            End If
            'クラス
            lngRet = InStr(1, varBuf, VBP_CLASS, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_CLASS) + 1)
                strChar = Right(varBuf, lngPos)
                strChar = Replace(strChar, DBL_QUOTATION, vbNullString)
                lngRet = InStr(1, strChar, SEMICOLON, vbTextCompare)
                lngPos = Len(strChar) - lngRet
                strChar = Right(strChar, lngPos)
                lngRet = InStr(1, strChar, PEARENT_DIR & PEARENT_DIR, vbTextCompare)
                '上位ディレクトリにあるファイルを読み飛ばす設定なら、それを読み飛ばす
                If (pintChkDir <> 1) Or (lngRet = 0) Then
                    ReDim Preserve .Class(plngClsCnt)
                    .Class(plngClsCnt).ClassName = StrGetFileName(strChar)
                    .Class(plngClsCnt).ClassPath = StrGetFilePath(strChar)
                    plngClsCnt = plngClsCnt + 1
                End If
            End If
        End With
    Loop

    'vbp ファイルを閉じる
    Close intFileNo

    IsReadVbpFile = True

    Exit Function

Exception:
    Call pclsMsg.ShowError(Err.Description)
    On Error GoTo 0

End Function

Private Function IsReadOtherFile(ByVal vstrFileNm As String, ByRef rnumRowNum As Long) As Boolean
'vbp ファイル以外のVB ファイルを読む

    Dim intFileNo                   As Integer
    Dim varBuf                      As Variant
    Dim lngPos                      As Long
    Dim strCmt                      As String
    Dim blnStart                    As Boolean
    Dim blnExistKey                 As Boolean
    Dim lngIdx                      As Long
    Dim blnExistFn                  As Boolean
    Dim blnExistSub                 As Boolean
    Dim strFnNm                     As String
    Dim strScope                    As String
    Dim lngRet                      As Long
    Dim lngHosei                    As Long
    Dim strRet                      As String
    Dim strLine                     As String
    Dim blnContinue                 As Boolean
    Dim varContinue                 As Variant
    Dim strParam                    As String
    Dim strReturn                   As String

    On Error GoTo Exception

    IsReadOtherFile = False

    rnumRowNum = 0

    blnStart = False
    blnExistKey = False
    strLine = vbNullString
    varContinue = vbNullString

    intFileNo = FreeFile

    'ファイルを開く
    Open vstrFileNm For Input Access Read As intFileNo

    Do
    '全行読む
        If EOF(intFileNo) Then
            Exit Do
        End If

        '1行読み込む
        Line Input #intFileNo, varBuf

        ' [Attribute] 行の読み込むが完了したら、以降の行をプログラムソースとみなす
        lngRet = InStr(1, varBuf, "Attribute", vbTextCompare)
        If (lngRet > 0) Then
            blnExistKey = True
        Else
            If (blnExistKey = True) Then
                blnStart = True
            End If
        End If

        'ファイルごとのソースコードの行数をカウントする
        If (blnStart) Then
            rnumRowNum = rnumRowNum + 1
        End If

        'アンダーライン（"_"）が表示された場合は、次行と連結する
        If Right(varBuf, 2) = " _" Then
            blnContinue = True
            varContinue = varContinue & varBuf
        Else
            blnContinue = False
            varBuf = varContinue & varBuf
            varContinue = vbNullString
        End If

        If (blnStart) And Not (blnContinue) Then
            '"関数上部のコメントは、その下の関数を説明している"にチェックされているなら、関数外部のコメントも取得
            If (pintChkFunc = 1) And (blnExistFn = False) And (blnExistSub = False) Then
                'シングルクォーテーションが2つ以上続く行を読み飛ばす設定にされていない、もしくは2つ以上続くシングルクォーテーションが存在しなければ、コメントを取得
                lngRet = InStr(1, varBuf, SGL_QUOTATION & SGL_QUOTATION, vbTextCompare)
                If (pintChkCmt <> 1) Or (lngRet = 0) Then
                    'コメントを取得
                    lngRet = InStr(1, varBuf, SGL_QUOTATION, vbTextCompare)
                    If (lngRet > 0) Then
                        strLine = Right(varBuf, Len(varBuf) - lngRet)
                        If (pintChkTrim = 1) Then
                            lngPos = Len(varBuf) - Len(LTrim(varBuf))
                            For lngIdx = 0 To lngPos
                                strLine = Space(1) & strLine
                            Next lngIdx
                        End If
                        strCmt = strCmt & strLine & vbLf
                    End If
                End If
            End If

            'Function 関数を検索
            If (blnExistFn = False) Then
                lngRet = InStr(1, varBuf, "Public Function")
                If (lngRet = 1) Then
                    blnExistFn = True
                    lngHosei = Len("Public Function") + 1
                    strScope = "Public"
                Else
                    lngRet = InStr(1, varBuf, "Private Function")
                    If (lngRet = 1) Then
                        blnExistFn = True
                        lngHosei = Len("Private Function") + 1
                        strScope = "Private"
                    Else
                        strRet = Left(varBuf, 8)
                        If (strRet = "Function") Then
                            blnExistFn = True
                            lngHosei = Len("Function") + 1
                            strScope = "Public"
                        End If
                    End If
                End If
                'Function 関数名を取得
                If (blnExistFn = True) Then
                    strFnNm = Right(varBuf, Len(varBuf) - lngHosei)
                    lngRet = InStr(1, strFnNm, "(")
                    If (lngRet > 0) Then
                        strFnNm = Replace(strFnNm, "_ ", vbNullString)
                        strFnNm = Replace(strFnNm, "  ", vbNullString)
                        strParam = Mid(strFnNm, lngRet + 1, InStrRev(strFnNm, ")") - lngRet - 1)
                        strReturn = Right(strFnNm, Len(strFnNm) - InStrRev(strFnNm, "As") - Len("As"))
                        strFnNm = Trim(Left(strFnNm, lngRet - 1))
                        plngFuncCnt = plngFuncCnt + 1
                        ReDim Preserve ptypFunction(plngFuncCnt)
                        ptypFunction(plngFuncCnt).Name = strFnNm
                        ptypFunction(plngFuncCnt).Scope = strScope
                        ptypFunction(plngFuncCnt).FileName = StrGetFileName(vstrFileNm)
                        ptypFunction(plngFuncCnt).Param = strParam
                        ptypFunction(plngFuncCnt).Return = strReturn
                        Debug.Print vbNullString
                        Debug.Print "------------------------------------------------------------"
                        Debug.Print "-  関数名　　：" & strFnNm
                        Debug.Print "-  スコープ　：" & strScope
                        Debug.Print "-  ファイル名：" & StrGetFileName(vstrFileNm)
                        Debug.Print "-  パラメータ：" & strParam
                        Debug.Print "-  戻り値　　：" & strReturn
                        Debug.Print "------------------------------------------------------------"
                    End If
                End If
            Else
            'Function 関数の終了を検索
                lngRet = InStr(1, varBuf, "End Function")
                If (lngRet = 1) Then
                    Debug.Print strCmt
                    ptypFunction(plngFuncCnt).Comment = strCmt
                    strCmt = vbNullString
                    blnExistFn = False
                Else
                    'シングルクォーテーションが2つ以上続く行を読み飛ばす設定にされていない、もしくは2つ以上続くシングルクォーテーションが存在しなければ、コメントを取得
                    lngRet = InStr(1, varBuf, SGL_QUOTATION & SGL_QUOTATION, vbTextCompare)
                    If (pintChkCmt <> 1) Or (lngRet = 0) Then
                        'コメントを取得
                        lngRet = InStr(1, varBuf, SGL_QUOTATION, vbTextCompare)
                        If (lngRet > 0) Then
                            strLine = Right(varBuf, Len(varBuf) - lngRet)
                            If (pintChkTrim = 1) Then
                                lngPos = Len(varBuf) - Len(LTrim(varBuf))
                                For lngIdx = 0 To lngPos
                                    strLine = Space(1) & strLine
                                Next lngIdx
                            End If
                            strCmt = strCmt & strLine & vbLf
                        End If
                    End If
                End If
            End If
            'Sub 関数を検索
            If (blnExistSub = False) Then
                lngRet = InStr(1, varBuf, "Public Sub")
                If (lngRet = 1) Then
                    blnExistSub = True
                    lngHosei = Len("Public Sub") + 1
                    strScope = "Public"
                Else
                    lngRet = InStr(1, varBuf, "Private Sub")
                    If (lngRet = 1) Then
                        blnExistSub = True
                        lngHosei = Len("Private Sub") + 1
                        strScope = "Private"
                    Else
                        strRet = Left(varBuf, 3)
                        If (strRet = "Sub") Then
                            blnExistSub = True
                            lngHosei = Len("Sub") + 1
                            strScope = "Public"
                        End If
                    End If
                End If
                'Sub 関数名を取得
                If (blnExistSub = True) Then
                    strFnNm = Right(varBuf, Len(varBuf) - lngHosei)
                    lngRet = InStr(1, strFnNm, "(")
                    If (lngRet > 0) Then
                        strFnNm = Replace(strFnNm, "_ ", vbNullString)
                        strFnNm = Replace(strFnNm, "  ", vbNullString)
                        strParam = Mid(strFnNm, lngRet + 1, InStrRev(strFnNm, ")") - lngRet - 1)
                        strReturn = vbNullString
                        strFnNm = Trim(Left(strFnNm, lngRet - 1))
                        plngFuncCnt = plngFuncCnt + 1
                        ReDim Preserve ptypFunction(plngFuncCnt)
                        ptypFunction(plngFuncCnt).Name = strFnNm
                        ptypFunction(plngFuncCnt).Scope = strScope
                        ptypFunction(plngFuncCnt).FileName = StrGetFileName(vstrFileNm)
                        ptypFunction(plngFuncCnt).Param = strParam
                        ptypFunction(plngFuncCnt).Return = strReturn
                        Debug.Print vbNullString
                        Debug.Print "------------------------------------------------------------"
                        Debug.Print "-  関数名　　：" & strFnNm
                        Debug.Print "-  スコープ　：" & strScope
                        Debug.Print "-  ファイル名：" & StrGetFileName(vstrFileNm)
                        Debug.Print "-  パラメータ：" & strParam
                        Debug.Print "-  戻り値　　：" & strReturn
                        Debug.Print "------------------------------------------------------------"
                    End If
                End If
            Else
            'Sub 関数の終了を検索
                lngRet = InStr(1, varBuf, "End Sub")
                If (lngRet = 1) Then
                    Debug.Print strCmt
                    ptypFunction(plngFuncCnt).Comment = strCmt
                    strCmt = vbNullString
                    blnExistSub = False
                Else
                    'シングルクォーテーションが2つ以上続く行を読み飛ばす設定にされていない、もしくは2つ以上続くシングルクォーテーションが存在しなければ、コメントを取得
                    lngRet = InStr(1, varBuf, SGL_QUOTATION & SGL_QUOTATION, vbTextCompare)
                    If (pintChkCmt <> 1) Or (lngRet = 0) Then
                        'コメントを取得
                        lngRet = InStr(1, varBuf, SGL_QUOTATION, vbTextCompare)
                        If (lngRet > 0) Then
                            strLine = Right(varBuf, Len(varBuf) - lngRet)
                            If (pintChkTrim = 1) Then
                                lngPos = Len(varBuf) - Len(LTrim(varBuf))
                                For lngIdx = 0 To lngPos
                                    strLine = Space(1) & strLine
                                Next lngIdx
                            End If
                            strCmt = strCmt & strLine & vbLf
                        End If
                    End If
                End If
            End If
        End If
    Loop

    Close intFileNo

    IsReadOtherFile = True

    Exit Function

'例外処理
Exception:
    'エラーメッセージ表示
    Call pclsMsg.ShowError(Err.Description)
    On Error GoTo 0

End Function

Private Function IsExcelOut() As Boolean
'構造体の内容をエクセルテンプレートに展開

    Dim strXltPath                  As String
    Dim objExcel                    As Object
    Dim objWorkBook                 As Object
    Dim objWorksheet                As Object
    Dim objRange                    As Object
    Dim strRet                      As String
    Dim lngIdx                      As Long
    Dim strChar                     As String
    Dim lngRowPos                   As Long
    Dim lngRowHeight                As Long
    Dim strMsg                      As String

    Const XLTFILENM                 As String = "VBSpec.xlt"
    Const SHEET_HYOUSHI             As String = "表紙": Const HYOUSHI_POS = 1
    Const SHEET_FILE                As String = "ﾌｧｲﾙ構成": Const FILE_POS = 5
    Const SHEET_SYORI               As String = "処理概要": Const SYORI_POS = 4
    Const SHEET_FUNCTION            As String = "関数一覧": Const FUNCTION_POS = 8
    Const ROW_HEIGHT                As Long = 13.5
    Const ROW_START                 As Long = 5
    Const COL_START                 As String = "B"
    Const xlContinuous              As Long = 1
    Const xlDouble                  As Long = -4119
    Const xlEdgeLeft                As Long = 7
    Const xlEdgeTop                 As Long = 8
    Const xlEdgeBottom              As Long = 9
    Const xlEdgeRight               As Long = 10

    IsExcelOut = False

    On Error GoTo Exception

    strXltPath = App.Path & DIR_SEPARATE & XLTFILENM

    'テンプレートの存在チェック
    strRet = Dir(strXltPath)
    If (strRet = vbNullString) Then
        Call pclsMsg.ShowMessage("XLT ファイルが存在しません。")
        Exit Function
    End If

    '進捗画面の表示
    Load frmProgress
    With frmProgress
        .Show
        .Refresh
    End With

    'エクセルオブジェクト生成
    Set objExcel = CreateObject("Excel.Application")

    Set objWorkBook = objExcel.Workbooks.Add(strXltPath)

    '表紙
    Set objWorksheet = objWorkBook.Worksheets(HYOUSHI_POS)
    objWorksheet.Name = SHEET_HYOUSHI
    objWorksheet.Activate
    With objWorksheet
        'タイトル
        .Range("TITLE").Value = ptypVbFile.TITLE
        'バージョン
        .Range("VERSION").Value = "Version " & ptypVbFile.MajorVer & _
                                         "." & ptypVbFile.MinorVer & _
                                         "." & ptypVbFile.RevisionVer
        '開発会社
        .Range("COMPANY").Value = ptypVbFile.VersionCompanyName
        'プロジェクト名
        .Range("NAME").Value = ptypVbFile.Name
        '実行ファイル名
        .Range("EXENAME").Value = ptypVbFile.ExeName32
        'パラメータ
        .Range("COMMAND").Value = StrChgNullStr(ptypVbFile.Command32, MSG_NOTHING)
        'ヘルプファイル名
        .Range("HELPFILE").Value = StrChgNullStr(ptypVbFile.HELPFILE, MSG_NOTHING)
    End With

    'ﾌｧｲﾙ構成
    Set objWorksheet = objWorkBook.Worksheets(FILE_POS)
    objWorksheet.Activate
    objWorksheet.Name = SHEET_FILE
    With objWorksheet
        'フォーム
        strChar = vbNullString
        For lngIdx = 0 To plngFrmCnt - 1
            If (strChar = vbNullString) Then
                strChar = ptypVbFile.Form(lngIdx).FormName & " (" & CStr(ptypVbFile.Form(lngIdx).RowNum) & " 行)"
            Else
                strChar = strChar & vbLf & ptypVbFile.Form(lngIdx).FormName & " (" & CStr(ptypVbFile.Form(lngIdx).RowNum) & " 行)"
            End If
        Next lngIdx
        If (plngFrmCnt > 0) Then
            lngRowHeight = plngFrmCnt * ROW_HEIGHT
            On Error Resume Next
            .Range("FORM").RowHeight = lngRowHeight
            On Error GoTo 0
            On Error GoTo Exception
        End If
        .Range("FORM").Value = StrChgNullStr(strChar, MSG_NOTHING)
        'モジュール
        strChar = vbNullString
        For lngIdx = 0 To plngBasCnt - 1
            If (strChar = vbNullString) Then
                strChar = ptypVbFile.Module(lngIdx).ModuleName & " (" & CStr(ptypVbFile.Module(lngIdx).RowNum) & " 行)"
            Else
                strChar = strChar & vbLf & ptypVbFile.Module(lngIdx).ModuleName & " (" & CStr(ptypVbFile.Module(lngIdx).RowNum) & " 行)"
            End If
        Next lngIdx
        If (plngBasCnt > 0) Then
            lngRowHeight = plngBasCnt * ROW_HEIGHT
            On Error Resume Next
            .Range("MODULE").RowHeight = lngRowHeight
            On Error GoTo 0
            On Error GoTo Exception
        End If
        .Range("MODULE").Value = StrChgNullStr(strChar, MSG_NOTHING)
        'クラス
        strChar = vbNullString
        For lngIdx = 0 To plngClsCnt - 1
            If (strChar = vbNullString) Then
                strChar = ptypVbFile.Class(lngIdx).ClassName & " (" & CStr(ptypVbFile.Class(lngIdx).RowNum) & " 行)"
            Else
                strChar = strChar & vbLf & ptypVbFile.Class(lngIdx).ClassName & " (" & CStr(ptypVbFile.Class(lngIdx).RowNum) & " 行)"
            End If
        Next lngIdx
        If (plngClsCnt > 0) Then
            lngRowHeight = plngClsCnt * ROW_HEIGHT
            On Error Resume Next
            .Range("CLASS").RowHeight = lngRowHeight
            On Error GoTo 0
            On Error GoTo Exception
        End If
        .Range("CLASS").Value = StrChgNullStr(strChar, MSG_NOTHING)
    End With

    '処理概要
    Set objWorksheet = objWorkBook.Worksheets(SYORI_POS)
    objWorksheet.Activate
    objWorksheet.Name = SHEET_SYORI
    With objWorksheet
        '説明
        .Range("DESCRIPTION").Value = ptypVbFile.VersionFileDescription
    End With

    '関数一覧
    Set objWorksheet = objWorkBook.Worksheets(FUNCTION_POS)
    objWorksheet.Activate
    objWorksheet.Name = SHEET_FUNCTION
    With objWorksheet
        strMsg = "関数一覧を出力しています…"
        '関数が存在しなくても、関数一覧シートは作成する
        If (plngFuncCnt > -1) Then
            Call frmProgress.InitProgress(plngFuncCnt, strMsg)
            lngRowPos = ROW_START
            For lngIdx = 0 To plngFuncCnt
                .Range(COL_START & CStr(lngRowPos)).Value = ptypFunction(lngIdx).Name _
                                                    & "(" & ptypFunction(lngIdx).FileName & ")" & vbLf _
                                                    & "スコープ　：" & ptypFunction(lngIdx).Scope & vbLf _
                                                    & "パラメータ：" & ptypFunction(lngIdx).Param & vbLf _
                                                    & "戻り値　　：" & ptypFunction(lngIdx).Return & vbLf
                .Range(COL_START & CStr(lngRowPos)).Interior.Color = RGB(190, 190, 190)
                .Range(COL_START & CStr(lngRowPos)).Font.Color = RGB(255, 255, 255)
                lngRowHeight = IntCntLf(.Range(COL_START & CStr(lngRowPos)).Value) * ROW_HEIGHT
                .Range(COL_START & CStr(lngRowPos)).RowHeight = lngRowHeight
                lngRowPos = lngRowPos + 2
                strChar = ptypFunction(lngIdx).Comment
                .Range(COL_START & CStr(lngRowPos)).Value = strChar
                On Error Resume Next
                lngRowHeight = IntCntLf(strChar) * ROW_HEIGHT
                .Range(COL_START & CStr(lngRowPos)).RowHeight = lngRowHeight
                On Error GoTo 0
                On Error GoTo Exception
                lngRowPos = lngRowPos + 3
                Call frmProgress.Progress(lngIdx)
            Next lngIdx
        End If
    End With

    Unload frmProgress

    'エクセルファイルの表示
    Set objWorksheet = objWorkBook.Sheets(1)
    objWorksheet.Activate
    objExcel.Visible = True

    'エクセルオブジェクトの解放
    Set objRange = Nothing
    Set objWorksheet = Nothing
    Set objWorkBook = Nothing
    Set objExcel = Nothing

    IsExcelOut = True

    Exit Function

'例外処理
Exception:
    'エラーメッセージ表示
    Call pclsMsg.ShowError(Err.Description)
    Unload frmProgress
    'エクセルオブジェクトの解放
    Set objRange = Nothing
    Set objWorksheet = Nothing
    Set objWorkBook = Nothing
    Set objExcel = Nothing
    On Error GoTo 0

End Function

Private Function IsTextOut() As Boolean
'構造体の内容をテキストに展開

    Dim strChar                     As String
    Dim lngIdx                      As Long
    Dim strMsg                      As String
    Dim blnRet                      As Boolean
    Dim strFileNm                   As String
    Dim strDefFileNm                As String

    Const ITEM_START                As String = "    ・"

    On Error GoTo Exception

    IsTextOut = False

    '「ファイルを保存」ダイアログを表示し、テキストファイルの保存先を選択させる
    strDefFileNm = vbNullString
    blnRet = IsSaveTxtFileDlg(frmMain.cmnDlg, strFileNm, strDefFileNm)
    If (blnRet = False) Then
        IsTextOut = True
        Exit Function
    End If
    pclsLog.LogFileName = strFileNm
    pclsLog.Clear

    '進捗画面の表示
    Load frmProgress
    With frmProgress
        .Show
        .Refresh
    End With

    Call pclsLog.LogMsg("■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■")
    'タイトル & バージョン
    Call pclsLog.LogMsg("■  " & ptypVbFile.TITLE _
                               & "Version " & ptypVbFile.MajorVer & _
                                        "." & ptypVbFile.MinorVer & _
                                        "." & ptypVbFile.RevisionVer)
    '説明
    Call pclsLog.LogMsg("■" & vbNullString)
    Call pclsLog.LogMsg("■  " & ptypVbFile.VersionFileDescription)
    'プロジェクト名
    Call pclsLog.LogMsg("■" & vbNullString)
    Call pclsLog.LogMsg("■  プロジェクト名　：" & ptypVbFile.Name)
    '実行ファイル名
    Call pclsLog.LogMsg("■  実行ファイル名　：" & ptypVbFile.ExeName32)
    'パラメータ
    Call pclsLog.LogMsg("■  パラメータ　　　：" & StrChgNullStr(ptypVbFile.Command32, MSG_NOTHING))
    'ヘルプファイル名
    Call pclsLog.LogMsg("■  ヘルプファイル名：" & StrChgNullStr(ptypVbFile.HELPFILE, MSG_NOTHING))
    Call pclsLog.LogMsg("■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■")
    'フォーム
    Call pclsLog.LogMsg(vbNullString)
    Call pclsLog.LogMsg("○使用するフォームモジュール")
    strChar = vbNullString
    For lngIdx = 0 To plngFrmCnt - 1
        If (strChar = vbNullString) Then
            strChar = ITEM_START & ptypVbFile.Form(lngIdx).FormName & " (" & CStr(ptypVbFile.Form(lngIdx).RowNum) & " 行)"
        Else
            strChar = strChar & vbLf & ITEM_START & ptypVbFile.Form(lngIdx).FormName & " (" & CStr(ptypVbFile.Form(lngIdx).RowNum) & " 行)"
        End If
    Next lngIdx
    Call pclsLog.LogMsg(StrChgNullStr(strChar, MSG_NOTHING))
    'モジュール
    Call pclsLog.LogMsg(vbNullString)
    Call pclsLog.LogMsg("○使用する標準モジュール")
    strChar = vbNullString
    For lngIdx = 0 To plngBasCnt - 1
        If (strChar = vbNullString) Then
            strChar = ITEM_START & ptypVbFile.Module(lngIdx).ModuleName & " (" & CStr(ptypVbFile.Module(lngIdx).RowNum) & " 行)"
        Else
            strChar = strChar & vbLf & ITEM_START & ptypVbFile.Module(lngIdx).ModuleName & " (" & CStr(ptypVbFile.Module(lngIdx).RowNum) & " 行)"
        End If
    Next lngIdx
    Call pclsLog.LogMsg(StrChgNullStr(strChar, MSG_NOTHING))
    'クラス
    Call pclsLog.LogMsg(vbNullString)
    Call pclsLog.LogMsg("○使用するクラスモジュール")
    strChar = vbNullString
    For lngIdx = 0 To plngClsCnt - 1
        If (strChar = vbNullString) Then
            strChar = ITEM_START & ptypVbFile.Class(lngIdx).ClassName & " (" & CStr(ptypVbFile.Class(lngIdx).RowNum) & " 行)"
        Else
            strChar = strChar & vbLf & ITEM_START & ptypVbFile.Class(lngIdx).ClassName & " (" & CStr(ptypVbFile.Class(lngIdx).RowNum) & " 行)"
        End If
    Next lngIdx
    Call pclsLog.LogMsg(StrChgNullStr(strChar, MSG_NOTHING))
    '関数一覧
    Call pclsLog.LogMsg(vbNullString)
    If (plngFuncCnt > -1) Then
        strMsg = "関数一覧を出力しています…"
        Call frmProgress.InitProgress(plngFuncCnt, strMsg)
        For lngIdx = 0 To plngFuncCnt
            Call pclsLog.LogMsg("*****************************************")
            Call pclsLog.LogMsg("*" & vbNullString)
            Call pclsLog.LogMsg("*" & "    " & ptypFunction(lngIdx).Name & "(" & ptypFunction(lngIdx).FileName & ")")
            Call pclsLog.LogMsg("*" & "        スコープ　：" & ptypFunction(lngIdx).Scope)
            Call pclsLog.LogMsg("*" & "        パラメータ：" & ptypFunction(lngIdx).Param)
            Call pclsLog.LogMsg("*" & "        戻り値　　：" & ptypFunction(lngIdx).Return)
            Call pclsLog.LogMsg("*" & vbNullString)
            Call pclsLog.LogMsg("*****************************************")
            Call pclsLog.LogMsg(ptypFunction(lngIdx).Comment)
            Call frmProgress.Progress(lngIdx)
        Next lngIdx
    End If

    Unload frmProgress

    IsTextOut = True

    Exit Function

'例外処理
Exception:
    Call pclsMsg.ShowError(Err.Description)
    Unload frmProgress
    On Error GoTo 0

End Function

Private Function StrChgNullStr(ByVal vstrChk As String _
                             , ByVal vstrChg As String) As String
'引数1に指定した文字列が空白なら、引数2に指定した文字列を返す

    StrChgNullStr = vstrChk

    If (vstrChk = vbNullString) Then
        StrChgNullStr = vstrChg
    End If
End Function

Private Function IntCntLf(ByVal vstrChar As String) As Integer
'引数に含まれる改行文字列(vbLf)の数をカウントして返す

    Dim lngBef                      As Long
    Dim lngAft                      As Long
    Dim strAft                      As String
    Dim lngRet                      As Long

    lngBef = LenB(vstrChar)

    strAft = Replace(vstrChar, vbLf, vbNullString)
    lngAft = LenB(strAft)

    lngRet = lngBef - lngAft

    IntCntLf = lngRet / 2
End Function

