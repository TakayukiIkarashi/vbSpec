Attribute VB_Name = "basMain"
Option Explicit

'�ݒ��ʃ`�F�b�N�{�b�N�X���
Public pintChkXls                   As Integer
Public pintChkTxt                   As Integer
Public pintChkCmt                   As Integer
Public pintChkFunc                  As Integer
Public pintChkDir                   As Integer
Public pintChkTrim                  As Integer

'INI�t�@�C���֘A
Public pstrIniPath                  As String               'INI�t�@�C���t���p�X

'���ʃN���X�I�u�W�F�N�g
Public pclsMsg                      As New clsMsg           '���b�Z�[�W�o�̓N���X
Public pclsLog                      As New clsLog           '���O�o�̓N���X

'�������s�p�����[�^
Private Const AUTO_PARA             As String = "/a"

Sub Main()
'�N��������

    Dim strValue                    As String
    Dim strCmd                      As String
    Dim intRet                      As Integer

    'INI�t�@�C���̃p�X��ϐ��Ɋi�[
    pstrIniPath = App.Path & DIR_SEPARATE & INI_FILENM

    'INI�t�@�C������f�[�^���擾
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
'�������s�J�n

    Dim intPos                      As Integer
    Dim strVbpNm                    As String

    '�p�����[�^���vbp�t�@�C���̃p�X���擾
    intPos = InStr(1, vstrCmd, AUTO_PARA) + Len(AUTO_PARA)
    strVbpNm = Replace(Trim(Mid(vstrCmd, intPos)), """", vbNullString)

    Call IsExec(strVbpNm)

End Sub

Public Function IsExec(ByVal vstrVbpPath As String) As Boolean
'�d�l���쐬�J�n

    Dim strVbpNm                    As String
    Dim blnRet                      As Boolean
    Dim lngIdx                      As Long
    Dim strPath                     As String
    Dim strMsg                      As String
    Dim lngRowNum                   As Long

    IsExec = False

    'vbp �t�@�C���̑��݃`�F�b�N
    strVbpNm = Dir(vstrVbpPath)
    If (strVbpNm = vbNullString) Then
        Call pclsMsg.ShowMessage("�w�肳�ꂽ�t�H���_�� vbp �t�@�C����������܂���B")
        Exit Function
    End If

    '�֐������̏�����
    plngFuncCnt = -1

    'vbp �t�@�C���̓ǂݍ���
    blnRet = IsReadVbpFile(vstrVbpPath)
    If (blnRet = False) Then Exit Function

    '�i����ʂ̕\��
    Load frmProgress
    With frmProgress
        .Show
        .Refresh
    End With

    'frm �t�@�C���̓ǂݍ���
    If (plngFrmCnt > 0) Then
        strMsg = "�t�H�[����ǂݍ���ł��܂��c"
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

    'bas �t�@�C���̓ǂݍ���
    If (plngBasCnt > 0) Then
        strMsg = "�W�����W���[����ǂݍ���ł��܂��c"
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

    'cls �t�@�C���̓ǂݍ���
    If (plngClsCnt > 0) Then
        strMsg = "�N���X ���W���[����ǂݍ���ł��܂��c"
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

    '�i����ʂ����
    Unload frmProgress

    '�\���̂̓��e���G�N�Z���e���v���[�g�ɓW�J
    If (pintChkXls = 1) Then
        blnRet = IsExcelOut()
        If (blnRet = False) Then
            Exit Function
        End If
    End If

    '�\���̂̓��e���e�L�X�g�ɓW�J
    If (pintChkTxt = 1) Then
        blnRet = IsTextOut()
        If (blnRet = False) Then
            Exit Function
        End If
    End If

    IsExec = True
End Function

Private Function IsReadVbpFile(ByVal vstrVbpPath As String) As Boolean
'vbp �t�@�C���̓ǂݍ���

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

    'vbp �t�@�C�����J��
    Open vstrVbpPath For Input Access Read As intFileNo

    Do
    '�S�s�ǂݐ؂�܂ňȉ��̏����𑱂���
        If EOF(intFileNo) Then
            Exit Do
        End If

        '1�s�ǂݍ���
        Line Input #intFileNo, varBuf

        With ptypVbFile
            '�^�C�g��
            lngRet = InStr(1, varBuf, VBP_TITLE, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_TITLE) + 1)
                strChar = Right(varBuf, lngPos)
                .TITLE = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            '�v���W�F�N�g��
            lngRet = InStr(1, varBuf, VBP_NAME, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_NAME) + 1)
                strChar = Right(varBuf, lngPos)
                .Name = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            '���s�t�@�C����
            lngRet = InStr(1, varBuf, VBP_EXENAME32, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_EXENAME32) + 1)
                strChar = Right(varBuf, lngPos)
                .ExeName32 = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            '�p�����[�^
            lngRet = InStr(1, varBuf, VBP_COMMAND32, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_COMMAND32) + 1)
                strChar = Right(varBuf, lngPos)
                .Command32 = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            '�w���v�t�@�C����
            lngRet = InStr(1, varBuf, VBP_HELPFILE, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_HELPFILE) + 1)
                strChar = Right(varBuf, lngPos)
                .HELPFILE = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            '�R�����g
            lngRet = InStr(1, varBuf, VBP_VERSIONCOMMENTS, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_VERSIONCOMMENTS) + 1)
                strChar = Right(varBuf, lngPos)
                .VersionComments = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            '����
            lngRet = InStr(1, varBuf, VBP_VERSIONFILEDESCRIPTION, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_VERSIONFILEDESCRIPTION) + 1)
                strChar = Right(varBuf, lngPos)
                .VersionFileDescription = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            '��Ж�
            lngRet = InStr(1, varBuf, VBP_VERSIONCOMPANYNAME, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_VERSIONCOMPANYNAME) + 1)
                strChar = Right(varBuf, lngPos)
                .VersionCompanyName = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            '���쌠
            lngRet = InStr(1, varBuf, VBP_VERSIONLEGALCOPYRIGHT, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_VERSIONLEGALCOPYRIGHT) + 1)
                strChar = Right(varBuf, lngPos)
                .VersionLegalCopyright = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            '�o�[�W���� Major
            lngRet = InStr(1, varBuf, VBP_MAJORVER, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_MAJORVER) + 1)
                strChar = Right(varBuf, lngPos)
                .MajorVer = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            '�o�[�W���� Minor
            lngRet = InStr(1, varBuf, VBP_MINORVER, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_MINORVER) + 1)
                strChar = Right(varBuf, lngPos)
                .MinorVer = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            '�o�[�W���� Revision
            lngRet = InStr(1, varBuf, VBP_REVISIONVER, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_REVISIONVER) + 1)
                strChar = Right(varBuf, lngPos)
                .RevisionVer = Replace(strChar, DBL_QUOTATION, vbNullString)
            End If
            '�R���|�[�l���g
            lngRet = InStr(1, varBuf, VBP_OBJECT, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_FORM) + 1)
                strChar = Right(varBuf, lngPos)
                strChar = Replace(strChar, DBL_QUOTATION, vbNullString)
                lngRet = InStr(1, strChar, SEMICOLON, vbTextCompare)
                lngPos = Len(strChar) - lngRet
                .Object = Right(strChar, lngPos)
            End If
            '�t�H�[��
            lngRet = InStr(1, varBuf, VBP_FORM, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_FORM) + 1)
                strChar = Right(varBuf, lngPos)
                strChar = Replace(strChar, DBL_QUOTATION, vbNullString)
                lngRet = InStr(1, strChar, EQUALMARK, vbTextCompare)
                lngPos = Len(strChar) - lngRet
                strChar = Right(strChar, lngPos)
                lngRet = InStr(1, strChar, PEARENT_DIR & PEARENT_DIR, vbTextCompare)
                '��ʃf�B���N�g���ɂ���t�@�C����ǂݔ�΂��ݒ�Ȃ�A�����ǂݔ�΂�
                If (pintChkDir <> 1) Or (lngRet = 0) Then
                    ReDim Preserve .Form(plngFrmCnt)
                    .Form(plngFrmCnt).FormName = StrGetFileName(strChar)
                    .Form(plngFrmCnt).FormPath = StrGetFilePath(strChar)
                    plngFrmCnt = plngFrmCnt + 1
                End If
            End If
            '���W���[��
            lngRet = InStr(1, varBuf, VBP_MODULE, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_MODULE) + 1)
                strChar = Right(varBuf, lngPos)
                strChar = Replace(strChar, DBL_QUOTATION, vbNullString)
                lngRet = InStr(1, strChar, SEMICOLON, vbTextCompare)
                lngPos = Len(strChar) - lngRet
                strChar = Right(strChar, lngPos)
                lngRet = InStr(1, strChar, PEARENT_DIR & PEARENT_DIR, vbTextCompare)
                '��ʃf�B���N�g���ɂ���t�@�C����ǂݔ�΂��ݒ�Ȃ�A�����ǂݔ�΂�
                If (pintChkDir <> 1) Or (lngRet = 0) Then
                    ReDim Preserve .Module(plngBasCnt)
                    .Module(plngBasCnt).ModuleName = StrGetFileName(strChar)
                    .Module(plngBasCnt).ModulePath = StrGetFilePath(strChar)
                    plngBasCnt = plngBasCnt + 1
                End If
            End If
            '�N���X
            lngRet = InStr(1, varBuf, VBP_CLASS, vbTextCompare)
            If (lngRet = 1) Then
                lngPos = Len(varBuf) - (Len(VBP_CLASS) + 1)
                strChar = Right(varBuf, lngPos)
                strChar = Replace(strChar, DBL_QUOTATION, vbNullString)
                lngRet = InStr(1, strChar, SEMICOLON, vbTextCompare)
                lngPos = Len(strChar) - lngRet
                strChar = Right(strChar, lngPos)
                lngRet = InStr(1, strChar, PEARENT_DIR & PEARENT_DIR, vbTextCompare)
                '��ʃf�B���N�g���ɂ���t�@�C����ǂݔ�΂��ݒ�Ȃ�A�����ǂݔ�΂�
                If (pintChkDir <> 1) Or (lngRet = 0) Then
                    ReDim Preserve .Class(plngClsCnt)
                    .Class(plngClsCnt).ClassName = StrGetFileName(strChar)
                    .Class(plngClsCnt).ClassPath = StrGetFilePath(strChar)
                    plngClsCnt = plngClsCnt + 1
                End If
            End If
        End With
    Loop

    'vbp �t�@�C�������
    Close intFileNo

    IsReadVbpFile = True

    Exit Function

Exception:
    Call pclsMsg.ShowError(Err.Description)
    On Error GoTo 0

End Function

Private Function IsReadOtherFile(ByVal vstrFileNm As String, ByRef rnumRowNum As Long) As Boolean
'vbp �t�@�C���ȊO��VB �t�@�C����ǂ�

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

    '�t�@�C�����J��
    Open vstrFileNm For Input Access Read As intFileNo

    Do
    '�S�s�ǂ�
        If EOF(intFileNo) Then
            Exit Do
        End If

        '1�s�ǂݍ���
        Line Input #intFileNo, varBuf

        ' [Attribute] �s�̓ǂݍ��ނ�����������A�ȍ~�̍s���v���O�����\�[�X�Ƃ݂Ȃ�
        lngRet = InStr(1, varBuf, "Attribute", vbTextCompare)
        If (lngRet > 0) Then
            blnExistKey = True
        Else
            If (blnExistKey = True) Then
                blnStart = True
            End If
        End If

        '�t�@�C�����Ƃ̃\�[�X�R�[�h�̍s�����J�E���g����
        If (blnStart) Then
            rnumRowNum = rnumRowNum + 1
        End If

        '�A���_�[���C���i"_"�j���\�����ꂽ�ꍇ�́A���s�ƘA������
        If Right(varBuf, 2) = " _" Then
            blnContinue = True
            varContinue = varContinue & varBuf
        Else
            blnContinue = False
            varBuf = varContinue & varBuf
            varContinue = vbNullString
        End If

        If (blnStart) And Not (blnContinue) Then
            '"�֐��㕔�̃R�����g�́A���̉��̊֐���������Ă���"�Ƀ`�F�b�N����Ă���Ȃ�A�֐��O���̃R�����g���擾
            If (pintChkFunc = 1) And (blnExistFn = False) And (blnExistSub = False) Then
                '�V���O���N�H�[�e�[�V������2�ȏ㑱���s��ǂݔ�΂��ݒ�ɂ���Ă��Ȃ��A��������2�ȏ㑱���V���O���N�H�[�e�[�V���������݂��Ȃ���΁A�R�����g���擾
                lngRet = InStr(1, varBuf, SGL_QUOTATION & SGL_QUOTATION, vbTextCompare)
                If (pintChkCmt <> 1) Or (lngRet = 0) Then
                    '�R�����g���擾
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

            'Function �֐�������
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
                'Function �֐������擾
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
                        Debug.Print "-  �֐����@�@�F" & strFnNm
                        Debug.Print "-  �X�R�[�v�@�F" & strScope
                        Debug.Print "-  �t�@�C�����F" & StrGetFileName(vstrFileNm)
                        Debug.Print "-  �p�����[�^�F" & strParam
                        Debug.Print "-  �߂�l�@�@�F" & strReturn
                        Debug.Print "------------------------------------------------------------"
                    End If
                End If
            Else
            'Function �֐��̏I��������
                lngRet = InStr(1, varBuf, "End Function")
                If (lngRet = 1) Then
                    Debug.Print strCmt
                    ptypFunction(plngFuncCnt).Comment = strCmt
                    strCmt = vbNullString
                    blnExistFn = False
                Else
                    '�V���O���N�H�[�e�[�V������2�ȏ㑱���s��ǂݔ�΂��ݒ�ɂ���Ă��Ȃ��A��������2�ȏ㑱���V���O���N�H�[�e�[�V���������݂��Ȃ���΁A�R�����g���擾
                    lngRet = InStr(1, varBuf, SGL_QUOTATION & SGL_QUOTATION, vbTextCompare)
                    If (pintChkCmt <> 1) Or (lngRet = 0) Then
                        '�R�����g���擾
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
            'Sub �֐�������
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
                'Sub �֐������擾
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
                        Debug.Print "-  �֐����@�@�F" & strFnNm
                        Debug.Print "-  �X�R�[�v�@�F" & strScope
                        Debug.Print "-  �t�@�C�����F" & StrGetFileName(vstrFileNm)
                        Debug.Print "-  �p�����[�^�F" & strParam
                        Debug.Print "-  �߂�l�@�@�F" & strReturn
                        Debug.Print "------------------------------------------------------------"
                    End If
                End If
            Else
            'Sub �֐��̏I��������
                lngRet = InStr(1, varBuf, "End Sub")
                If (lngRet = 1) Then
                    Debug.Print strCmt
                    ptypFunction(plngFuncCnt).Comment = strCmt
                    strCmt = vbNullString
                    blnExistSub = False
                Else
                    '�V���O���N�H�[�e�[�V������2�ȏ㑱���s��ǂݔ�΂��ݒ�ɂ���Ă��Ȃ��A��������2�ȏ㑱���V���O���N�H�[�e�[�V���������݂��Ȃ���΁A�R�����g���擾
                    lngRet = InStr(1, varBuf, SGL_QUOTATION & SGL_QUOTATION, vbTextCompare)
                    If (pintChkCmt <> 1) Or (lngRet = 0) Then
                        '�R�����g���擾
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

'��O����
Exception:
    '�G���[���b�Z�[�W�\��
    Call pclsMsg.ShowError(Err.Description)
    On Error GoTo 0

End Function

Private Function IsExcelOut() As Boolean
'�\���̂̓��e���G�N�Z���e���v���[�g�ɓW�J

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
    Const SHEET_HYOUSHI             As String = "�\��": Const HYOUSHI_POS = 1
    Const SHEET_FILE                As String = "̧�ٍ\��": Const FILE_POS = 5
    Const SHEET_SYORI               As String = "�����T�v": Const SYORI_POS = 4
    Const SHEET_FUNCTION            As String = "�֐��ꗗ": Const FUNCTION_POS = 8
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

    '�e���v���[�g�̑��݃`�F�b�N
    strRet = Dir(strXltPath)
    If (strRet = vbNullString) Then
        Call pclsMsg.ShowMessage("XLT �t�@�C�������݂��܂���B")
        Exit Function
    End If

    '�i����ʂ̕\��
    Load frmProgress
    With frmProgress
        .Show
        .Refresh
    End With

    '�G�N�Z���I�u�W�F�N�g����
    Set objExcel = CreateObject("Excel.Application")

    Set objWorkBook = objExcel.Workbooks.Add(strXltPath)

    '�\��
    Set objWorksheet = objWorkBook.Worksheets(HYOUSHI_POS)
    objWorksheet.Name = SHEET_HYOUSHI
    objWorksheet.Activate
    With objWorksheet
        '�^�C�g��
        .Range("TITLE").Value = ptypVbFile.TITLE
        '�o�[�W����
        .Range("VERSION").Value = "Version " & ptypVbFile.MajorVer & _
                                         "." & ptypVbFile.MinorVer & _
                                         "." & ptypVbFile.RevisionVer
        '�J�����
        .Range("COMPANY").Value = ptypVbFile.VersionCompanyName
        '�v���W�F�N�g��
        .Range("NAME").Value = ptypVbFile.Name
        '���s�t�@�C����
        .Range("EXENAME").Value = ptypVbFile.ExeName32
        '�p�����[�^
        .Range("COMMAND").Value = StrChgNullStr(ptypVbFile.Command32, MSG_NOTHING)
        '�w���v�t�@�C����
        .Range("HELPFILE").Value = StrChgNullStr(ptypVbFile.HELPFILE, MSG_NOTHING)
    End With

    '̧�ٍ\��
    Set objWorksheet = objWorkBook.Worksheets(FILE_POS)
    objWorksheet.Activate
    objWorksheet.Name = SHEET_FILE
    With objWorksheet
        '�t�H�[��
        strChar = vbNullString
        For lngIdx = 0 To plngFrmCnt - 1
            If (strChar = vbNullString) Then
                strChar = ptypVbFile.Form(lngIdx).FormName & " (" & CStr(ptypVbFile.Form(lngIdx).RowNum) & " �s)"
            Else
                strChar = strChar & vbLf & ptypVbFile.Form(lngIdx).FormName & " (" & CStr(ptypVbFile.Form(lngIdx).RowNum) & " �s)"
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
        '���W���[��
        strChar = vbNullString
        For lngIdx = 0 To plngBasCnt - 1
            If (strChar = vbNullString) Then
                strChar = ptypVbFile.Module(lngIdx).ModuleName & " (" & CStr(ptypVbFile.Module(lngIdx).RowNum) & " �s)"
            Else
                strChar = strChar & vbLf & ptypVbFile.Module(lngIdx).ModuleName & " (" & CStr(ptypVbFile.Module(lngIdx).RowNum) & " �s)"
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
        '�N���X
        strChar = vbNullString
        For lngIdx = 0 To plngClsCnt - 1
            If (strChar = vbNullString) Then
                strChar = ptypVbFile.Class(lngIdx).ClassName & " (" & CStr(ptypVbFile.Class(lngIdx).RowNum) & " �s)"
            Else
                strChar = strChar & vbLf & ptypVbFile.Class(lngIdx).ClassName & " (" & CStr(ptypVbFile.Class(lngIdx).RowNum) & " �s)"
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

    '�����T�v
    Set objWorksheet = objWorkBook.Worksheets(SYORI_POS)
    objWorksheet.Activate
    objWorksheet.Name = SHEET_SYORI
    With objWorksheet
        '����
        .Range("DESCRIPTION").Value = ptypVbFile.VersionFileDescription
    End With

    '�֐��ꗗ
    Set objWorksheet = objWorkBook.Worksheets(FUNCTION_POS)
    objWorksheet.Activate
    objWorksheet.Name = SHEET_FUNCTION
    With objWorksheet
        strMsg = "�֐��ꗗ���o�͂��Ă��܂��c"
        '�֐������݂��Ȃ��Ă��A�֐��ꗗ�V�[�g�͍쐬����
        If (plngFuncCnt > -1) Then
            Call frmProgress.InitProgress(plngFuncCnt, strMsg)
            lngRowPos = ROW_START
            For lngIdx = 0 To plngFuncCnt
                .Range(COL_START & CStr(lngRowPos)).Value = ptypFunction(lngIdx).Name _
                                                    & "(" & ptypFunction(lngIdx).FileName & ")" & vbLf _
                                                    & "�X�R�[�v�@�F" & ptypFunction(lngIdx).Scope & vbLf _
                                                    & "�p�����[�^�F" & ptypFunction(lngIdx).Param & vbLf _
                                                    & "�߂�l�@�@�F" & ptypFunction(lngIdx).Return & vbLf
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

    '�G�N�Z���t�@�C���̕\��
    Set objWorksheet = objWorkBook.Sheets(1)
    objWorksheet.Activate
    objExcel.Visible = True

    '�G�N�Z���I�u�W�F�N�g�̉��
    Set objRange = Nothing
    Set objWorksheet = Nothing
    Set objWorkBook = Nothing
    Set objExcel = Nothing

    IsExcelOut = True

    Exit Function

'��O����
Exception:
    '�G���[���b�Z�[�W�\��
    Call pclsMsg.ShowError(Err.Description)
    Unload frmProgress
    '�G�N�Z���I�u�W�F�N�g�̉��
    Set objRange = Nothing
    Set objWorksheet = Nothing
    Set objWorkBook = Nothing
    Set objExcel = Nothing
    On Error GoTo 0

End Function

Private Function IsTextOut() As Boolean
'�\���̂̓��e���e�L�X�g�ɓW�J

    Dim strChar                     As String
    Dim lngIdx                      As Long
    Dim strMsg                      As String
    Dim blnRet                      As Boolean
    Dim strFileNm                   As String
    Dim strDefFileNm                As String

    Const ITEM_START                As String = "    �E"

    On Error GoTo Exception

    IsTextOut = False

    '�u�t�@�C����ۑ��v�_�C�A���O��\�����A�e�L�X�g�t�@�C���̕ۑ����I��������
    strDefFileNm = vbNullString
    blnRet = IsSaveTxtFileDlg(frmMain.cmnDlg, strFileNm, strDefFileNm)
    If (blnRet = False) Then
        IsTextOut = True
        Exit Function
    End If
    pclsLog.LogFileName = strFileNm
    pclsLog.Clear

    '�i����ʂ̕\��
    Load frmProgress
    With frmProgress
        .Show
        .Refresh
    End With

    Call pclsLog.LogMsg("������������������������������������������������������������")
    '�^�C�g�� & �o�[�W����
    Call pclsLog.LogMsg("��  " & ptypVbFile.TITLE _
                               & "Version " & ptypVbFile.MajorVer & _
                                        "." & ptypVbFile.MinorVer & _
                                        "." & ptypVbFile.RevisionVer)
    '����
    Call pclsLog.LogMsg("��" & vbNullString)
    Call pclsLog.LogMsg("��  " & ptypVbFile.VersionFileDescription)
    '�v���W�F�N�g��
    Call pclsLog.LogMsg("��" & vbNullString)
    Call pclsLog.LogMsg("��  �v���W�F�N�g���@�F" & ptypVbFile.Name)
    '���s�t�@�C����
    Call pclsLog.LogMsg("��  ���s�t�@�C�����@�F" & ptypVbFile.ExeName32)
    '�p�����[�^
    Call pclsLog.LogMsg("��  �p�����[�^�@�@�@�F" & StrChgNullStr(ptypVbFile.Command32, MSG_NOTHING))
    '�w���v�t�@�C����
    Call pclsLog.LogMsg("��  �w���v�t�@�C�����F" & StrChgNullStr(ptypVbFile.HELPFILE, MSG_NOTHING))
    Call pclsLog.LogMsg("������������������������������������������������������������")
    '�t�H�[��
    Call pclsLog.LogMsg(vbNullString)
    Call pclsLog.LogMsg("���g�p����t�H�[�����W���[��")
    strChar = vbNullString
    For lngIdx = 0 To plngFrmCnt - 1
        If (strChar = vbNullString) Then
            strChar = ITEM_START & ptypVbFile.Form(lngIdx).FormName & " (" & CStr(ptypVbFile.Form(lngIdx).RowNum) & " �s)"
        Else
            strChar = strChar & vbLf & ITEM_START & ptypVbFile.Form(lngIdx).FormName & " (" & CStr(ptypVbFile.Form(lngIdx).RowNum) & " �s)"
        End If
    Next lngIdx
    Call pclsLog.LogMsg(StrChgNullStr(strChar, MSG_NOTHING))
    '���W���[��
    Call pclsLog.LogMsg(vbNullString)
    Call pclsLog.LogMsg("���g�p����W�����W���[��")
    strChar = vbNullString
    For lngIdx = 0 To plngBasCnt - 1
        If (strChar = vbNullString) Then
            strChar = ITEM_START & ptypVbFile.Module(lngIdx).ModuleName & " (" & CStr(ptypVbFile.Module(lngIdx).RowNum) & " �s)"
        Else
            strChar = strChar & vbLf & ITEM_START & ptypVbFile.Module(lngIdx).ModuleName & " (" & CStr(ptypVbFile.Module(lngIdx).RowNum) & " �s)"
        End If
    Next lngIdx
    Call pclsLog.LogMsg(StrChgNullStr(strChar, MSG_NOTHING))
    '�N���X
    Call pclsLog.LogMsg(vbNullString)
    Call pclsLog.LogMsg("���g�p����N���X���W���[��")
    strChar = vbNullString
    For lngIdx = 0 To plngClsCnt - 1
        If (strChar = vbNullString) Then
            strChar = ITEM_START & ptypVbFile.Class(lngIdx).ClassName & " (" & CStr(ptypVbFile.Class(lngIdx).RowNum) & " �s)"
        Else
            strChar = strChar & vbLf & ITEM_START & ptypVbFile.Class(lngIdx).ClassName & " (" & CStr(ptypVbFile.Class(lngIdx).RowNum) & " �s)"
        End If
    Next lngIdx
    Call pclsLog.LogMsg(StrChgNullStr(strChar, MSG_NOTHING))
    '�֐��ꗗ
    Call pclsLog.LogMsg(vbNullString)
    If (plngFuncCnt > -1) Then
        strMsg = "�֐��ꗗ���o�͂��Ă��܂��c"
        Call frmProgress.InitProgress(plngFuncCnt, strMsg)
        For lngIdx = 0 To plngFuncCnt
            Call pclsLog.LogMsg("*****************************************")
            Call pclsLog.LogMsg("*" & vbNullString)
            Call pclsLog.LogMsg("*" & "    " & ptypFunction(lngIdx).Name & "(" & ptypFunction(lngIdx).FileName & ")")
            Call pclsLog.LogMsg("*" & "        �X�R�[�v�@�F" & ptypFunction(lngIdx).Scope)
            Call pclsLog.LogMsg("*" & "        �p�����[�^�F" & ptypFunction(lngIdx).Param)
            Call pclsLog.LogMsg("*" & "        �߂�l�@�@�F" & ptypFunction(lngIdx).Return)
            Call pclsLog.LogMsg("*" & vbNullString)
            Call pclsLog.LogMsg("*****************************************")
            Call pclsLog.LogMsg(ptypFunction(lngIdx).Comment)
            Call frmProgress.Progress(lngIdx)
        Next lngIdx
    End If

    Unload frmProgress

    IsTextOut = True

    Exit Function

'��O����
Exception:
    Call pclsMsg.ShowError(Err.Description)
    Unload frmProgress
    On Error GoTo 0

End Function

Private Function StrChgNullStr(ByVal vstrChk As String _
                             , ByVal vstrChg As String) As String
'����1�Ɏw�肵�������񂪋󔒂Ȃ�A����2�Ɏw�肵���������Ԃ�

    StrChgNullStr = vstrChk

    If (vstrChk = vbNullString) Then
        StrChgNullStr = vstrChg
    End If
End Function

Private Function IntCntLf(ByVal vstrChar As String) As Integer
'�����Ɋ܂܂����s������(vbLf)�̐����J�E���g���ĕԂ�

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

