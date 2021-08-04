Attribute VB_Name = "basVbSpecStructure"
Option Explicit

'vbp ファイル 構造体用 From
Public Type TFORM
    FormName                        As String           'フォーム名
    FormPath                        As String           'フォームパス
    RowNum                          As Long             '行数
End Type

'vbp ファイル 構造体用 Module
Public Type TMODULE
    ModuleName                      As String           'モジュール名
    ModulePath                      As String           'モジュールパス
    RowNum                          As Long             '行数
End Type

'vbp ファイル 構造体用 Class
Public Type TCLASS
    ClassName                       As String           'クラス名
    ClassPath                       As String           'クラスパス
    RowNum                          As Long             '行数
End Type

'vbp ファイル
Public Type TVBPFILE
    TITLE                           As String           'タイトル
    Name                            As String           'プロジェクト名
    ExeName32                       As String           '実行ファイル名
    Command32                       As String           'パラメータ
    HELPFILE                        As String           'ヘルプファイル名
    VersionComments                 As String           'コメント
    VersionFileDescription          As String           '説明
    VersionCompanyName              As String           '会社名
    VersionLegalCopyright           As String           '著作権
    MajorVer                        As String           'バージョン Major
    MinorVer                        As String           'バージョン Minor
    RevisionVer                     As String           'バージョン Revision
    Object                          As String           'コンポーネント
    Form()                          As TFORM            'フォーム
    Module()                        As TMODULE          'モジュール
    Class()                         As TCLASS           'クラス
End Type
Public ptypVbFile                   As TVBPFILE         'vbp ファイル 構造体

Public plngObjCnt                   As Long             '使用コンポーネント数
Public plngFrmCnt                   As Long             '使用フォーム数
Public plngBasCnt                   As Long             '使用標準モジュール数
Public plngClsCnt                   As Long             '使用クラスモジュール数


'VB 関数
Public Type TFUNCTION
    Name                            As String           '関数名
    Scope                           As String           '関数のスコープ
    FileName                        As String           '関数が定義されているファイル名
    Comment                         As String           '関数内のコメント
    Param                           As String           '関数のパラメータ
    Return                          As String           '関数の戻り値
End Type
Public ptypFunction()               As TFUNCTION        'VB 関数 構造体

Public plngFuncCnt                  As Long             '関数件数

