Attribute VB_Name = "basVbSpecConst"
Option Explicit

'VBP ファイル
Public Const VBP_TITLE                  As String = "TITLE"                     'タイトル
Public Const VBP_NAME                   As String = "Name"                      'プロジェクト名
Public Const VBP_EXENAME32              As String = "ExeName32"                 '実行ファイル名
Public Const VBP_COMMAND32              As String = "Command32"                 'パラメータ
Public Const VBP_HELPFILE               As String = "HELPFILE"                  'ヘルプファイル名
Public Const VBP_VERSIONCOMMENTS        As String = "VersionComments"           'コメント
Public Const VBP_VERSIONFILEDESCRIPTION As String = "VersionFileDescription"    '説明
Public Const VBP_VERSIONCOMPANYNAME     As String = "VersionCompanyName"        '会社名
Public Const VBP_VERSIONLEGALCOPYRIGHT  As String = "VersionLegalCopyright"     '著作権
Public Const VBP_MAJORVER               As String = "MajorVer"                  'バージョン Major
Public Const VBP_MINORVER               As String = "MinorVer"                  'バージョン Minor
Public Const VBP_REVISIONVER            As String = "RevisionVer"               'バージョン Revision
Public Const VBP_OBJECT                 As String = "Object"                    'コンポーネント
Public Const VBP_FORM                   As String = "Form"                      'フォーム
Public Const VBP_MODULE                 As String = "Module"                    'モジュール
Public Const VBP_CLASS                  As String = "Class"                     'クラス

'ファイル操作
Public Const CURRENT_DIR                As String = "."
Public Const PEARENT_DIR                As String = ".."
Public Const DIR_SEPARATE               As String = "\"
Public Const DBL_QUOTATION              As String = """"

'記号
Public Const EQUALMARK                  As String = "="
Public Const SEMICOLON                  As String = ";"
Public Const SGL_QUOTATION              As String = "'"
Public Const COLON                      As String = ":"

'INIファイル文字列
Public Const INI_FILENM                 As String = "vbSpec.ini"                'INIファイル名
Public Const INI_SEC_CHK                As String = "CHKBOX"                    'セクション
Public Const INI_KEY_XLS                As String = "XLS"                       'エクセル出力
Public Const INI_KEY_TXT                As String = "TXT"                       'テキスト出力
Public Const INI_KEY_CMT                As String = "CMT"                       '設定キー1
Public Const INI_KEY_FUNC               As String = "FUNC"                      '設定キー2
Public Const INI_KEY_TRIM               As String = "TRIM"                      '設定キー3
Public Const INI_KEY_DIR                As String = "DIR"                       '設定キー4

'ファイル出力メッセージ
Public Const MSG_NOTHING                As String = "なし"
Public Const MSG_FUNCTION               As String = "関数："

