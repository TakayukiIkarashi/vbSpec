VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  '固定(実線)
   ClientHeight    =   4515
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "MS UI Gothic"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3810.127
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   6315
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox txtVbp 
      Height          =   285
      Left            =   300
      OLEDropMode     =   1  '手動
      TabIndex        =   1
      Top             =   570
      Width           =   5385
   End
   Begin VB.CommandButton cmdVbp 
      Caption         =   "..."
      Height          =   255
      Left            =   5700
      TabIndex        =   2
      Top             =   600
      Width           =   315
   End
   Begin VB.CheckBox chkXls 
      Caption         =   "Excel ファイルに出力"
      Height          =   195
      Left            =   540
      TabIndex        =   3
      Top             =   1080
      Width           =   2355
   End
   Begin VB.CheckBox chkTxt 
      Caption         =   "テキストファイルに出力"
      Height          =   195
      Left            =   540
      TabIndex        =   4
      Top             =   1380
      Width           =   2355
   End
   Begin VB.Frame fraOption 
      Caption         =   "詳細設定"
      Height          =   1875
      Left            =   300
      TabIndex        =   5
      Top             =   1770
      Width           =   5715
      Begin VB.CheckBox chkTrim 
         Caption         =   "コメント行出力の開始位置は、ソースコードと同じにする"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   5175
      End
      Begin VB.CheckBox chkDir 
         Caption         =   $"frmMain.frx":0CCA
         Height          =   555
         Left            =   240
         TabIndex        =   9
         Top             =   1170
         Width           =   5295
      End
      Begin VB.CheckBox chkFunc 
         Caption         =   "関数上部のコメントは、その下の関数を説明している"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   660
         Width           =   5175
      End
      Begin VB.CheckBox chkCmt 
         Caption         =   "シングルクォーテーション(')が２つ以上続く行を読み飛ばす"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame fraBtn 
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   375
      Left            =   3420
      TabIndex        =   10
      Top             =   4080
      Width           =   2775
      Begin VB.CommandButton cmdCancel 
         Caption         =   "キャンセル"
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Top             =   30
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   315
         Left            =   0
         TabIndex        =   11
         Top             =   30
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   5520
      Top             =   1230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblVbp 
      Caption         =   "Visual Basic Project(&V):"
      Height          =   225
      Left            =   330
      TabIndex        =   0
      Top             =   330
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnEventExit               As Boolean

Private Sub Form_Load()
'フォームロード時処理

    '見出しを設定する
    Me.Caption = App.TITLE & " Ver." & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)

    'チェックボックス状態を変数から取得する
    chkXls.Value = pintChkXls
    chkTxt.Value = pintChkTxt
    chkCmt.Value = pintChkCmt
    chkDir.Value = pintChkDir
    chkFunc.Value = pintChkFunc
    chkTrim.Value = pintChkTrim

    '画面の状態を初期化する
    txtVbp.Text = vbNullString
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'フォームクエリーアンロード時処理

    'INIファイルにデータを保存
    Call SetIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_XLS, CStr(pintChkXls))
    Call SetIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_TXT, CStr(pintChkTxt))
    Call SetIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_CMT, CStr(pintChkCmt))
    Call SetIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_DIR, CStr(pintChkDir))
    Call SetIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_FUNC, CStr(pintChkFunc))
    Call SetIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_TRIM, CStr(pintChkTrim))
End Sub

Private Sub cmdOK_Click()
'OKボタンクリック時処理

    Dim blnRet                      As Boolean
    Dim strVbpPath                  As String

    '仕様書の出力先ファイルが指定されていなければ、処理しない
    If (chkXls.Value = 0) And (chkTxt.Value = 0) Then
        Call pclsMsg.ShowMessage("出力先ファイルが指定されていません。")
        Exit Sub
    End If

    'マウスポインタを砂時計にする
    Screen.MousePointer = vbHourglass

    strVbpPath = txtVbp.Text

    'IsExec関数を実行する
    blnRet = basMain.IsExec(strVbpPath)
    If (blnRet = False) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    'マウスポインタを元に戻す
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancel_Click()
'キャンセルボタンクリック時処理

    Unload Me
End Sub

Private Sub txtVbp_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'vbpファイル入力テキストボックス OLEDragDrop

    On Error Resume Next

    txtVbp.Text = Data.Files(1)

    On Error GoTo 0
End Sub

Private Sub cmdVbp_Click()
'vbpファイル参照ボタンクリック時処理

    Dim strFileNm                   As String
    Dim strDefFileNm                As String
    Dim blnRet                      As Boolean

    strDefFileNm = vbNullString

    '「ファイルを開く」ダイアログを表示する
    blnRet = IsOpenVbpFileDlg(cmnDlg, strFileNm, strDefFileNm)
    If (blnRet = False) Then
        Exit Sub
    End If

    txtVbp.Text = strFileNm
End Sub

Private Sub chkXls_Click()
'エクセル出力チェックボックスクリック時処理

    '変数にチェック状態を格納
    pintChkXls = chkXls.Value
End Sub

Private Sub chkTxt_Click()
'テキスト出力チェックボックスクリック時処理

    '変数にチェック状態を格納
    pintChkTxt = chkTxt.Value
End Sub

Private Sub chkCmt_Click()
'「シングルクォーテーション(')が２つ以上続く行を読み飛ばす」
'チェックボックスクリック時処理

    '変数にチェック状態を格納
    pintChkCmt = chkCmt.Value
End Sub

Private Sub chkDir_Click()
'「関数上部のコメントは、その下の関数を説明している」
'チェックボックスクリック時処理

    '変数にチェック状態を格納
    pintChkDir = chkDir.Value
End Sub

Private Sub chkFunc_Click()
'「コメント行の開始位置は、プログラムソースと同じにする」
'チェックボックスクリック時処理

    '変数にチェック状態を格納
    pintChkFunc = chkFunc.Value
End Sub

Private Sub chkTrim_Click()
'「指定したvbpファイルより上位のディレクトリに存在するファイルは、共通ファイルであるとみなし、読み飛ばす」
'チェックボックスクリック時処理

    '変数にチェック状態を格納
    pintChkTrim = chkTrim.Value
End Sub
