VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mstrLogFileName             As String

'ログファイル名の設定
Public Property Let LogFileName(ByVal vstrData As String)
Attribute LogFileName.VB_Description = "出力ログファイル名を設定する。参照も可能。"
    mstrLogFileName = vstrData
End Property
Public Property Get LogFileName() As String
    LogFileName = mstrLogFileName
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "ログファイルのクリアを行う。ファイルの削除は行わない。"
'ログファイルの初期化

    Dim intFileNo                   As Integer

    If (RTrim(mstrLogFileName) = vbNullString) Then
        Exit Sub
    End If

    On Error Resume Next

    intFileNo = FreeFile

    Open mstrLogFileName For Output As #intFileNo

    Close #intFileNo

    On Error GoTo 0
End Sub

Public Sub LogMsg(ByVal vstrMsg As String)
Attribute LogMsg.VB_Description = "LogFileNameﾌﾟﾛﾊﾟﾃｨに設定されているファイルにメッセージを出力する。"
'ログファイルにメッセージを追加する

    Dim intFileNo                   As Integer

    On Error Resume Next

    If (RTrim(mstrLogFileName) = "") Then
        Exit Sub
    End If

    intFileNo = FreeFile

    Open mstrLogFileName For Append As intFileNo

    Print #intFileNo, vstrMsg

    Close intFileNo

    On Error GoTo 0
End Sub

Private Sub Class_Initialize()
'ログ出力クラスの初期化処理
    mstrLogFileName = vbNullString
End Sub

