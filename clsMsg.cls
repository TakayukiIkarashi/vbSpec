VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowError(ByVal vstrMessage As String)
'エラーメッセージを表示
    Call MsgBox(vstrMessage, vbOKOnly + vbCritical, "エラー")
End Sub

Public Sub ShowMessage(ByVal vstrMessage As String)
'注意メッセージを表示
    Call MsgBox(vstrMessage, vbOKOnly + vbExclamation, "注意")
End Sub

Public Sub ShowInformation(ByVal vstrInformation As String)
'情報メッセージを表示
    Call MsgBox(vstrInformation, vbOKOnly + vbInformation, "情報")
End Sub

Public Function ShowQuestion(ByVal vstrQuestion As String) As Integer
'質問メッセージを表示
    ShowQuestion = MsgBox(vstrQuestion, vbYesNo + vbQuestion + vbDefaultButton2, "質問")
End Function

