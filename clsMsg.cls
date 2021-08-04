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
'�G���[���b�Z�[�W��\��
    Call MsgBox(vstrMessage, vbOKOnly + vbCritical, "�G���[")
End Sub

Public Sub ShowMessage(ByVal vstrMessage As String)
'���Ӄ��b�Z�[�W��\��
    Call MsgBox(vstrMessage, vbOKOnly + vbExclamation, "����")
End Sub

Public Sub ShowInformation(ByVal vstrInformation As String)
'��񃁃b�Z�[�W��\��
    Call MsgBox(vstrInformation, vbOKOnly + vbInformation, "���")
End Sub

Public Function ShowQuestion(ByVal vstrQuestion As String) As Integer
'���⃁�b�Z�[�W��\��
    ShowQuestion = MsgBox(vstrQuestion, vbYesNo + vbQuestion + vbDefaultButton2, "����")
End Function
