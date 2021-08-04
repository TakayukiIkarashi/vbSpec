VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  '�Œ�(����)
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
   ScaleMode       =   0  'հ�ް
   ScaleWidth      =   6315
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox txtVbp 
      Height          =   285
      Left            =   300
      OLEDropMode     =   1  '�蓮
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
      Caption         =   "Excel �t�@�C���ɏo��"
      Height          =   195
      Left            =   540
      TabIndex        =   3
      Top             =   1080
      Width           =   2355
   End
   Begin VB.CheckBox chkTxt 
      Caption         =   "�e�L�X�g�t�@�C���ɏo��"
      Height          =   195
      Left            =   540
      TabIndex        =   4
      Top             =   1380
      Width           =   2355
   End
   Begin VB.Frame fraOption 
      Caption         =   "�ڍאݒ�"
      Height          =   1875
      Left            =   300
      TabIndex        =   5
      Top             =   1770
      Width           =   5715
      Begin VB.CheckBox chkTrim 
         Caption         =   "�R�����g�s�o�͂̊J�n�ʒu�́A�\�[�X�R�[�h�Ɠ����ɂ���"
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
         Caption         =   "�֐��㕔�̃R�����g�́A���̉��̊֐���������Ă���"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   660
         Width           =   5175
      End
      Begin VB.CheckBox chkCmt 
         Caption         =   "�V���O���N�H�[�e�[�V����(')���Q�ȏ㑱���s��ǂݔ�΂�"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame fraBtn 
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "Frame1"
      Height          =   375
      Left            =   3420
      TabIndex        =   10
      Top             =   4080
      Width           =   2775
      Begin VB.CommandButton cmdCancel 
         Caption         =   "�L�����Z��"
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
'�t�H�[�����[�h������

    '���o����ݒ肷��
    Me.Caption = App.TITLE & " Ver." & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)

    '�`�F�b�N�{�b�N�X��Ԃ�ϐ�����擾����
    chkXls.Value = pintChkXls
    chkTxt.Value = pintChkTxt
    chkCmt.Value = pintChkCmt
    chkDir.Value = pintChkDir
    chkFunc.Value = pintChkFunc
    chkTrim.Value = pintChkTrim

    '��ʂ̏�Ԃ�����������
    txtVbp.Text = vbNullString
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'�t�H�[���N�G���[�A�����[�h������

    'INI�t�@�C���Ƀf�[�^��ۑ�
    Call SetIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_XLS, CStr(pintChkXls))
    Call SetIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_TXT, CStr(pintChkTxt))
    Call SetIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_CMT, CStr(pintChkCmt))
    Call SetIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_DIR, CStr(pintChkDir))
    Call SetIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_FUNC, CStr(pintChkFunc))
    Call SetIniData(pstrIniPath, INI_SEC_CHK, INI_KEY_TRIM, CStr(pintChkTrim))
End Sub

Private Sub cmdOK_Click()
'OK�{�^���N���b�N������

    Dim blnRet                      As Boolean
    Dim strVbpPath                  As String

    '�d�l���̏o�͐�t�@�C�����w�肳��Ă��Ȃ���΁A�������Ȃ�
    If (chkXls.Value = 0) And (chkTxt.Value = 0) Then
        Call pclsMsg.ShowMessage("�o�͐�t�@�C�����w�肳��Ă��܂���B")
        Exit Sub
    End If

    '�}�E�X�|�C���^�������v�ɂ���
    Screen.MousePointer = vbHourglass

    strVbpPath = txtVbp.Text

    'IsExec�֐������s����
    blnRet = basMain.IsExec(strVbpPath)
    If (blnRet = False) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '�}�E�X�|�C���^�����ɖ߂�
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancel_Click()
'�L�����Z���{�^���N���b�N������

    Unload Me
End Sub

Private Sub txtVbp_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'vbp�t�@�C�����̓e�L�X�g�{�b�N�X OLEDragDrop

    On Error Resume Next

    txtVbp.Text = Data.Files(1)

    On Error GoTo 0
End Sub

Private Sub cmdVbp_Click()
'vbp�t�@�C���Q�ƃ{�^���N���b�N������

    Dim strFileNm                   As String
    Dim strDefFileNm                As String
    Dim blnRet                      As Boolean

    strDefFileNm = vbNullString

    '�u�t�@�C�����J���v�_�C�A���O��\������
    blnRet = IsOpenVbpFileDlg(cmnDlg, strFileNm, strDefFileNm)
    If (blnRet = False) Then
        Exit Sub
    End If

    txtVbp.Text = strFileNm
End Sub

Private Sub chkXls_Click()
'�G�N�Z���o�̓`�F�b�N�{�b�N�X�N���b�N������

    '�ϐ��Ƀ`�F�b�N��Ԃ��i�[
    pintChkXls = chkXls.Value
End Sub

Private Sub chkTxt_Click()
'�e�L�X�g�o�̓`�F�b�N�{�b�N�X�N���b�N������

    '�ϐ��Ƀ`�F�b�N��Ԃ��i�[
    pintChkTxt = chkTxt.Value
End Sub

Private Sub chkCmt_Click()
'�u�V���O���N�H�[�e�[�V����(')���Q�ȏ㑱���s��ǂݔ�΂��v
'�`�F�b�N�{�b�N�X�N���b�N������

    '�ϐ��Ƀ`�F�b�N��Ԃ��i�[
    pintChkCmt = chkCmt.Value
End Sub

Private Sub chkDir_Click()
'�u�֐��㕔�̃R�����g�́A���̉��̊֐���������Ă���v
'�`�F�b�N�{�b�N�X�N���b�N������

    '�ϐ��Ƀ`�F�b�N��Ԃ��i�[
    pintChkDir = chkDir.Value
End Sub

Private Sub chkFunc_Click()
'�u�R�����g�s�̊J�n�ʒu�́A�v���O�����\�[�X�Ɠ����ɂ���v
'�`�F�b�N�{�b�N�X�N���b�N������

    '�ϐ��Ƀ`�F�b�N��Ԃ��i�[
    pintChkFunc = chkFunc.Value
End Sub

Private Sub chkTrim_Click()
'�u�w�肵��vbp�t�@�C������ʂ̃f�B���N�g���ɑ��݂���t�@�C���́A���ʃt�@�C���ł���Ƃ݂Ȃ��A�ǂݔ�΂��v
'�`�F�b�N�{�b�N�X�N���b�N������

    '�ϐ��Ƀ`�F�b�N��Ԃ��i�[
    pintChkTrim = chkTrim.Value
End Sub
