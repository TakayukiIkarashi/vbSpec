Attribute VB_Name = "basVbSpecStructure"
Option Explicit

'vbp �t�@�C�� �\���̗p From
Public Type TFORM
    FormName                        As String           '�t�H�[����
    FormPath                        As String           '�t�H�[���p�X
    RowNum                          As Long             '�s��
End Type

'vbp �t�@�C�� �\���̗p Module
Public Type TMODULE
    ModuleName                      As String           '���W���[����
    ModulePath                      As String           '���W���[���p�X
    RowNum                          As Long             '�s��
End Type

'vbp �t�@�C�� �\���̗p Class
Public Type TCLASS
    ClassName                       As String           '�N���X��
    ClassPath                       As String           '�N���X�p�X
    RowNum                          As Long             '�s��
End Type

'vbp �t�@�C��
Public Type TVBPFILE
    TITLE                           As String           '�^�C�g��
    Name                            As String           '�v���W�F�N�g��
    ExeName32                       As String           '���s�t�@�C����
    Command32                       As String           '�p�����[�^
    HELPFILE                        As String           '�w���v�t�@�C����
    VersionComments                 As String           '�R�����g
    VersionFileDescription          As String           '����
    VersionCompanyName              As String           '��Ж�
    VersionLegalCopyright           As String           '���쌠
    MajorVer                        As String           '�o�[�W���� Major
    MinorVer                        As String           '�o�[�W���� Minor
    RevisionVer                     As String           '�o�[�W���� Revision
    Object                          As String           '�R���|�[�l���g
    Form()                          As TFORM            '�t�H�[��
    Module()                        As TMODULE          '���W���[��
    Class()                         As TCLASS           '�N���X
End Type
Public ptypVbFile                   As TVBPFILE         'vbp �t�@�C�� �\����

Public plngObjCnt                   As Long             '�g�p�R���|�[�l���g��
Public plngFrmCnt                   As Long             '�g�p�t�H�[����
Public plngBasCnt                   As Long             '�g�p�W�����W���[����
Public plngClsCnt                   As Long             '�g�p�N���X���W���[����


'VB �֐�
Public Type TFUNCTION
    Name                            As String           '�֐���
    Scope                           As String           '�֐��̃X�R�[�v
    FileName                        As String           '�֐�����`����Ă���t�@�C����
    Comment                         As String           '�֐����̃R�����g
    Param                           As String           '�֐��̃p�����[�^
    Return                          As String           '�֐��̖߂�l
End Type
Public ptypFunction()               As TFUNCTION        'VB �֐� �\����

Public plngFuncCnt                  As Long             '�֐�����

