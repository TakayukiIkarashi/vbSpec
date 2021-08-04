VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "仕様書を作成しています･･･"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4905
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS UI Gothic"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox picProg 
      Height          =   375
      Left            =   180
      ScaleHeight     =   315
      ScaleWidth      =   4455
      TabIndex        =   1
      Top             =   600
      Width           =   4515
      Begin VB.Label lblProg 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.Label lblMsg 
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   4515
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngProgMax                 As Long

Private Sub Form_Load()
'フォームロード処理

    '画面の状態を初期化する
    lblMsg.Caption = vbNullString
End Sub

Public Sub InitProgress(ByVal vlngMax As Long _
                      , ByVal vstrMsg As String)
'進捗バーの初期化

    If (vlngMax > 0) Then
        lblProg.Width = 0
        mlngProgMax = vlngMax
    End If
    lblMsg.Caption = vstrMsg
    Me.Refresh
End Sub

Public Sub Progress(ByVal vlngIdx As Long)
'進捗バーの更新

    With lblProg
        .Width = picProg.Width * vlngIdx / mlngProgMax
        .Refresh
    End With
End Sub

