VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "エラーの保存と復元"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Copyright 2022-2024 SAKUMA, Takao
Option Explicit

Private Const CodeName As String = "エラーの保存と復元"
Private 事 As 事

Private Sub Class_Initialize()
  事 = 主.エラー()
End Sub

Public Property Get エラー() As 事
Attribute エラー.VB_UserMemId = 0
  'Attribute ErrCopy.VB_UserMemId = 0
  エラー = 事
End Property

Private Sub Class_Terminate()
  主.エラー() = 事
End Sub
