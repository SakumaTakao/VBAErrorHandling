VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "範"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_Initialize()
  Debug.Print "Class「範」は抽象クラスでありインスタンスを生成できません。"
  Err.Raise 1
End Sub

Public Property Get 見出Address() As String

End Property

Public Property Get 最終行判定列番() As Long

End Property

'Public Property Get フィルタ() As Excel.AutoFilter
'
'End Property

Public Function 出() As 出

End Function
