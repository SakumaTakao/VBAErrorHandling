VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "主"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Enum 上書
  サブクラスの処理のみ行う = 65536
  サブクラスの処理後スーパークラスの処理を行う
End Enum

Public Function 受け渡す(Optional ByRef 物 As Variant) As Variant
  On Error GoTo OnError
  Static 在庫 As Variant
  If IsEmpty(在庫) Then
    If IsObject(物) Then Set 在庫 = 物 Else 在庫 = 物
  Else
    If IsObject(在庫) Then Set 受け渡す = 在庫 Else 受け渡す = 在庫
    在庫 = Empty
  End If
Exit Function
OnError:
End Function

Public Sub フォルダ内の全ファイルを処理する(ルートパス As String, 訪 As 訪)
  On Error GoTo OnError
  受け渡す 訪
  With New 訪
    .フォルダ内の全ファイルを処理する ルートパス
  End With
Exit Sub
OnError:
End Sub
