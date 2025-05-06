VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Implements 訪
Implements 範
Implements 出

Private Enum 行
  見出 = 2
  始 = 3
  Excel最大 = 1048576
End Enum

Private Enum 列
  ライブラリ = 1
  モジュール
  メンバー
  属性
  種類
  型
End Enum

Private 書込行番 As Long

Public Sub テスト()
  Debug.Print "TypeOf Me Is Worksheet", TypeOf Me Is Worksheet
  Debug.Print "TypeOf Me Is 範", TypeOf Me Is 範
  Debug.Print "TypeOf Me Is Workbook", TypeOf Me Is Workbook
  Debug.Print "Not TypeOf Me Is Worksheet", Not TypeOf Me Is Worksheet
  Debug.Print "Not TypeOf Me IsNot 範", Not TypeOf Me Is 範
  Debug.Print "Not TypeOf Me IsNot Workbook", Not TypeOf Me Is Workbook
  
End Sub

Private Property Get 範_見出Address() As String
  範_見出Address = "$B$2:$E$2"
End Property

Private Property Get 範_最終行判定列番() As Long
  範_最終行判定列番 = 列.モジュール
End Property

Private Function 範_出() As 巡庫.出
  On Error GoTo OnError
  Copy
  With Parent.Workbooks
    Set 範_出 = .Item(.Count)
  End With
Exit Function
OnError:

End Function

Private Sub 訪_処理する(ファイル As IWshRuntimeLibrary.IFile)
  On Error GoTo OnError
  If InStr(ファイル.Type, "Excel") = 0 Then Err.Raise 上書.サブクラスの処理のみ行う
  '結果として何も処理せずに抜ける
Exit Sub
OnError:
  Err.Raise 上書.サブクラスの処理のみ行う
End Sub

Private Function 訪_対象外(フォルダ As IWshRuntimeLibrary.IFolder) As Boolean

End Function

Private Sub 出_Initialize()
  AutoFilter = False
  Range(Cells(行.始, 1), Cells(行.Excel最大, 1)).EntireRow.ClearContents
  書込行番 = 行.始
End Sub

Private Sub 出_Terminate()
  Range(範_見出Address).Resize(最終行 - 行.見出 + 1&).AutoFilter
End Sub

Private Sub 出_転記する(データ As Variant)
  On Error GoTo OnError
  Cells(書込行番, 1&).Resize(UBound(データ, 2), UBound(データ, 1)).Value = データ
  書込行番 = 最終行 + 1&
Exit Sub
OnError:

End Sub

Private Property Get 最終行() As Long
  On Error GoTo OnError
  最終行 = Cells(行.Excel最大, 範_最終行判定列番()).End(xlUp).Row()
  If 最終行 < 行.見出 Then Err().Raise 1
Exit Property
OnError:
End Property
  
