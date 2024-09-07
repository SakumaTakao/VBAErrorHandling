Attribute VB_Name = "記"
'Copyright 2022-2024 SAKUMA, Takao
Option Explicit

Public Enum 述
  引数①ニハEnum②ノ定義値シカ許サレマセンガ③ガ与エラレマシタ  ' = &H80000000
  引数①ノ値ヲ②ニ書キ換エテ続行シマス
  引数①ニNothingヲ与エルコトハ許サレマセン
  ①ニNothingヲ与エルコトハ許サレマセン 'Set Propertyの場合
  ①デえらーヲ生ジマシタ
  ①ニ与エタ引数ガ不正デス
  Class①ヲNewデ生成スルコトハ許サレマセン
  ①ヲ使ッテクダサイ
End Enum

Private Const CodeName As String = "記"

Private Const ① As Long = 9312 'ChrW("①")
Private Const 切 As String = vbCr
Private Const 述の連結 As String = _
    "引数「①」にはEnum「②」の定義値しか許されませんが ③ が与えられました。" & 切 & _
    "引数「①」の値を ② に書き換えて続行します。" & 切 & _
    "引数「①」にNothingを与えることは許されません。" & 切 & _
    "「①」にNothingを与えることは許されません。" & 切 & _
    "「①」でエラーを生じました。" & 切 & _
    "「①」に与えた引数が不正です。" & 切 & _
    "Class「①」をNewで生成することは許されません。" & 切 & _
    "「①」を使ってください。"

Private 述簿 As Variant '述簿() As String

Public Function 述ぶ(ByVal Enum述 As 述, ParamArray ①等の値() As Variant) As String
  Dim エラー As エラーの保存と復元
  Set エラー = New エラーの保存と復元
  
  On Error GoTo OnError
1:
  述ぶ = 述簿(Enum述)
2:
  Dim i As Long
  For i = LBound(①等の値) To UBound(①等の値)
    述ぶ = Replace(述ぶ, ChrW(① + i), ①等の値(i))
  Next
Exit Function
OnError:
  With Err()
    Select Case .Number
      Case ErrNumber.型違い, ErrNumber.Null不正, ErrNumber.インスタンス未設定
        If Erl() = 1 Then
          述簿 = Split(述の連結, 切)
          Resume
        End If
        録す "記庫.記.述ぶ", "引数「①等の値」の" & CStr(i + 1&) & "番目の要素がString型に変換できません。"
        Resume Next
      Case ErrNumber.番外
        録す "記庫.記.述ぶ", "引数「Enum述」に該当する文字列が登録されていません。"
        述ぶ = "（Function「述ぶ」がエラーになりました。"
      Case Else
         録す "記庫.記.述ぶ"
        If 述ぶ = "" Then 述ぶ = "（Function「述ぶ」がエラーになりました。）"
    End Select
  End With
End Function

