Attribute VB_Name = "記"
'Copyright 2022-2024 SAKUMA, Takao
Option Explicit

Public Enum 述
  引数�@ニハEnum�Aノ定義値シカ許サレマセンガ�Bガ与エラレマシタ  ' = &H80000000
  引数�@ノ値ヲ�Aニ書キ換エテ続行シマス
  引数�@ニNothingヲ与エルコトハ許サレマセン
  �@ニNothingヲ与エルコトハ許サレマセン 'Set Propertyの場合
  �@デえらーヲ生ジマシタ
  �@ニ与エタ引数ガ不正デス
  Class�@ヲNewデ生成スルコトハ許サレマセン
  �@ヲ使ッテクダサイ
End Enum

Private Const CodeName As String = "記"

Private Const �@ As Long = 9312 'ChrW("�@")
Private Const 切 As String = vbCr
Private Const 述の連結 As String = _
    "引数「�@」にはEnum「�A」の定義値しか許されませんが �B が与えられました。" & 切 & _
    "引数「�@」の値を �A に書き換えて続行します。" & 切 & _
    "引数「�@」にNothingを与えることは許されません。" & 切 & _
    "「�@」にNothingを与えることは許されません。" & 切 & _
    "「�@」でエラーを生じました。" & 切 & _
    "「�@」に与えた引数が不正です。" & 切 & _
    "Class「�@」をNewで生成することは許されません。" & 切 & _
    "「�@」を使ってください。"

Private 述簿 As Variant '述簿() As String

Public Function 述ぶ(ByVal Enum述 As 述, ParamArray �@等の値() As Variant) As String
  Dim エラー As エラーの保存と復元
  Set エラー = New エラーの保存と復元
  
  On Error GoTo OnError
1:
  述ぶ = 述簿(Enum述)
2:
  Dim i As Long
  For i = LBound(�@等の値) To UBound(�@等の値)
    述ぶ = Replace(述ぶ, ChrW(�@ + i), �@等の値(i))
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
        録す "記庫.記.述ぶ", "引数「�@等の値」の" & CStr(i + 1&) & "番目の要素がString型に変換できません。"
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

