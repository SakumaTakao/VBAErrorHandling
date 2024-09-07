VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "録"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Copyright 2022-2024 SAKUMA, Takao
Option Explicit

Public Sub 録す(事 As 事)
  With 事
    Debug.Print .When, .Source, .Description
  End With
End Sub

'Public Sub 録す(ByVal 所※プロシージャ名など As String, Optional ByVal 述 As String)
'  Dim ラベル As String
'  With Err()
'    If 所※プロシージャ名など <> "" Then
'      .Source = 所※プロシージャ名など
'      If Erl() And .Number Then ラベル = " " & "ラベル" & CStr(Erl())
'    End If
'    If 述 = "" Then 述 = .Description
'    If .Number Then 述 = CStr(.Number) & ": " & 述
'    'エラーが発生していないときはラベルを表示しないようにする必要がある
'    Debug.Print Now, .Source & ラベル, 述
'  End With
'End Sub
