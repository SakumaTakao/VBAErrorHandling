Attribute VB_Name = "主"
'Copyright 2022-2024 SAKUMA, Takao
Option Explicit

Public Type 事
  ID As Long
  Source As String
  Description As String
  HelpFile As String
  HelpContext As Long
  LastDllError As Long
  
  Erl As Long '行ラベルの設定可能値は0～2147483647（Long型のうち負でない値）
  When As Date
End Type

Public Enum ErrNumber
  索引該当無 = 5& 'CollectionやDictionaryの文字列キーに該当無
  集が空 = 5&
  パスが空文字 = 5&
  オーバーフロー = 6&
  番外 = 9& '配列、Collection、Dictionaryのインデクスが範囲外
  ReDim未済 = 9&
  型違い = 13&
  ファイル名不正 = 52&
  ファイル未作成 = 53&
  ファイル既開 = 55&
  ファイル読取専用 = 70&
  ファイルパス無効 = 75&
  ファイルパス不明 = 76&
  未生成 = 91&
  インスタンス未設定 = 91&
  Null不正 = 94&
  インスタンス要 = 424&
  数やStringにインスタンスを代入 = 438&
  メソッド無し = 438&
  属性欠如 = 438&
  生成失敗 = 440&
  索引登録済 = 457&
  
  上限 = 65535
End Enum

Public Enum ErrID※65536以上
  呼出先エラー = ErrNumber.上限 + 1&
  用法相違
  引数相違
End Enum

Private Const CodeName As String = "主"
Private Const LastDllError連結句 As String = " LastDllError = "

Private ErrCopy As 事
Private 事の初期値 As 事
Private 病むで発生させたエラー As Boolean

Private 録インスタンス As New 録

Public Sub 録す(ByVal 所※プロシージャ名など As String, Optional ByVal 述 As String)
  Dim エラー As エラーの保存と復元
  Set エラー = New エラーの保存と復元
  
  On Error GoTo OnError
  Dim 事 As 事
  With エラー.エラー
    事.Source = IIf(所※プロシージャ名など = "", .Source, 所※プロシージャ名など)
    事.Description = IIf(述 = "", .Description, 述)

    If .ID Then
      事.Description = CStr(.ID) & ": " & 事.Description
      If .Erl Then 事.Source = 事.Source & " ラベル" & CStr(Erl())
    End If
  End With
  事.When = Now()

  録インスタンス.録す 事
Exit Sub
OnError:
  録す CodeName & ".録す"
  Resume Next
End Sub

Public Sub 病む(ByVal ErrID As ErrID※65536以上, Optional ByVal 所※プロシージャ名など As String, Optional ByVal 述 As String)
  If ErrID <= ErrNumber.上限 Then
    録す 所※プロシージャ名など, "「病む」の引数「ErrID」には65536以上の整数しか許されませんが " & ErrID & " が与えられました。"
    ErrID = -ErrID
    If ErrID = 0 Then ErrID = ErrID※65536以上.引数相違
  End If
  
  Dim 事 As 事
  With 事
    .ID = -ErrID
    .Source = 所※プロシージャ名など
    .Description = 述
    
    'Err().Raiseの際に引数HelpFileを省略するとデフォルト値が設定されてしまう
    'それを防ぐため空文字を設定する
    .HelpFile = ""
    
    '引数HelpContextに0を与えるとErr().Raiseの際に1000440に変換されてしまうので、-1を与える
    .HelpContext = -1
  End With
  エラー() = 事
  病むで発生させたエラー = True
  
  '本プロシージャのOptional引数が省略されていた場合は、Err().Raiseの際に対応する引数を省略しデフォルト値を設定させる
  With Err() '直前の「エラー() = 事」により更新されている
    If 所※プロシージャ名など = "" Then
      If 述 = "" Then
        .Raise .Number, , , .HelpFile, .HelpContext
      Else
        .Raise .Number, , .Description, .HelpFile, .HelpContext
      End If
    Else
      If 述 = "" Then
        .Raise .Number, .Source, , .HelpFile, .HelpContext
      Else
        .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
      End If
    End If
  End With
End Sub

Public Sub 録し病む※OnError(ByVal 所※プロシージャ名など As String)
  With エラー()
    If .ID Then
      録す 所※プロシージャ名など
      If .ID < -ErrNumber.上限 And .HelpFile = "" Then 'これがTrueになることは無い？
        病む -.ID
      Else
        病む 呼出先エラー, 所※プロシージャ名など, "「" & 所※プロシージャ名など & "」でエラーを生じました。"
      End If
    Else
      録す 所※プロシージャ名など, "エラーが生じていないのに「録し病む※OnError」を呼ぶことは許されません。"
    End If
  End With
End Sub

Public Property Set 録(ByVal 録 As 録)
  On Error GoTo OnError
  If 録 Is Nothing Then Set 録 = New 録
  Set 録インスタンス = 録
Exit Property
OnError:
  録し病む※OnError "Set 録"
End Property

Public Property Get エラー() As 事
'  If (ErrCopy.ID <> 0) And (Err().Number <> 0) Then
  If CBool(ErrCopy.ID) * Err().Number Then
    With ErrCopy
      エラー.ID = .ID
      エラー.Source = .Source
      エラー.Description = .Description
      エラー.HelpFile = .HelpFile
      エラー.HelpContext = .HelpContext
      
      If 病むで発生させたエラー Then
        エラー.LastDllError = Err().LastDllError
        エラー.Erl = Erl()
        病むで発生させたエラー = False
      Else
        エラー.LastDllError = .LastDllError
        エラー.Erl = .Erl
      End If
    End With
  Else
    With Err()
      エラー.ID = .Number
      エラー.Source = .Source
      エラー.Description = .Description
      エラー.HelpFile = .HelpFile
      エラー.HelpContext = .HelpContext
      エラー.LastDllError = .LastDllError
      エラー.Erl = Erl()
    End With
  End If
  ErrCopy = 事の初期値
End Property

Public Property Let エラー(事 As 事)
  With ErrCopy
    .ID = 事.ID
    .Source = 事.Source
    .Description = 事.Description
    .HelpFile = 事.HelpFile
    '引数HelpContextに0を与えるとErr().Raise時に1000440に変換されてしまうので、-1を与える
    .HelpContext = IIf(事.HelpContext, -1, 事.HelpContext)
    .LastDllError = 事.LastDllError
    .Erl = 事.Erl

    Err().Number = IIf(.ID > ErrNumber.上限, -.ID, .ID)
    Err().Source = .Source
    Err().Description = .Description
    Err().HelpFile = .HelpFile
    Err().HelpContext = .HelpContext
    
    'Err().LastDllErrorは読取専用なので書き換えるべきときはErr().Descriptionに追記
    If Err().LastDllError <> .LastDllError _
      Then Err().Description = .Description & LastDllError連結句 & .LastDllError
    
  End With
End Property


