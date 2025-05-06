VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "訪"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private 実 As 訪
Private デフォルト処理 As Excel処理

Private Sub Class_Initialize()
  Set 実 = 主.受け渡す()
End Sub

Friend Sub フォルダ内の全ファイルを処理する(ルートパス As String)
  On Error GoTo OnError
  With New IWshRuntimeLibrary.FileSystemObject
    訪れる .GetFolder(ルートパス)
  End With
Exit Sub
OnError:
End Sub

Private Sub 訪れる(フォルダ As IWshRuntimeLibrary.Folder)
  On Error GoTo OnError
  
  Dim ファイル As IWshRuntimeLibrary.File
  For Each ファイル In フォルダ.Files
    処理する ファイル
  Next ファイル
  Dim サブフォルダ As IWshRuntimeLibrary.Folder
  For Each サブフォルダ In フォルダ.SubFolders
    If Not 実.対象外(サブフォルダ) Then 訪れる サブフォルダ
  Next サブフォルダ
Exit Sub
OnError:
End Sub

Public Sub 処理する(ファイル As IWshRuntimeLibrary.File)
  On Error GoTo OnError
  実.処理する ファイル
  デフォルト処理.処理する ファイル.Path
Exit Sub
OnError:
  Select Case Err().Number
    Case 91
      主.受け渡す 実
      Set デフォルト処理 = New Excel処理
      Resume
    Case 上書.サブクラスの処理のみ行う
      Exit Sub
    Case Else
  End Select
End Sub

'抽象メソッド
Public Function 対象外(フォルダ As IWshRuntimeLibrary.Folder) As Boolean
End Function
