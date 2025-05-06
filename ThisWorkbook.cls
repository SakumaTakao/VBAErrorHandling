VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const ルートパス As String = "C:\Users\dqg01\OneDrive\VBA\巡庫"

Public Sub Sheet1の処理をルートフォルダ内の全ファイルに対して行う()
  巡庫.主.フォルダ内の全ファイルを処理する ルートパス, Sheet1
End Sub

