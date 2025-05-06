VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "出"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private 実 As 出

Private Sub Class_Initialize()
  Set 実 = 主.受け渡す
  実.Initialize
End Sub

Public Sub Initialize()
End Sub

Public Sub 転記する(データ As Variant)
  On Error GoTo OnError
Exit Sub
OnError:

End Sub

Public Sub Terminate()
End Sub

Private Sub Class_Terminate()
  実.Terminate
End Sub
