VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Excel処理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim エクセル As Application
Private 範 As 範
Private 出 As New 出

Private Const Excel最大行 = 1048576
Private Const Excel最大列 = 16384

Private Sub Class_Initialize()
  On Error GoTo OnError
  Set エクセル = New Application
  Set 範 = 主.受け渡す()
Exit Sub
OnError:
  
End Sub

Public Sub 処理する(パス As String)
  On Error GoTo OnError
  With エクセル.Workbooks
    Dim i As Long
    For i = 1 To .Count
      If .Item(i).FullName = パス Then
        Exit For
      Else
        .Item(i).Close False
      End If
    Next
    If .Count = 0 Then .Open パス, ReadOnly:=True
    
    Dim シート As Worksheet
    For Each シート In .Item(1).Worksheets
      抽出して転記する シート
    Next シート
  End With
Exit Sub
OnError:

End Sub
  
Private Sub 抽出して転記する(シート As Worksheet)
  On Error GoTo OnError
  With シート
    .AutoFilterMode = False
    Dim 見出し As Range
    Set 見出し = .Range(範.見出Address)
    With 見出し.Cells(1)
      Dim 開始行 As Long
      開始行 = .Row()
      Dim 開始列 As Long
      開始列 = .Column()
    End With
    
    Dim 最終行 As Long
    最終行 = .Cells(Excel最大行, 範.最終行判定列番).End(xlUp)
    Dim 最終列 As Long
    最終列 = .Cells(開始行, Excel最大列).End(xlToLeft)
  End With
  
  Dim 表領域 As Range
  Set 表領域 = 見出し.Resize(最終行 - 開始行 + 1&)
  
  Dim 範Asシート As Worksheet
  Set 範Asシート = 範
  If 範Asシート.AutoFilterMode Then
    表領域.AutoFilter
    
    With 範Asシート.AutoFilter.Filters
      Dim i As Long
  On Error Resume Next
      For i = 1 To .Count
        With .Item(i)
          If .On Then
            表領域.AutoFilter i, .Criteria1
            表領域.AutoFilter i, .Criteria1, .Operator
            表領域.AutoFilter i, .Criteria1, .Operator, .Criteria2
          End If
        End With
      Next
  On Error GoTo OnError
    End With
  End If
  
  With シート
    最終行 = .Cells(Excel最大行, 範.最終行判定列番).End(xlUp)
    If 最終行 > 開始行 _
      Then 出.転記する .Range(.Cells(開始行 + 1&, 1&), .Cells(最終行, 最終列)).SpecialCells(xlCellTypeVisible).Value
  End With
Exit Sub
OnError:
  If Err.Number = 91 Then
    主.受け渡す 範.出
    Set 出 = New 出
    Resume
  End If
End Sub

Private Sub Class_Terminate()
  エクセル.Quit
End Sub
