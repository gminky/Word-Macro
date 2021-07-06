VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 枠表示切替form 
   Caption         =   "ツール"
   ClientHeight    =   2064
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   5496
   OleObjectBlob   =   "枠表示切替form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "枠表示切替form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub 行末句点スタイルボタン_Click()
    If Selection.Paragraphs.FarEastLineBreakControl = True Then
        Selection.Paragraphs.FarEastLineBreakControl = False
        枠表示切替form.行末句点スタイルボタン.Caption = "設定する"
    Else
        Selection.Paragraphs.FarEastLineBreakControl = True
        枠表示切替form.行末句点スタイルボタン.Caption = "設定しない"
    End If
End Sub

Private Sub 表示ボタン_Click()
    If 表示ボタン.Caption = "表示する" Then
        '非表示にする場合
        表示ボタン.Caption = "表示しない"
        Call 閣議請議用枠表示
    Else
        '表示する場合
        表示ボタン.Caption = "表示する"
        Call 閣議請議用枠非表示
    
    End If
End Sub
