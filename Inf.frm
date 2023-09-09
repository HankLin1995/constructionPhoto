VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Inf 
   Caption         =   "歡迎使用"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5835
   OleObjectBlob   =   "Inf.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "Inf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image2_Click()
ActiveWorkbook.FollowHyperlink Address:="https://hankvba.blogspot.com/2018/03/autocad-vba.html", NewWindow:=True
End Sub

Private Sub Label12_Click()
ActiveWorkbook.FollowHyperlink Address:="https://www.youtube.com/watch?v=jvgl9tHrMi0&ab_channel=LinChunHan", NewWindow:=True
End Sub

Private Sub Label14_Click()
ActiveWorkbook.FollowHyperlink Address:="https://docs.google.com/presentation/d/1lKRRtUrsWIv2JSE8on-qyosYz8N29Ms_/edit?usp=sharing&ouid=112944893851556117594&rtpof=true&sd=true", NewWindow:=True
End Sub
