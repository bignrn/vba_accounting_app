Attribute VB_Name = "Module1"
'*********************
'Main Class
'FaileName :Module1
'FainalDate:2020.11.25
'
'@ather nora
'@version 1.0
'*********************

'グローバル変数
Option Explicit
Public Const sheetNameGlobal = "Sheet1"
Public Const sheetNameGlobal2 = "Sheet2"
Public vertion As String
Public LastLine As Integer

'====================
'
'関数：UserFoemOPEN_Click
'用途：クリックするとユーザーフォームが立ち上がる
'
'====================
Sub UserFormOPEN_Click()
    UserForm1.Show
End Sub

'====================
'
'関数：EraseUserFoemOPEN_Click
'用途：クリックするとユーザーフォームが立ち上がる
'
'====================
Sub EraseUserFormOPEN_Click()
    UserForm2.Show
End Sub

'====================
'
'関数：BackSheet_Click
'用途：クリックしたら前のシートに戻る
'
'====================
Sub BackSheet_Click()
    Worksheets(sheetNameGlobal).Activate
End Sub

