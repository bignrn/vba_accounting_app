Attribute VB_Name = "Module1"
'*********************
'Main Class
'FaileName :Module1
'FainalDate:2020.11.25
'
'@ather nora
'@version 1.0
'
'参考資料
'https://www.sejuku.net/blog/35484
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
On Error GoTo Exception
    UserForm1.Show
    Exit Sub
Exception:
        MsgBox Err.Number & vbCrLf & Err.Description
End Sub

'====================
'
'関数：EraseUserFoemOPEN_Click
'用途：クリックするとユーザーフォームが立ち上がる
'
'====================
Sub EraseUserFormOPEN_Click()
On Error GoTo Exception
    UserForm2.Show
    Exit Sub
Exception:
        MsgBox Err.Number & vbCrLf & Err.Description
End Sub

'====================
'
'関数：BackSheet_Click
'用途：クリックしたら前のシートに戻る
'
'====================
Sub BackSheet_Click()
On Error GoTo Exception
    Worksheets(sheetNameGlobal).Activate
    Exit Sub
Exception:
        MsgBox Err.Number & vbCrLf & Err.Description
End Sub

