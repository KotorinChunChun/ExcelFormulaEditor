Attribute VB_Name = "AppMain"
Rem
Rem @appname ExcelFormulaEditor - エクセル数式らくらく入力アドイン
Rem
Rem @module AppMain
Rem
Rem @author @KotorinChunChun
Rem
Rem @update
Rem    2020/07/27 0.10 初回版
Rem    2020/07/28 0.11 若干バグ修正
Rem    2021/01/15 0.12 配列数式対応と操作性向上
Rem
Option Explicit
Option Private Module

Public Const APP_NAME = "エクセル数式らくらく入力アドイン"
Public Const APP_CREATER = "@KotorinChunChun"
Public Const APP_VERSION = "0.12"
Public Const APP_UPDATE = "2021/01/15"
Public Const APP_URL = "https://github.com/KotorinChunChun/ExcelFormulaEditor"

Public instExcelFormulaEditerMonitor As ExcelFormulaEditerMonitor

Rem --------------------------------------------------
Rem アドイン実行時
Sub AddinStart()
    MsgBox "数式拡張入力を有効にします！！！" & vbLf & _
            "" & vbLf & _
            "", _
                vbInformation + vbOKOnly, APP_NAME
    Call MonitorStart
End Sub

Rem アドイン一時停止時
Sub AddinStop()
    Dim item
    For Each item In Array( _
        "監視を停止しますか？", _
        "ほんとにやめちゃうの？")
        If MsgBox(item, vbExclamation + vbYesNo, APP_NAME) = vbNo Then
            MsgBox "ありがと～～～", vbOKOnly, APP_NAME
            Exit Sub
        End If
    Next
    MsgBox "またあそんでね？", vbOKOnly, APP_NAME
    Call MonitorStop
End Sub

Rem アドイン設定表示
Sub AddinConfig():: End Sub '  Call SettingForm.Show

Rem アドイン情報表示
Sub AddinInfo()
    Select Case MsgBox(ThisWorkbook.Name & vbLf & vbLf & _
            "バージョン : " & APP_VERSION & vbLf & _
            "更新日　　 : " & APP_UPDATE & vbLf & _
            "開発者　　 : " & APP_CREATER & vbLf & _
            "実行パス　 : " & ThisWorkbook.Path & vbLf & _
            "公開ページ : " & APP_URL & vbLf & _
            vbLf & _
            "使い方や最新版を探しに公開ページを開きますか？" & _
            "", vbInformation + vbYesNo, "バージョン情報")
        Case vbNo
            '
        Case vbYes
            CreateObject("Wscript.Shell").Run APP_URL, 3
    End Select
End Sub

Rem アドイン完全終了
Sub AddinEnd(): ThisWorkbook.Close False: End Sub
Rem --------------------------------------------------

Rem 監視開始
Rem Workbook_Openから呼ばれる
Sub MonitorStart()
    Set instExcelFormulaEditerMonitor = New ExcelFormulaEditerMonitor
    'Ctrl+F2
    Application.OnKey "^2", "OpenFormulaEditorForm"
End Sub

Rem 監視停止
Sub MonitorStop()
Set instExcelFormulaEditerMonitor = Nothing
    Application.OnKey "^2"
End Sub

Rem フォーム実行
Sub OpenFormulaEditorForm()
    Dim rng As Range: Set rng = Excel.ActiveCell
    Dim fmr: fmr = ExcelFormulaEditorForm.OpenForm(rng)
    If Not IsNull(fmr) Then
        rng.Formula = fmr
    End If
End Sub
