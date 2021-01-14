Attribute VB_Name = "AppMain"
Rem
Rem @appname ExcelFormulaEditor - �G�N�Z�������炭�炭���̓A�h�C��
Rem
Rem @module AppMain
Rem
Rem @author @KotorinChunChun
Rem
Rem @update
Rem    2020/07/27 0.10 �����
Rem    2020/07/28 0.11 �኱�o�O�C��
Rem    2021/01/15 0.12 �z�񐔎��Ή��Ƒ��쐫����
Rem
Option Explicit
Option Private Module

Public Const APP_NAME = "�G�N�Z�������炭�炭���̓A�h�C��"
Public Const APP_CREATER = "@KotorinChunChun"
Public Const APP_VERSION = "0.12"
Public Const APP_UPDATE = "2021/01/15"
Public Const APP_URL = "https://github.com/KotorinChunChun/ExcelFormulaEditor"

Public instExcelFormulaEditerMonitor As ExcelFormulaEditerMonitor

Rem --------------------------------------------------
Rem �A�h�C�����s��
Sub AddinStart()
    MsgBox "�����g�����͂�L���ɂ��܂��I�I�I" & vbLf & _
            "" & vbLf & _
            "", _
                vbInformation + vbOKOnly, APP_NAME
    Call MonitorStart
End Sub

Rem �A�h�C���ꎞ��~��
Sub AddinStop()
    Dim item
    For Each item In Array( _
        "�Ď����~���܂����H", _
        "�ق�Ƃɂ�߂��Ⴄ�́H")
        If MsgBox(item, vbExclamation + vbYesNo, APP_NAME) = vbNo Then
            MsgBox "���肪�Ɓ`�`�`", vbOKOnly, APP_NAME
            Exit Sub
        End If
    Next
    MsgBox "�܂�������łˁH", vbOKOnly, APP_NAME
    Call MonitorStop
End Sub

Rem �A�h�C���ݒ�\��
Sub AddinConfig():: End Sub '  Call SettingForm.Show

Rem �A�h�C�����\��
Sub AddinInfo()
    Select Case MsgBox(ThisWorkbook.Name & vbLf & vbLf & _
            "�o�[�W���� : " & APP_VERSION & vbLf & _
            "�X�V���@�@ : " & APP_UPDATE & vbLf & _
            "�J���ҁ@�@ : " & APP_CREATER & vbLf & _
            "���s�p�X�@ : " & ThisWorkbook.Path & vbLf & _
            "���J�y�[�W : " & APP_URL & vbLf & _
            vbLf & _
            "�g������ŐV�ł�T���Ɍ��J�y�[�W���J���܂����H" & _
            "", vbInformation + vbYesNo, "�o�[�W�������")
        Case vbNo
            '
        Case vbYes
            CreateObject("Wscript.Shell").Run APP_URL, 3
    End Select
End Sub

Rem �A�h�C�����S�I��
Sub AddinEnd(): ThisWorkbook.Close False: End Sub
Rem --------------------------------------------------

Rem �Ď��J�n
Rem Workbook_Open����Ă΂��
Sub MonitorStart()
    Set instExcelFormulaEditerMonitor = New ExcelFormulaEditerMonitor
    'Ctrl+F2
    Application.OnKey "^2", "OpenFormulaEditorForm"
End Sub

Rem �Ď���~
Sub MonitorStop()
Set instExcelFormulaEditerMonitor = Nothing
    Application.OnKey "^2"
End Sub

Rem �t�H�[�����s
Sub OpenFormulaEditorForm()
    Dim rng As Range: Set rng = Excel.ActiveCell
    Dim fmr: fmr = ExcelFormulaEditorForm.OpenForm(rng)
    If Not IsNull(fmr) Then
        rng.Formula = fmr
    End If
End Sub
