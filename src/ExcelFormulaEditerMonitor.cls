VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelFormulaEditerMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem
Rem @module
Rem   FormulaEditMonitor
Rem
Rem @description
Rem   �����̓������Z�����_�u���N���b�N�����Ƃ� ExcelEvaluateForm ��\������
Rem
Rem @note
Rem�@ �C���X�^���X�𐶐����邾���ŊĎ����܂��B
Rem
Rem @author
Rem   @KotorinChunChun
Rem
Rem @history
Rem   2020/07/27
Rem
Option Explicit

Public WithEvents app As Excel.Application
Attribute app.VB_VarHelpID = -1

Private Sub app_SheetBeforeDoubleClick(ByVal sh As Object, ByVal Target As Range, Cancel As Boolean)
    Call OpenFormulaEditor(sh, Target, Cancel)
End Sub

Sub OpenFormulaEditor(ByVal sh As Object, ByVal Target As Range, ByRef Cancel As Boolean)
    Set Target = Target(1, 1)
    If Not Target.HasFormula Then Exit Sub
    
    Dim Result
    Result = ExcelFormulaEditorForm.OpenForm(Target)
    
    If Not IsNull(Result) Then
        Target.Formula = Result
    End If
    
    Cancel = True
End Sub

'----------------------------------------
'�R���X�g���N�^
Private Sub Class_Initialize()
    Set app = Application
End Sub