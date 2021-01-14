VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExcelFormulaEditorForm 
   Caption         =   "�G�N�Z�������炭�炭���̓t�H�[��"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15765
   OleObjectBlob   =   "ExcelFormulaEditorForm.frm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "ExcelFormulaEditorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        ExcelFormulaEditorForm
Rem
Rem  @description   Excel�̐������ȒP�ɓ��́E�ҏW����t�H�[��
Rem
Rem  @update        2020/07/27
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem  @references
Rem    Excel.Application
Rem
Rem --------------------------------------------------------------------------------
Rem  @refModules
Rem    FuncExcelFormula
Rem    FuncMSForms
Rem
Rem --------------------------------------------------------------------------------
Rem  @history
Rem    2020/07/27 �J��
Rem
Rem --------------------------------------------------------------------------------
Rem
Rem �e�X�g�f�[�^
Rem  =IF(TRUE,IF(TRUE,IF(TRUE,1,2),3),4)
Rem
Option Explicit

Private IsCanceled As Boolean
Private Target As Excel.Range

Rem �R���X�g���N�^
Rem
Rem  @note
Rem     �t�H�[���𒼐ڎ��s���Ă��Ӗ��͂���܂���B
Rem     OpenForm�ŊJ���悤�ɂ��Ă��������B
Rem
Private Sub UserForm_Initialize()
    With ListBox
        .ColumnCount = 3
        .ColumnHeads = True
        .ColumnWidths = "200;10;100"
        '"�l;����
    End With
    
    With TabStrip
        .Tabs(0).caption = "0:Single"
        Dim i As Long
        For i = 1 To 9
            .Tabs.add "" & i & ":Block"
        Next
        .Tabs(9).caption = "9:Tree"
        .Value = 5 - 1
    End With
    
    TextBoxFormated.MultiLine = True
    TextBoxFormated.Locked = True
    
    TextBoxInput.EnterKeyBehavior = True   'Enter�̃t�H�[�J�X�ړ����֎~���������������ʖ��������B
    
    SpinButtonFontSize.Value = TextBoxInput.Font.Size
    SpinButtonFontSize.Min = 8
    SpinButtonFontSize.Max = 30
    
    Me.Hide
End Sub

Rem �������w�肵�ăt�H�[����\������
Rem
Rem @param arr_listitems        ���X�g�ɕ\������S�A�C�e���z��
Rem @param arr_defaultvalues    �N�����ɑI����Ԃɂ���A�C�e���z��
Rem @parem can_multiselect      �����I���̉�
Rem
Rem @return As Variant(0 to n)  �I������Ă����A�C�e���z��
Rem                             �L�����Z�����FNull
Public Function OpenForm(Optional targetRange As Range) As Variant
    Set Target = targetRange
    If Target Is Nothing Then Set Target = Excel.ActiveCell
    
    '�ė��p���ꂽ�ꍇ�̂��ߏ�����
    IsCanceled = False
    
    '��C���f���g��Ԃŕ\��
    On Error Resume Next
    Dim fmr: fmr = Target.Formula
    CheckBoxFormulaArray.Value = Target.HasArray
    On Error GoTo 0
    Call TextBoxSetFormula(fmr)
    
    Me.Show '�����[�_���t�H�[���ł͂�����VBA���~�܂�
    
    '�t�H�[�����I�����ꂽ�Ƃ�/Unload���ꂽ�Ƃ��G���[���o��
    On Error Resume Next
    OpenForm = Null
    OpenForm = Me.Result
    On Error GoTo 0
End Function

Private Sub SpinButtonFontSize_Change()
    On Error Resume Next
    TextBoxInput.Font.Size = SpinButtonFontSize.Value
End Sub

Rem �����𔽉f����
Sub TextBoxSetFormula(fmr)
    
    If InStr(fmr, vbLf) > 0 Then
        TextBoxInput.MultiLine = True
        TextBoxInput.Height = 200
    Else
        TextBoxInput.MultiLine = False
        TextBoxInput.Height = 24
    End If
    
    ListBox.Top = TextBoxInput.Top + TextBoxInput.Height + 10
    TextBoxInput.Text = fmr
End Sub

Rem �������擾����
Function TextBoxGetFormula() As String
    Dim fmr: fmr = TextBoxInput.Text
    If Not fmr Like "=*" Then fmr = "=" & fmr
    TextBoxGetFormula = fmr
End Function

Rem ���ʂ̎擾
Rem
Rem @return As Variant(0 to n)  �I������Ă����A�C�e���z��
Rem                             �L�����Z�����FNull
Rem @note
Rem  ���[�h���X�Ή��̂��߂Ɍ��J���Ă��邪�A
Rem  �Ăяo������Unload�΍�̏��������G�ɂȂ邽��
Rem  �ł������OpenForm�֐��̖߂�l���g���ׂ�
Rem
Public Property Get Result() As Variant
    If IsCanceled Then Result = Null: Exit Property
    Result = TextBoxInput.Text
End Property

Rem �����ύX��
Private Sub TextBoxInput_Change()
    Dim fmr: fmr = TextBoxGetFormula()
    Dim v: v = kccFuncExcelFormula.EvaluateEx(fmr, Excel.ActiveCell)
    Dim txtFormula: txtFormula = "" '�Z���Q�Ƃ�����l�ɒu��������֐�������
    
    '���̗L�����`�F�b�N
    TextBoxInput.BackColor = IIf(IsError(v), 16764159, vbWhite)
    
    '���̏������X�g�ɒǉ�
    If IsError(v) Then v = CStr(v)
    Call kccFuncMSForms.ListBox_AddItem(ListBox, Array(fmr, txtFormula, v), 0)
    ListBox.listIndex = 0
    
    '�����t�H�[�}�b�g���ʂ�\��
    Call AutoFormat(fmr)
End Sub

Rem �^�u�ύX��
Private Sub TabStrip_Change()
    Dim fmr: fmr = TextBoxGetFormula()
    Call AutoFormat(fmr)
    Dim v: Set v = TabStrip.Tabs(TabStrip.Value)
    '�A�N�e�B�u�ȃ^�u�ɒ��F�Ƃ����Ėڗ����������E�E���A�I�[�i�[�h���[���K�{�Ƃ킩��f�O
End Sub

Rem �����t�H�[�}�b�g
Private Sub AutoFormat(fmr)
    If TabStrip.Value = 9 Then
        TextBoxFormated.Text = kccFuncExcelFormula.FormulaIndentTree(fmr)
    Else
        Dim level: level = TabStrip.Value + 1
        TextBoxFormated.Text = kccFuncExcelFormula.FormulaIndentBlock(fmr, level)
    End If
End Sub

Rem ----------

'Sub OK_Button_Click()
'    Call WriteFormula(False)
'End Sub

Sub Cancel_Button_Click()
    Unload Me
End Sub

Sub WriteFormula(isCSE As Boolean)
    Dim fmr: fmr = TextBoxGetFormula()
    Dim v: v = kccFuncExcelFormula.EvaluateEx(fmr, Excel.ActiveCell)
    If IsError(v) Then Exit Sub
    
    fmr = kccFuncExcelFormula.ReplaceByRange(fmr, Target)
    If MsgBox(IIf(isCSE, "�z��", "") & "�������Z���ɓ��͂��܂�" & vbLf & fmr, vbOKCancel, "��������") = vbCancel Then Exit Sub
    
    If isCSE Then
        Excel.ActiveCell.FormulaArray = Replace(fmr, vbCr, "")
    Else
        Excel.ActiveCell.Formula = Replace(fmr, vbCr, "")
    End If
    
    Unload Me
End Sub

Rem ----------

Sub TextBoxInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): Call Any_KeyDown(KeyCode, Shift): End Sub
Sub ListBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): Call Any_KeyDown(KeyCode, Shift): End Sub
Sub TextBoxFormated_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): Call Any_KeyDown(KeyCode, Shift): End Sub
Sub TabStrip_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): Call Any_KeyDown(KeyCode, Shift): End Sub
Sub SpinButtonFontSize_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): Call Any_KeyDown(KeyCode, Shift): End Sub
Sub CheckBoxFormulaArray_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): Call Any_KeyDown(KeyCode, Shift): End Sub

Sub Any_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim PressShift As Boolean: PressShift = Shift And 1
    Dim PressCtrl As Boolean: PressCtrl = Shift And 2
    Dim PressAlt As Boolean: PressAlt = Shift And 4
    If KeyCode = 16 Or KeyCode = 17 Or KeyCode = 18 Then Exit Sub
    
    Select Case KeyCode.Value
    
        'Ctrl+Tab�Ń^�u���͂𖳌���
        Case vbKeyTab
            If PressCtrl And PressShift Then
                Call CtrlValueIncDec(TabStrip, -1)
                KeyCode = 0
            ElseIf PressCtrl Then
                Call CtrlValueIncDec(TabStrip, 1)
                KeyCode = 0
            End If
'
        'ESC�Ńt�H�[�������
        Case vbKeyEscape: Cancel_Button_Click
        
        'F5�Ńt�H�[�}�b�g�ς݂̎����̗p
        Case vbKeyF5: Call TextBoxSetFormula(TextBoxFormated.Text)
        
        'Enter�ő��M (���sON����Ctrl+Enter�j / CSE�Ŕz�񐔎� / �`�F�b�NON�ł��z�񐔎�
        Case vbKeyReturn
            If PressCtrl And PressShift Then
                KeyCode = 0
                Call WriteFormula(True)
            ElseIf TextBoxInput.MultiLine = False Or PressCtrl Then
                KeyCode = 0
                Call WriteFormula(False Or CheckBoxFormulaArray.Value)
            End If
            
        'Alt+0�`9�Ń^�u�؂�ւ� / �A�����s�ō̗p
        Case vbKey0 To vbKey9
            If PressAlt Then
                If TabStrip.Value = KeyCode - 48 Then
                    Call TextBoxSetFormula(TextBoxFormated.Text)
                Else
                    TabStrip.Value = KeyCode - 48
                    KeyCode = 0
                End If
            End If
    End Select
End Sub

'�͈͐����̂���v���p�e�B�ɂ͂ݏo���Ȃ��悤�ɉ��Z���Z���s��
Sub CtrlValueIncDec(ctl As Object, num As Long)
    Dim Min As Long, Max As Long, Value As Long
    If TypeName(ctl) = "TabStrip" Then
        Min = 0
        Max = ctl.Tabs.Count - 1
        Value = ctl.Value
    ElseIf TypeName(ctl) = "SpinButton" Then
        Min = ctl.Min
        Max = ctl.Max
        Value = ctl.Value
    Else
        Debug.Print TypeName(ctl)
        Stop
    End If
    Value = Value + num
    Value = IIf(Value > Max, Max, IIf(Value < Min, Min, Value))
    ctl.Value = Value
End Sub
