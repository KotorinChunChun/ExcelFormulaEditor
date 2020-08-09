VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kccFuncExcelFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccFuncExcelFormula
Rem
Rem  @description   Excel�̊֐������p�[�X���ď������郂�W���[��
Rem
Rem  @update        2020/07/27
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem  @references
Rem    �s�v
Rem
Rem --------------------------------------------------------------------------------
Rem  @refModules
Rem    ClsFormulaExpression
Rem    ClsFormulaStack
Rem    ClsFormulaToken
Rem
Rem --------------------------------------------------------------------------------
Rem  @note
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Rem --------------------------------------------------------------------------------
Rem
Rem  @note          ���̃v���O�����͉��L���Q�l�ɓƎ��̉��ς��s�������̂ł��B
Rem
Rem  @quote         Excel�Ńl�X�g����If�֐���VBA�ŃC���f���g���ĕ��͂��₷������
Rem
Rem  @url           https://thom.hateblo.jp/entry/2017/08/20/122525
Rem
Rem --------------------------------------------------------------------------------
Public Enum TokenType
    Target      'If
    BeginParen  '(
    EndParen    ')
    Comma       ',
    LineFeed    'vbLF
    Exclamation '!
    ErrCode     '#�`!
    Other
End Enum

Public Enum CharType
    Alphabet
    Number
    BeginParen
    EndParen
    Comma
    DoubleQuote
    SingleQuote
    LineFeed
    Exclamation
    NumberSign
    None
    Other
End Enum

Private Function GetCharType(C) As CharType
    If C = "" Then GetCharType = CharType.None: Exit Function
    Dim ret As CharType
    Select Case Asc(C)
        Case Asc("a") To Asc("z"), Asc("A") To Asc("Z")
            ret = CharType.Alphabet
        Case Asc("0") To Asc("9")
            ret = CharType.Number
        Case Else
            Select Case True
                Case C = "(": ret = CharType.BeginParen
                Case C = ")": ret = CharType.EndParen
                Case C = ",": ret = CharType.Comma
                Case C = """": ret = CharType.DoubleQuote
                Case C = "'": ret = CharType.SingleQuote
                Case C = vbLf: ret = CharType.LineFeed
                Case C = "!": ret = CharType.Exclamation
                Case C = "#": ret = CharType.NumberSign
                Case Else
                    ret = CharType.Other
            End Select
    End Select
    GetCharType = ret
End Function

Private Function IsIn(target_, ParamArray check()) As Boolean
    Dim i As Long, ret As Boolean: ret = False
    For i = LBound(check) To UBound(check)
        ret = ret Or check(i) = target_
    Next
    IsIn = ret
End Function

Public Function GetTokens(targetExpression As ClsFormulaExpression) As Collection
    Dim ret As Collection: Set ret = New Collection
    
    Dim t As ClsFormulaToken
    Do While targetExpression.hasNext
        Set t = New ClsFormulaToken
        t.tString = targetExpression.getNext
        Select Case GetCharType(t.tString)
            Case CharType.Alphabet
                '�p����n�܂�p���̓Z���A�h���X�̂��ߌ���
                Do While IsIn(GetCharType(targetExpression.checkNext), CharType.Alphabet, CharType.Number)
                    t.AddChar targetExpression.getNext
                Loop
            Case CharType.Number
                '�����E�E�E�͐��l�̂��ߌ���
                Do While GetCharType(targetExpression.checkNext) = CharType.Number
                    t.AddChar targetExpression.getNext
                Loop
            Case CharType.DoubleQuote
                '�_�u���N�H�[�e�[�V�����͕���܂Ō���
                Do While GetCharType(targetExpression.checkNext) <> CharType.DoubleQuote
                    t.AddChar targetExpression.getNext
                Loop
                t.AddChar targetExpression.getNext
            Case CharType.SingleQuote
                '�V���O���N�H�[�e�[�V�����͕���܂Ō���
                Do While GetCharType(targetExpression.checkNext) <> CharType.SingleQuote
                    t.AddChar targetExpression.getNext
                Loop
                t.AddChar targetExpression.getNext
            Case CharType.NumberSign
                '�V���[�v�̓G���[�R�[�h�B!�ŕ���܂Ō���
                Do While GetCharType(targetExpression.checkNext) <> CharType.Exclamation
                    t.AddChar targetExpression.getNext
                Loop
                t.AddChar targetExpression.getNext
            Case CharType.BeginParen
            Case CharType.EndParen
            Case CharType.Comma
            Case CharType.LineFeed
            Case CharType.Exclamation
            Case CharType.Other
        End Select
        ret.add t
    Loop
    
    '!�ɂ�錋��
    '1.Alphabet or Number or Other + ! + Other��OK
    Dim ret2 As Collection: Set ret2 = New Collection
    Dim t1 As ClsFormulaToken, t2 As ClsFormulaToken, t3 As ClsFormulaToken
    Dim i As Long
    For i = 1 To ret.Count
        Set t1 = ret(i + 0)
        If i < ret.Count - 1 Then
            Set t2 = ret(i + 1)
            If t2.tType = TokenType.Exclamation And i < ret.Count - 2 Then
                Set t3 = ret(i + 2)
                If t3.tType = TokenType.Other Then
                    Set t = New ClsFormulaToken
                    t.tString = t1.tString & t2.tString & t3.tString
                    ret2.add t
                    i = i + 2
                Else
                    ret2.add t1
                End If
            Else
                ret2.add t1
            End If
        Else
            ret2.add t1
        End If
    Next
    
    Set GetTokens = ret2
End Function

Rem �֐������C�C�����ɉ��s�����ăC���f���g����Q�i�b��Łj
Rem
Rem  @param fmr         Excel�֐����̕�����
Rem
Rem  @return As String  �C���f���g�ς݂̕�����
Rem
Rem  @note
Rem    https://thom.hateblo.jp/entry/2017/08/20/122525
Rem    Excel�Ńl�X�g����If�֐���VBA�ŃC���f���g���ĕ��͂��₷������
Rem    EXCEL�֐��̍\����̓R�[�h
Rem
Rem  @example
Rem    Before
Rem      =IF(C3="","",IF(C3>60,"��",IF(C3>30,"��","�~"))
Rem
Rem    After
Rem      =IF(C3="",
Rem          "",
Rem          IF(C3>60,
Rem              "��",
Rem              IF(C3>30,
Rem                  "��",
Rem                  "�~"))
Rem
Public Function FormulaIndentTree(ByVal fmr) As String
    If Not fmr Like "=*" Then FormulaIndentTree = fmr: Exit Function
    fmr = FormulaIndentTree_RemoveFormat(fmr)
        
    Dim targetExpression As ClsFormulaExpression: Set targetExpression = New ClsFormulaExpression
    targetExpression.ExpressionString = fmr
    
    Dim retVal As String
    Dim tokens As Collection: Set tokens = GetTokens(targetExpression)
    Dim t As ClsFormulaToken
    Dim i As Long
    Dim st As ClsFormulaStack: Set st = New ClsFormulaStack
    Dim tabCount As Long
    
    For i = 1 To tokens.Count
        Set t = tokens(i)
        retVal = retVal & t.tString
        Select Case t.tType
            Case TokenType.BeginParen
                If tokens(i - 1).tType = TokenType.Target Then
                    st.Push True
                    tabCount = tabCount + 1
                Else
                    st.Push False
                End If
            Case TokenType.EndParen
                If st.Pop Then tabCount = tabCount - 1
            Case TokenType.Comma
                If st.Top Then
                    retVal = retVal & vbLf & String(tabCount * 4, " ")
                End If
            Case Else
        End Select
    Next
    
    FormulaIndentTree = retVal
End Function

Public Function FormulaIndentTree_RemoveFormat(ByVal fmr As String) As String
    If Not fmr Like "=*" Then FormulaIndentTree_RemoveFormat = fmr: Exit Function
    fmr = Replace(fmr, vbCrLf, vbLf)
    fmr = Replace(fmr, vbCr, vbLf)
        
    Dim targetExpression As ClsFormulaExpression: Set targetExpression = New ClsFormulaExpression
    targetExpression.ExpressionString = fmr
        
    Dim retVal As String
    Dim tokens As Collection: Set tokens = GetTokens(targetExpression)
    Dim t As ClsFormulaToken
    Dim i As Long
    Dim st As ClsFormulaStack: Set st = New ClsFormulaStack
    Dim tabCount As Long
    
    For i = 1 To tokens.Count
        Set t = tokens(i)
        If t.tString <> " " And t.tString <> vbLf Then
            retVal = retVal & t.tString
        End If
    Next
    
    FormulaIndentTree_RemoveFormat = retVal
End Function

Rem �֐������C�C�����ɉ��s�����ăC���f���g����P�i�b��Łj
Rem
Rem  @param fmr         Excel�֐����̕�����
Rem  @param indentLevel �ő剽��܂ŃC���f���g���邩(0~)
Rem
Rem  @return As String  �C���f���g�ς݂̕�����
Rem
Rem  @note
Rem    �ۑ�
Rem      ��������n��̎�����C���f���g������������
Rem      ��U��ԉ��̃C���f���g�����Z�肷��͓̂���B������߁B
Rem
Rem      ��ԉ��̊֐������͉��s���Ȃ��Ƃ��A�������ŏ�������߂�Ƃ��������B
Rem
Rem    �d�l
Rem      �s���ƂɃg���~���O����1�s�ɍ���
Rem      (��������A���s�{�C���N�������g
Rem      ,��������A���s
Rem      )��������A���s�{�f�N�������g
Rem
Rem      �s���ȕ��J�b�R���o����ȍ~�������Ȃ�
Rem
Rem  @example
Rem    Program
Rem        FormulaIndentBlock(Selection.formula,3)
Rem
Rem    Before
Rem        =IF(C3="","",IF(C3>60,"��",IF(C3>30,"��","�~"))
Rem
Rem    After
Rem        =IF(
Rem          C3="",
Rem          "",
Rem          IF(
Rem            C3>60,
Rem            "��",
Rem            IF(C3>30,"��","�~")
Rem          )
Rem        )
Rem
Public Function FormulaIndentBlock(ByVal fmr, Optional ByVal indentLevel = 2) As String
    Dim ins As Variant
    Dim ous As String
    fmr = Replace(fmr, vbCrLf, vbLf)
    fmr = Replace(fmr, vbCr, vbLf)
    For Each ins In Split(fmr, vbLf)
        ous = ous & Trim(ins)
    Next
    
    '������""�Ɉ͂܂�Ă���X�y�[�X�Ɖ��s�͎��O�ɒu������
    '�ȍ~�̃v���O�������ȑf������̂�����Ǝv����B
    
    '���s�}��
    ous = Replace(ous, "(", "(" & vbLf)
    ous = Replace(ous, ")", vbLf & ")")
    ous = Replace(ous, ",", "," & vbLf)
    
    Dim i As Long, j As Long, k As Long
    Dim indent As Long
    
    '���s�񐔕��͂���
    'n�ȏ�Ȃ�([LF]����[LF])�܂ł�[LF]�������B
    Dim ids As String
    ids = Left(ous, 1)
    indent = 0
    For i = 2 To Len(ous)
        If Mid(ous, i - 1, 2) = vbLf & ")" Then indent = indent - 1
        If Mid(ous, i - 1, 2) = "(" & vbLf Then indent = indent + 1
        If Mid(ous, i, 1) <> vbLf Or indent < indentLevel Then
            ids = ids & Mid(ous, i, 1)
        End If
    Next
    ous = ids
    
    '�X�y�[�Xx2�}��
    Dim s1() As String
    Dim s2() As String
    s1 = Split(ous, vbLf)
    ous = ""
    indent = 0
    For i = LBound(s1) To UBound(s1)
        If s1(i) Like ")*" Then indent = indent - 1
        If indent < 0 Then
            '���J�b�R�ߑ��ɂ��������f
            ous = ous & s1(i)
        Else
            If i <> LBound(s1) Then ous = ous & vbLf
            ous = ous & String(indent * 2, " ") & s1(i)
            If s1(i) Like "*(" Then indent = indent + 1
        End If
    Next
    
    FormulaIndentBlock = ous
End Function

Rem rng�̃v���p�e�B�Œu���ł���Evaluate�̃��b�p�[�֐�
Rem
Rem  @param fmr         Evaluate�ɂ����鐔��������
Rem  @param rng         Range�I�u�W�F�N�g
Rem
Rem  @return As Variant �����̌v�Z���ʁ@�v�Z���s���͉����Ȃ��G���[��Ԃ�
Rem
Public Function EvaluateEx(ByVal fmr, Optional ByVal rng As Range) As Variant
    '�����t�H�[�}�b�gOFF�ɂ��Ȃ���Evaluate�ł��Ȃ�
    fmr = FormulaIndentBlock(fmr, 0)
    
    Dim f: f = ReplaceByRange(fmr, rng)
    Dim v: v = Application.Evaluate(f)
    If IsError(v) Then
        Debug.Print Join(Array("EvaluateEx Error", _
                            "Format      := " & fmr, _
                            "Formula     := " & f, _
                            "Evaluate    := " & CStr(v), _
                            ""), vbLf)
    End If
    
    EvaluateEx = v
End Function

Rem ������̂����K��̕�����Range�̃v���p�e�B�֒u��
Rem
Rem  @param fmr         �ϊ��������� �K��̃��[�h��[�`]
Rem  @param rng         Range�I�u�W�F�N�g
Rem
Rem  @return As String  �ϊ��㕶����
Rem
Public Function ReplaceByRange(ByVal fmr, ByVal rng As Range) As String
    ReplaceByRange = fmr
    If rng Is Nothing Then Exit Function
    Dim f: f = fmr
    If InStr(f, "[value]") > 0 Then f = Replace(f, "[value]", rng.Value)
    If InStr(f, "[text]") > 0 Then f = Replace(f, "[text]", """" & rng.Text & """") '0001�΍�
    If InStr(f, "[formula]") > 0 Then f = Replace(f, "[formula]", rng.Formula)
    If InStr(f, "[formular1c1]") > 0 Then f = Replace(f, "[formular1c1]", rng.FormulaR1C1)
    If InStr(f, "[row]") > 0 Then f = Replace(f, "[row]", rng.Row)
    If InStr(f, "[col]") > 0 Then f = Replace(f, "[col]", rng.Column)
    If InStr(f, "[column]") > 0 Then f = Replace(f, "[column]", rng.Column)
    ReplaceByRange = f
End Function


