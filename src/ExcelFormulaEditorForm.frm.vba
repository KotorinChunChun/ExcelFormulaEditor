VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExcelFormulaEditorForm 
   Caption         =   "エクセル数式らくらく入力フォーム"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15765
   OleObjectBlob   =   "ExcelFormulaEditorForm.frm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
Rem  @description   Excelの数式を簡単に入力・編集するフォーム
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
Rem    2020/07/27 開発
Rem
Rem --------------------------------------------------------------------------------
Rem
Rem テストデータ
Rem  =IF(TRUE,IF(TRUE,IF(TRUE,1,2),3),4)
Rem
Option Explicit

Private IsCanceled As Boolean
Private Target As Excel.Range

Rem コンストラクタ
Rem
Rem  @note
Rem     フォームを直接実行しても意味はありません。
Rem     OpenFormで開くようにしてください。
Rem
Private Sub UserForm_Initialize()
    With ListBox
        .ColumnCount = 3
        .ColumnHeads = True
        .ColumnWidths = "200;10;100"
        '"値;数式
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
    
    TextBoxInput.EnterKeyBehavior = True   'Enterのフォーカス移動を禁止したかったが効果無かった。
    
    SpinButtonFontSize.Value = TextBoxInput.Font.Size
    SpinButtonFontSize.Min = 8
    SpinButtonFontSize.Max = 30
    
    Me.Hide
End Sub

Rem 引数を指定してフォームを表示する
Rem
Rem @param arr_listitems        リストに表示する全アイテム配列
Rem @param arr_defaultvalues    起動時に選択状態にするアイテム配列
Rem @parem can_multiselect      複数選択の可否
Rem
Rem @return As Variant(0 to n)  選択されていたアイテム配列
Rem                             キャンセル時：Null
Public Function OpenForm(Optional targetRange As Range) As Variant
    Set Target = targetRange
    If Target Is Nothing Then Set Target = Excel.ActiveCell
    
    '再利用された場合のため初期化
    IsCanceled = False
    
    '非インデント状態で表示
    On Error Resume Next
    Dim fmr: fmr = Target.Formula
    CheckBoxFormulaArray.Value = Target.HasArray
    On Error GoTo 0
    Call TextBoxSetFormula(fmr)
    
    Me.Show '←モーダルフォームではここでVBAが止まる
    
    'フォームが終了されたとき/Unloadされたときエラーが出る
    On Error Resume Next
    OpenForm = Null
    OpenForm = Me.Result
    On Error GoTo 0
End Function

Private Sub SpinButtonFontSize_Change()
    On Error Resume Next
    TextBoxInput.Font.Size = SpinButtonFontSize.Value
End Sub

Rem 数式を反映する
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

Rem 数式を取得する
Function TextBoxGetFormula() As String
    Dim fmr: fmr = TextBoxInput.Text
    If Not fmr Like "=*" Then fmr = "=" & fmr
    TextBoxGetFormula = fmr
End Function

Rem 結果の取得
Rem
Rem @return As Variant(0 to n)  選択されていたアイテム配列
Rem                             キャンセル時：Null
Rem @note
Rem  モードレス対応のために公開しているが、
Rem  呼び出し元でUnload対策の処理が複雑になるため
Rem  できる限りOpenForm関数の戻り値を使うべき
Rem
Public Property Get Result() As Variant
    If IsCanceled Then Result = Null: Exit Property
    Result = TextBoxInput.Text
End Property

Rem 数式変更時
Private Sub TextBoxInput_Change()
    Dim fmr: fmr = TextBoxGetFormula()
    Dim v: v = kccFuncExcelFormula.EvaluateEx(fmr, Excel.ActiveCell)
    Dim txtFormula: txtFormula = "" 'セル参照だけを値に置き換える関数未実装
    
    '式の有効性チェック
    TextBoxInput.BackColor = IIf(IsError(v), 16764159, vbWhite)
    
    '式の情報をリストに追加
    If IsError(v) Then v = CStr(v)
    Call kccFuncMSForms.ListBox_AddItem(ListBox, Array(fmr, txtFormula, v), 0)
    ListBox.listIndex = 0
    
    '自動フォーマット結果を表示
    Call AutoFormat(fmr)
End Sub

Rem タブ変更時
Private Sub TabStrip_Change()
    Dim fmr: fmr = TextBoxGetFormula()
    Call AutoFormat(fmr)
    Dim v: Set v = TabStrip.Tabs(TabStrip.Value)
    'アクティブなタブに着色とかして目立たせたい・・が、オーナードローが必須とわかり断念
End Sub

Rem 自動フォーマット
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
    If MsgBox(IIf(isCSE, "配列", "") & "数式をセルに入力します" & vbLf & fmr, vbOKCancel, "数式入力") = vbCancel Then Exit Sub
    
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
    
        'Ctrl+Tabでタブ入力を無効化
        Case vbKeyTab
            If PressCtrl And PressShift Then
                Call CtrlValueIncDec(TabStrip, -1)
                KeyCode = 0
            ElseIf PressCtrl Then
                Call CtrlValueIncDec(TabStrip, 1)
                KeyCode = 0
            End If
'
        'ESCでフォームを閉じる
        Case vbKeyEscape: Cancel_Button_Click
        
        'F5でフォーマット済みの式を採用
        Case vbKeyF5: Call TextBoxSetFormula(TextBoxFormated.Text)
        
        'Enterで送信 (改行ON時はCtrl+Enter） / CSEで配列数式 / チェックONでも配列数式
        Case vbKeyReturn
            If PressCtrl And PressShift Then
                KeyCode = 0
                Call WriteFormula(True)
            ElseIf TextBoxInput.MultiLine = False Or PressCtrl Then
                KeyCode = 0
                Call WriteFormula(False Or CheckBoxFormulaArray.Value)
            End If
            
        'Alt+0〜9でタブ切り替え / 連続実行で採用
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

'範囲制限のあるプロパティにはみ出さないように加算減算を行う
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
