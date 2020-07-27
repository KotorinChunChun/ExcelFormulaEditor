VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsFormulaStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        ClsFormulaStack
Rem
Rem  @description   Stackクラス
Rem
Rem  @update        2020/07/27
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem
Rem  @note          このプログラムは下記を参考に独自の改変を行ったものです。
Rem
Rem  @quote         ExcelでネストしたIf関数をVBAでインデントして分析しやすくする
Rem
Rem  @url           https://thom.hateblo.jp/entry/2017/08/20/122525
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Private items() As Variant

Property Get Count() As Integer
    Count = UBound(items)
End Property

Property Get Top() As Variant
    Top = items(UBound(items))
End Property

Public Function Pop() As Variant
    If UBound(items) > 0 Then
        Pop = items(UBound(items))
        ReDim Preserve items(UBound(items) - 1)
    Else
        Pop = Empty
    End If
End Function

Public Sub Push(ByRef x As Variant)
    ReDim Preserve items(UBound(items) + 1)
    items(UBound(items)) = x
End Sub

Private Sub Class_Initialize()
    ReDim items(0)
End Sub
