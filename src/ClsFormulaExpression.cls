VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsFormulaExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        ClsFormulaExpression
Rem
Rem  @description   Expressionクラス
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

Public ExpressionString
Private cursor

Private Sub Class_Initialize()
    cursor = 1
End Sub

Function hasNext() As Boolean
    hasNext = Len(ExpressionString) > cursor - 1
End Function

Function getNext() As String
    getNext = Mid(ExpressionString, cursor, 1)
    cursor = cursor + 1
End Function

Function checkNext() As String
    If hasNext Then
        checkNext = Mid(ExpressionString, cursor, 1)
    Else
        'MsgBox "error"
    End If
End Function

Sub Reset()
    cursor = 1
End Sub
