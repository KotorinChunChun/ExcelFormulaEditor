VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsFormulaToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        ClsFormulaToken
Rem
Rem  @description   Tokenクラス
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

Public tString

'文字列を後ろに追記
Sub AddChar(C)
    tString = tString & C
End Sub

Property Get tType() As TokenType
    Dim ret As TokenType
    Select Case UCase(tString)
        Case "IF": ret = TokenType.Target
        Case "(": ret = TokenType.BeginParen
        Case ")": ret = TokenType.EndParen
        Case ",": ret = TokenType.Comma
        Case vbLf: ret = TokenType.LineFeed
        Case "!": ret = TokenType.Exclamation
        Case Else
            If Left(tString, 1) = "#" And Right(tString, 1) = "!" Then
                ret = TokenType.ErrCode
            Else
                ret = TokenType.Other
            End If
    End Select
    tType = ret
End Property
