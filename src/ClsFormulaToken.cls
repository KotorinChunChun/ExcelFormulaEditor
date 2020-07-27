VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsFormulaToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public tString

'•¶Žš—ñ‚ðŒã‚ë‚É’Ç‹L
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


