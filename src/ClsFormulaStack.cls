VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsFormulaStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
