VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kccFuncMSForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccFuncMSForms
Rem
Rem  @description   MSFormsのイケてないコントロールを、イイ感じに使うための関数群
Rem
Rem  @update        2020/07/27
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Rem リストボックスにアイテムを追加する
Rem
Rem
Rem  @note
Rem     標準のAddItemメソッドは配列に対応していないため必要
Rem     渡された配列の要素数が、表示可能な列数を超えていても切り捨てられる
Rem
Public Function ListBox_AddItem(lb As MSForms.ListBox, insertRowData, Optional ByVal insertRowIndex As Long = -1) As Long
    If insertRowIndex = -1 Then
        insertRowIndex = lb.ListCount
    End If
    
    If Not IsArray(insertRowData) Then
        lb.addItem insertRowData, insertRowIndex
        ListBox_AddItem = insertRowIndex
        Exit Function
    End If
    
    lb.addItem "", insertRowIndex
    Dim columnIndex As Long, itemIndex As Long
    itemIndex = LBound(insertRowData)
    For columnIndex = 0 To lb.ColumnCount - 1
        lb.List(insertRowIndex, columnIndex) = insertRowData(itemIndex)
        itemIndex = itemIndex + 1
    Next
    ListBox_AddItem = insertRowIndex
End Function
