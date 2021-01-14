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
Rem  @description   MSForms�̃C�P�ĂȂ��R���g���[�����A�C�C�����Ɏg�����߂̊֐��Q
Rem
Rem  @update        2020/07/27
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Rem ���X�g�{�b�N�X�ɃA�C�e����ǉ�����
Rem
Rem
Rem  @note
Rem     �W����AddItem���\�b�h�͔z��ɑΉ����Ă��Ȃ����ߕK�v
Rem     �n���ꂽ�z��̗v�f�����A�\���\�ȗ񐔂𒴂��Ă��Ă��؂�̂Ă���
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
