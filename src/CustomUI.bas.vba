Attribute VB_Name = "CustomUI"
Rem
Rem CustomUI
Rem
Rem �{���W���[���͎����CustomUI�G�f�B�^���玩�����������C�x���g�n���h���ł��B
Rem

Sub onAction_AddinStart(control As IRibbonControl): Call AddinStart: FinalUseCommand = "AddinStart": End Sub
Sub onAction_AddinStop(control As IRibbonControl): Call AddinStop: FinalUseCommand = "AddinStop": End Sub
Sub onAction_AddinConfig(control As IRibbonControl): Call AddinConfig: FinalUseCommand = "AddinConfig": End Sub
Sub onAction_AddinInfo(control As IRibbonControl): Call AddinInfo: FinalUseCommand = "AddinInfo": End Sub
Sub onAction_AddinEnd(control As IRibbonControl): Call AddinEnd: FinalUseCommand = "AddinEnd": End Sub

Sub onAction_OpenFormulaEditorForm(control As IRibbonControl): Call OpenFormulaEditorForm: FinalUseCommand = "OpenFormulaEditorForm": End Sub
