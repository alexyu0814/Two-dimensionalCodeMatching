object Form1: TForm1
  Left = 375
  Top = 124
  Caption = 'Form1'
  ClientHeight = 214
  ClientWidth = 444
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  OnKeyDown = FormKeyDown
  PixelsPerInch = 96
  TextHeight = 13
  object DBLookupComboboxEh1: TDBLookupComboboxEh
    Left = 80
    Top = 104
    Width = 265
    Height = 65
    ControlLabel.Width = 80
    ControlLabel.Height = 13
    ControlLabel.Caption = 'CheckComboBox'
    ControlLabel.Visible = True
    AutoSize = False
    DynProps = <>
    DataField = ''
    EditButton.DropDownFormParams.DropDownForm = DropDownForm2.Owner
    EditButtons = <>
    ReadOnly = True
    TabOrder = 0
    Visible = True
    WordWrap = True
    OnCloseDropDownForm = DBLookupComboboxEh1CloseDropDownForm
    OnOpenDropDownForm = DBLookupComboboxEh1OpenDropDownForm
  end
end
