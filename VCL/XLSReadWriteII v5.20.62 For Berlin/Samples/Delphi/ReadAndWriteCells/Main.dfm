object frmMain: TfrmMain
  Left = 736
  Top = 163
  Width = 691
  Height = 567
  Caption = 'Read and write cells'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 115
  TextHeight = 16
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 673
    Height = 119
    Align = alTop
    TabOrder = 0
    object btnRead: TButton
      Left = 10
      Top = 10
      Width = 92
      Height = 31
      Caption = 'Read'
      TabOrder = 0
      OnClick = btnReadClick
    end
    object btnWrite: TButton
      Left = 10
      Top = 44
      Width = 92
      Height = 31
      Caption = 'Write'
      TabOrder = 1
      OnClick = btnWriteClick
    end
    object edReadFilename: TEdit
      Left = 110
      Top = 12
      Width = 502
      Height = 21
      TabOrder = 2
    end
    object edWriteFilename: TEdit
      Left = 110
      Top = 46
      Width = 502
      Height = 21
      TabOrder = 3
    end
    object btnDlgOpen: TButton
      Left = 619
      Top = 10
      Width = 35
      Height = 31
      Caption = '...'
      TabOrder = 4
      OnClick = btnDlgOpenClick
    end
    object btnDlgSave: TButton
      Left = 619
      Top = 43
      Width = 35
      Height = 31
      Caption = '...'
      TabOrder = 5
      OnClick = btnDlgSaveClick
    end
    object Button1: TButton
      Left = 10
      Top = 79
      Width = 92
      Height = 31
      Caption = 'Close'
      TabOrder = 6
      OnClick = Button1Click
    end
    object btnAddCells: TButton
      Left = 108
      Top = 79
      Width = 93
      Height = 31
      Caption = 'Add cells'
      TabOrder = 7
      OnClick = btnAddCellsClick
    end
  end
  object Grid: TStringGrid
    Left = 0
    Top = 119
    Width = 673
    Height = 404
    Align = alClient
    ColCount = 4
    DefaultRowHeight = 16
    FixedCols = 0
    RowCount = 101
    TabOrder = 1
    ColWidths = (
      64
      65
      163
      261)
  end
  object dlgSave: TSaveDialog
    Filter = 'Excel files |*.xlsx;*.xlsm;*.xls;*.xlm|All files (*.*)|*.*'
    Left = 132
    Top = 100
  end
  object dlgOpen: TOpenDialog
    Filter = 'Excel files|*.xlsx;*.xls|All files (*.*)|*.*'
    Left = 72
    Top = 100
  end
  object XLS: TXLSReadWriteII5
    ComponentVersion = '5.20.12'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 100
    Top = 148
  end
  object XPManifest1: TXPManifest
    Left = 184
    Top = 152
  end
end
