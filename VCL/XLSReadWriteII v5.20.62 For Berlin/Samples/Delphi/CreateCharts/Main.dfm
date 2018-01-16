object frmMain: TfrmMain
  Left = 692
  Top = 247
  BorderStyle = bsDialog
  Caption = 'Create Chart Sample'
  ClientHeight = 506
  ClientWidth = 545
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 110
  TextHeight = 16
  object Label1: TLabel
    Left = 10
    Top = 15
    Width = 52
    Height = 16
    Caption = 'Filename'
  end
  object Label7: TLabel
    Left = 420
    Top = 40
    Width = 117
    Height = 101
    AutoSize = False
    Caption = 'Charts can only be craeted in Excel  97 (XLS) files.'
    WordWrap = True
  end
  object Button1: TButton
    Left = 448
    Top = 468
    Width = 92
    Height = 30
    Caption = 'Close'
    TabOrder = 0
    OnClick = Button1Click
  end
  object edFilename: TEdit
    Left = 69
    Top = 10
    Width = 425
    Height = 21
    TabOrder = 1
  end
  object Button3: TButton
    Left = 497
    Top = 10
    Width = 26
    Height = 26
    Caption = '...'
    TabOrder = 2
    OnClick = Button3Click
  end
  object Button4: TButton
    Left = 448
    Top = 428
    Width = 92
    Height = 31
    Caption = 'Save'
    TabOrder = 3
    OnClick = Button4Click
  end
  object PageControl1: TPageControl
    Left = 5
    Top = 44
    Width = 400
    Height = 454
    ActivePage = TabSheet1
    TabOrder = 4
    object TabSheet1: TTabSheet
      Caption = 'Basic'
      object Label2: TLabel
        Left = 10
        Top = 15
        Width = 61
        Height = 16
        Caption = 'Chart style'
      end
      object Title: TLabel
        Left = 197
        Top = 94
        Width = 25
        Height = 16
        Caption = 'Title'
      end
      object Label6: TLabel
        Left = 98
        Top = 251
        Width = 46
        Height = 16
        Caption = 'Markers'
      end
      object Label3: TLabel
        Left = 217
        Top = 0
        Width = 43
        Height = 16
        Caption = 'Options'
      end
      object Button2: TButton
        Left = 0
        Top = 143
        Width = 92
        Height = 31
        Caption = 'Column'
        TabOrder = 6
        OnClick = Button2Click
      end
      object cbLegend: TCheckBox
        Left = 197
        Top = 30
        Width = 100
        Height = 20
        Caption = 'Has legend'
        Checked = True
        State = cbChecked
        TabOrder = 0
      end
      object cb3D: TCheckBox
        Left = 197
        Top = 59
        Width = 105
        Height = 21
        Caption = '3D'
        TabOrder = 1
      end
      object edTitle: TEdit
        Left = 231
        Top = 89
        Width = 149
        Height = 21
        TabOrder = 2
      end
      object Button5: TButton
        Left = 0
        Top = 74
        Width = 92
        Height = 31
        Caption = 'Bar'
        TabOrder = 4
        OnClick = Button5Click
      end
      object Button6: TButton
        Left = 0
        Top = 39
        Width = 92
        Height = 31
        Caption = 'Area'
        TabOrder = 3
        OnClick = Button6Click
      end
      object Button7: TButton
        Left = 0
        Top = 108
        Width = 92
        Height = 31
        Caption = 'Bubble'
        TabOrder = 5
        OnClick = Button7Click
      end
      object Button8: TButton
        Left = 0
        Top = 177
        Width = 92
        Height = 31
        Caption = 'Cylinder'
        TabOrder = 7
        OnClick = Button8Click
      end
      object Button9: TButton
        Left = 0
        Top = 212
        Width = 92
        Height = 30
        Caption = 'Doughnut'
        TabOrder = 8
        OnClick = Button9Click
      end
      object Button10: TButton
        Left = 0
        Top = 246
        Width = 92
        Height = 31
        Caption = 'Line'
        TabOrder = 9
        OnClick = Button10Click
      end
      object Button11: TButton
        Left = 0
        Top = 281
        Width = 92
        Height = 30
        Caption = 'Pie'
        TabOrder = 10
        OnClick = Button11Click
      end
      object Button12: TButton
        Left = 0
        Top = 315
        Width = 92
        Height = 31
        Caption = 'Radar'
        TabOrder = 11
        OnClick = Button12Click
      end
      object Button13: TButton
        Left = 0
        Top = 350
        Width = 92
        Height = 30
        Caption = 'Scatter (XY)'
        TabOrder = 12
        OnClick = Button13Click
      end
      object Button14: TButton
        Left = 0
        Top = 384
        Width = 92
        Height = 31
        Caption = 'Surface'
        TabOrder = 14
        OnClick = Button14Click
      end
      object Button22: TButton
        Left = 98
        Top = 350
        Width = 194
        Height = 30
        Caption = 'Scatter (XY), multiple series'
        TabOrder = 13
        OnClick = Button22Click
      end
      object Button23: TButton
        Left = 197
        Top = 123
        Width = 92
        Height = 31
        Caption = 'Font A...'
        TabOrder = 15
        OnClick = Button23Click
      end
      object Button24: TButton
        Left = 197
        Top = 162
        Width = 92
        Height = 31
        Caption = 'Font B...'
        TabOrder = 16
        OnClick = Button24Click
      end
      object cbLineMarkers: TComboBox
        Left = 153
        Top = 246
        Width = 129
        Height = 24
        Style = csDropDownList
        ItemHeight = 16
        TabOrder = 17
        Items.Strings = (
          'None'
          'Square'
          'Diamond'
          'Triangle'
          'X'
          'Star'
          'Dow Jones'
          'Standard Deviation'
          'Circle'
          'Plus')
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Picture'
      ImageIndex = 1
      object Label4: TLabel
        Left = 5
        Top = 10
        Width = 60
        Height = 16
        Caption = 'Picture file'
      end
      object Button15: TButton
        Left = 5
        Top = 64
        Width = 92
        Height = 31
        Caption = 'Add chart'
        TabOrder = 0
        OnClick = Button15Click
      end
      object edPictFile: TEdit
        Left = 5
        Top = 30
        Width = 346
        Height = 21
        TabOrder = 1
        Text = 'Pig.jpg'
      end
      object Button16: TButton
        Left = 354
        Top = 30
        Width = 31
        Height = 25
        Caption = '...'
        TabOrder = 2
        OnClick = Button16Click
      end
      object Memo1: TMemo
        Left = 5
        Top = 103
        Width = 380
        Height = 110
        BorderStyle = bsNone
        Color = clBtnFace
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -15
        Font.Name = 'Courier New'
        Font.Style = []
        Lines.Strings = (
          'Creates a chart with the picture file as '
          'background in the chart and legend area.')
        ParentFont = False
        ReadOnly = True
        TabOrder = 3
      end
    end
    object TabSheet3: TTabSheet
      Caption = 'Gradient'
      ImageIndex = 2
      object shpColor1: TShape
        Left = 103
        Top = 15
        Width = 61
        Height = 21
      end
      object shpColor2: TShape
        Left = 103
        Top = 49
        Width = 61
        Height = 21
      end
      object Label5: TLabel
        Left = 79
        Top = 89
        Width = 78
        Height = 16
        Caption = 'Gradient style'
      end
      object Button17: TButton
        Left = 5
        Top = 123
        Width = 92
        Height = 31
        Caption = 'Add chart'
        TabOrder = 0
        OnClick = Button17Click
      end
      object Button18: TButton
        Left = 167
        Top = 10
        Width = 93
        Height = 31
        Caption = 'Color 1'
        TabOrder = 1
        OnClick = Button18Click
      end
      object Button19: TButton
        Left = 167
        Top = 44
        Width = 93
        Height = 31
        Caption = 'Color 2'
        TabOrder = 2
        OnClick = Button19Click
      end
      object cbGradientStyle: TComboBox
        Left = 167
        Top = 84
        Width = 179
        Height = 24
        Style = csDropDownList
        ItemHeight = 16
        TabOrder = 3
        Items.Strings = (
          'Horizontal'
          'Vertical'
          'Diagonal up'
          'Diagonal down'
          'From corner'
          'From center')
      end
      object Memo2: TMemo
        Left = 0
        Top = 167
        Width = 380
        Height = 110
        BorderStyle = bsNone
        Color = clBtnFace
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -15
        Font.Name = 'Courier New'
        Font.Style = []
        Lines.Strings = (
          'Creates a chart with gradient fill style '
          'in '
          'the chart and legend area.')
        ParentFont = False
        ReadOnly = True
        TabOrder = 4
      end
    end
    object TabSheet4: TTabSheet
      Caption = 'Color bar'
      ImageIndex = 3
      object Button20: TButton
        Left = 10
        Top = 10
        Width = 92
        Height = 31
        Caption = 'Add chart'
        TabOrder = 0
        OnClick = Button20Click
      end
      object Memo3: TMemo
        Left = 5
        Top = 103
        Width = 380
        Height = 110
        BorderStyle = bsNone
        Color = clBtnFace
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -15
        Font.Name = 'Courier New'
        Font.Style = []
        Lines.Strings = (
          'Creates a chart with randomly colored '
          'bars.')
        ParentFont = False
        ReadOnly = True
        TabOrder = 1
      end
    end
    object TabSheet5: TTabSheet
      Caption = 'Labels'
      ImageIndex = 4
      object Button21: TButton
        Left = 10
        Top = 10
        Width = 92
        Height = 31
        Caption = 'Add chart'
        TabOrder = 0
        OnClick = Button21Click
      end
      object Memo4: TMemo
        Left = 5
        Top = 103
        Width = 380
        Height = 110
        BorderStyle = bsNone
        Color = clBtnFace
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -15
        Font.Name = 'Courier New'
        Font.Style = []
        Lines.Strings = (
          'Creates a chart with labels (month names) '
          'attached to the bars.')
        ParentFont = False
        ReadOnly = True
        TabOrder = 1
      end
    end
  end
  object dlgSave: TSaveDialog
    Filter = 'Excel files (*.xls)|*.xls|All files (*.*)|*.*'
    Left = 404
    Top = 36
  end
  object dlgOpen: TOpenDialog
    Filter = 
      'Picture files (*.bmp;*.jpg;*.png;*.wmf)|*.bmp;*.jpg;*.png;*.wmf|' +
      'All files (*.*)|*.*'
    Left = 404
    Top = 68
  end
  object dlgFont: TFontDialog
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Left = 404
    Top = 100
  end
  object XLS: TXLSReadWriteII5
    ComponentVersion = '5.20.18'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 436
    Top = 164
  end
  object XPManifest1: TXPManifest
    Left = 436
    Top = 200
  end
end
