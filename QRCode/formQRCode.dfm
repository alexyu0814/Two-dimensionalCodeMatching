object frmQRCode: TfrmQRCode
  Left = 0
  Top = 0
  Caption = #28895#21378#21152#24037#21830#20108#32500#30721#21435#37325#12289#21305#37197#20108#32500#30721#36719#20214' V1.0'
  ClientHeight = 637
  ClientWidth = 944
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 944
    Height = 637
    ActivePage = TabMulti
    Align = alClient
    TabOrder = 0
    object TabSheet1: TTabSheet
      Caption = #20027#30028#38754'('#20108#32500#30721')'
      TabVisible = False
      object lblResult: TLabel
        Left = 0
        Top = 596
        Width = 936
        Height = 13
        Align = alBottom
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        ExplicitWidth = 3
      end
      object Label29: TLabel
        Left = 0
        Top = 321
        Width = 537
        Height = 275
        Align = alLeft
        AutoSize = False
        Caption = 
          #29983#25104#20108#32500#30721#35828#26126#65306#13#10'1'#12289#35831#20808#22312'OA'#30003#35831#8220#20108#32500#30721#30003#35831#8221#27969#31243#65292#20877#22797#21046'OA'#30340#20027#34920#21333#25454#21495#13#10'2'#12289#29983#25104#20108#32500#30721#20043#21069#35831#20180#32454#26680#23545#24037#21333#32534#21495#12289#21378#23478#31867 +
          #22411#12289#32593#22336#31561#20449#24687#13#10'3'#12289#22914#20108#32500#30721#25968#37327#22823#20110#21333#20010#25968#25454#21253#19978#38480#65292#21017#20250#36827#34892#25286#21253#65292#22914#65306#30003#35831#30340#20108#32500#30721#25968#37327#20026'120'#19975#65292#21333#20010#25968#25454#21253#19978#38480#20026'50'#19975#65292#21017#20250 +
          #25286#25104'3'#20010#21253#65292#20998#21035#20026'50'#19975#12289'50'#19975#12289'20'#19975
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clPurple
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        WordWrap = True
        ExplicitLeft = -4
        ExplicitTop = 289
        ExplicitHeight = 177
      end
      object btnStart: TButton
        Left = 537
        Top = 321
        Width = 399
        Height = 275
        Align = alClient
        Caption = #24320#22987#29983#25104#20108#32500#30721
        TabOrder = 1
        OnClick = btnStartClick
      end
      object Panel1: TPanel
        Left = 0
        Top = 0
        Width = 936
        Height = 321
        Align = alTop
        TabOrder = 0
        object Label1: TLabel
          Left = 17
          Top = 170
          Width = 111
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #21697#29260#20195#30721#65288'2'#20301#65289#65306
        end
        object Label2: TLabel
          Left = 351
          Top = 170
          Width = 296
          Height = 13
          AutoSize = False
          Caption = '01'#65306#21452#21916'('#32463#20856#24037#22346') 02'#65306#21452#21916'('#24425#24742') 03'#65306#21452#21916'('#21916#30334#24180')'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object Label3: TLabel
          Left = 2
          Top = 200
          Width = 126
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #21360#21047#21378#20195#30721#65288'2'#20301#65289#65306
        end
        object Label4: TLabel
          Left = 351
          Top = 200
          Width = 288
          Height = 13
          AutoSize = False
          Caption = '02'#65306#21170#22025
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object Label5: TLabel
          Left = 49
          Top = 228
          Width = 79
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #24180#26376#65288'4'#20301#65289#65306
        end
        object Label6: TLabel
          Left = 49
          Top = 6
          Width = 79
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #20027#34920#21333#25454#21495#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label7: TLabel
          Left = 49
          Top = 143
          Width = 79
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #32593#22336#65306
        end
        object Label8: TLabel
          Left = 49
          Top = 259
          Width = 79
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #20108#32500#30721#25968#37327#65306
        end
        object Label9: TLabel
          Left = 351
          Top = 6
          Width = 208
          Height = 13
          AutoSize = False
          Caption = 'OA'#21333#25454#21495#65306#24517#39035#20197'QR'#24320#22836
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label10: TLabel
          Left = 351
          Top = 123
          Width = 111
          Height = 13
          AutoSize = False
          Caption = '0'#65306#24191#19996#20013#28895
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object Label11: TLabel
          Left = 49
          Top = 119
          Width = 79
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #21378#23478#31867#22411#65306
        end
        object Label12: TLabel
          Left = 351
          Top = 228
          Width = 104
          Height = 13
          AutoSize = False
          Caption = #26684#24335#65306'YYMM'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label13: TLabel
          Left = 49
          Top = 286
          Width = 79
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #24037#20316#32447#31243#25968#65306
          Visible = False
        end
        object Label14: TLabel
          Left = 351
          Top = 283
          Width = 208
          Height = 13
          AutoSize = False
          Caption = #21407#21017#26159'10'#19975#25968#25454#21253#24320#36767#19968#20010#32447#31243
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          Visible = False
        end
        object Label16: TLabel
          Left = 49
          Top = 89
          Width = 79
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #26465#30418'/'#23567#30418#65306
        end
        object Label20: TLabel
          Left = 351
          Top = 258
          Width = 112
          Height = 13
          AutoSize = False
          Caption = #21333#20010#25968#25454#21253#19978#38480#65306
        end
        object Label21: TLabel
          Left = 49
          Top = 33
          Width = 79
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #24037#21333#32534#21495#65306
        end
        object Label22: TLabel
          Left = 49
          Top = 59
          Width = 79
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #21360#21697#32534#21495#65306
        end
        object edtNo: TEdit
          Left = 136
          Top = 3
          Width = 201
          Height = 21
          TabOrder = 0
          OnExit = edtNoExit
        end
        object edtMode: TEdit
          Left = 136
          Top = 113
          Width = 201
          Height = 21
          ReadOnly = True
          TabOrder = 5
          Text = '0'
        end
        object edtWebAddr: TEdit
          Left = 136
          Top = 140
          Width = 201
          Height = 21
          ReadOnly = True
          TabOrder = 6
          Text = 'https://hd.shuangxi100.com/'
        end
        object edtBrandCode: TEdit
          Left = 136
          Top = 167
          Width = 201
          Height = 21
          ReadOnly = True
          TabOrder = 7
          Text = '01'
        end
        object edtPrinterCode: TEdit
          Left = 136
          Top = 194
          Width = 201
          Height = 21
          ReadOnly = True
          TabOrder = 8
          Text = '02'
        end
        object edtNy: TEdit
          Left = 136
          Top = 222
          Width = 201
          Height = 21
          ReadOnly = True
          TabOrder = 9
        end
        object edtCount: TEdit
          Left = 136
          Top = 251
          Width = 201
          Height = 21
          ReadOnly = True
          TabOrder = 10
        end
        object edtThreadCount: TEdit
          Left = 136
          Top = 278
          Width = 201
          Height = 21
          ReadOnly = True
          TabOrder = 12
          Text = '1'
          Visible = False
        end
        object edtType: TComboBox
          Left = 136
          Top = 83
          Width = 201
          Height = 21
          Style = csDropDownList
          TabOrder = 4
          Items.Strings = (
            #26465#30418
            #23567#30418)
        end
        object edtMax: TMaskEdit
          Left = 448
          Top = 255
          Width = 113
          Height = 21
          EditMask = '!999999;1;_'
          MaxLength = 6
          TabOrder = 11
          Text = '500000'
        end
        object edtGdbh: TEdit
          Left = 136
          Top = 30
          Width = 203
          Height = 21
          ReadOnly = True
          TabOrder = 1
          OnExit = edtNoExit
        end
        object edtYpbh: TEdit
          Left = 136
          Top = 56
          Width = 113
          Height = 21
          ReadOnly = True
          TabOrder = 2
          OnExit = edtNoExit
        end
        object edtYpmc: TEdit
          Left = 248
          Top = 56
          Width = 313
          Height = 21
          ReadOnly = True
          TabOrder = 3
          OnExit = edtNoExit
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = #21305#37197#21382#21490#24211'('#20108#32500#30721')'
      ImageIndex = 1
      TabVisible = False
      object Label28: TLabel
        Left = 0
        Top = 283
        Width = 537
        Height = 326
        Align = alLeft
        AutoSize = False
        Caption = 
          #21305#37197#21382#21490#24211#35268#21017#35828#26126#65306#13#10#22914#26524#21516#19968#25209#20108#32500#30721#36229#36807'50'#19975#65292#20026'120'#19975#65292#21017#20250#25286#20998#25104#19977#20010#21253#65292#20551#35774#21253#21517#20026#65306'QR00000001_1(50'#19975')' +
          #12289'QR00000001_2(50'#19975')'#12289'QR00000001_3(20'#19975')'#13#10'1'#12289#21305#37197'QR00000001_1(50'#19975')'#33258#36523#65292#33509#21305 +
          #37197#21040#37325#22797#30721#65292#21017#21024#25481'QR00000001_1'#21253#30340#37325#22797#30721#13#10'2'#12289#21305#37197'QR00000001_2(50'#19975')'#33258#36523#65292#33509#21305#37197#21040#37325#22797#30721#65292#21017#21024#25481'Q' +
          'R00000001_2'#21253#30340#37325#22797#30721#13#10'3'#12289#21305#37197'QR00000001_3(20'#19975')'#33258#36523#65292#33509#21305#37197#21040#37325#22797#30721#65292#21017#21024#25481'QR00000001_' +
          '3'#21253#30340#37325#22797#30721#13#10#21382#21490#24211#35828#26126#65306#21516#19968#24180#26376#26085#19979#26377#29983#25104#36807#20108#32500#30721#25968#25454#21253#65292#20551#35774#21253#21517#20026#65306'QR00000011_1'#12289'QR00000022_1'#13#10 +
          '1'#12289#21305#37197'QR00000001_1(50'#19975')'#21644'QR00000011_1'#65292#33509#21305#37197#21040#37325#22797#30721#65292#21017#21024#25481'QR00000001_1'#21253#30340#37325#22797#30721#13 +
          #10'2'#12289#21305#37197'QR00000001_1(50'#19975')'#21644'QR00000022_1'#65292#33509#21305#37197#21040#37325#22797#30721#65292#21017#21024#25481'QR00000001_1'#21253#30340#37325#22797#30721 +
          #13#10'3'#12289#21305#37197'QR00000001_2(50'#19975')'#21644'QR00000011_1'#65292#33509#21305#37197#21040#37325#22797#30721#65292#21017#21024#25481'QR00000001_2'#21253#30340#37325#22797 +
          #30721#13#10'4'#12289#21305#37197'QR00000001_2(50'#19975')'#21644'QR00000022_1'#65292#33509#21305#37197#21040#37325#22797#30721#65292#21017#21024#25481'QR00000001_2'#21253#30340#37325 +
          #22797#30721#13#10'5'#12289#21305#37197'QR00000001_3(20'#19975')'#21644'QR00000011_1'#65292#33509#21305#37197#21040#37325#22797#30721#65292#21017#21024#25481'QR00000001_3'#21253#30340 +
          #37325#22797#30721#13#10'6'#12289#21305#37197'QR00000001_3(20'#19975')'#21644'QR00000022_1'#65292#33509#21305#37197#21040#37325#22797#30721#65292#21017#21024#25481'QR00000001_3'#21253 +
          #30340#37325#22797#30721
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clPurple
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        WordWrap = True
        ExplicitLeft = -4
        ExplicitTop = 289
        ExplicitHeight = 177
      end
      object btnMatch: TButton
        Left = 797
        Top = 283
        Width = 139
        Height = 326
        Align = alRight
        Caption = #24320#22987#21305#37197#37325#22797#30721
        TabOrder = 1
        OnClick = btnMatchClick
      end
      object Memo1: TMemo
        Left = 0
        Top = 0
        Width = 936
        Height = 283
        Align = alTop
        ReadOnly = True
        ScrollBars = ssBoth
        TabOrder = 0
      end
    end
    object TabSheet3: TTabSheet
      Caption = #29983#25104#25991#20214'('#20108#32500#30721')'
      ImageIndex = 2
      TabVisible = False
      object Label30: TLabel
        Left = 0
        Top = 329
        Width = 537
        Height = 280
        Align = alLeft
        AutoSize = False
        Caption = 
          #29983#25104#25991#20214#35828#26126#65306#13#10'1'#12289#28857#20987#8220#24320#22987#29983#25104#25991#20214#8221#21518#65292#20250#29983#25104#20027#30028#38754'('#20108#32500#30721')'#30340#20027#34920#21333#25454#21495#19979#30340#25152#26377#21253#30340#25991#20214#65292#22914#20027#34920#21333#25454#21495#20026#65306'QR00000' +
          '001'#65292#25968#37327#20026'120'#19975#65292#23545#24212#26377'3'#20010#21253#65306#20998#21035#20026'50'#19975#12289'50'#19975#12289'20'#19975#65292#21017#20250#29983#25104'3'#20010#25991#20214#13#10'2'#12289#29983#25104#30340#25991#20214#35831#22312#8220#29983#25104#25991#20214#30340#36335#24452#8221#19979#26597 +
          #25214#65292#29983#25104#25991#20214#30340#36335#24452#21487#20197#20462#25913#13#10'3'#12289#29983#25104#30340#25991#20214#21517#30340#26684#24335#20026#65306#20027#34920#21333#25454#21495'('#24037#21333#21495#12304#12305','#21360#21697#32534#21495#12304#12305','#21360#21697#21517#31216#12304#12305','#30418#22411#12304#12305','#20108#32500#30721#25968#37327 +
          #12304#12305')'#13#10'4'#12289#29983#25104#30340#25991#20214#20869#23481#30340#26684#24335#20026#65306'27'#20301#32593#22336'+2'#20301#21697#29260#20195#30721'+2'#20301#21360#21047#21378#20195#30721'+4'#20301#24180#26376'+20'#20301#38543#26426#30721'+6'#20301#26657#39564#30721#65292#24635#20849#20026'61'#20301
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clPurple
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        WordWrap = True
        ExplicitTop = 326
        ExplicitHeight = 134
      end
      object btnCreateFile: TButton
        Left = 537
        Top = 329
        Width = 399
        Height = 280
        Align = alClient
        Caption = #24320#22987#29983#25104#25991#20214
        TabOrder = 2
        OnClick = btnCreateFileClick
      end
      object Memo2: TMemo
        Left = 0
        Top = 0
        Width = 936
        Height = 295
        Align = alTop
        ReadOnly = True
        ScrollBars = ssBoth
        TabOrder = 0
      end
      object Panel2: TPanel
        Left = 0
        Top = 295
        Width = 936
        Height = 34
        Align = alTop
        TabOrder = 1
        object Label15: TLabel
          Left = 50
          Top = 12
          Width = 114
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #29983#25104#25991#20214#30340#36335#24452#65306
        end
        object edtPath: TEdit
          Left = 170
          Top = 4
          Width = 203
          Height = 21
          ReadOnly = True
          TabOrder = 0
          Text = 'E:\QRCode'
        end
        object CheckBox1: TCheckBox
          Left = 396
          Top = 6
          Width = 137
          Height = 17
          Alignment = taLeftJustify
          Caption = #29983#25104#30340#25991#20214#26159#21542#24102#32593#22336
          Checked = True
          State = cbChecked
          TabOrder = 1
        end
      end
    end
    object TabSheet5: TTabSheet
      Caption = #23548#20837#23458#25143#25991#20214
      ImageIndex = 4
      object grp1: TGroupBox
        Left = 0
        Top = 42
        Width = 936
        Height = 233
        Align = alTop
        Caption = #25991#20214#21015#34920
        TabOrder = 0
        object pnl1: TPanel
          Left = 2
          Top = 190
          Width = 932
          Height = 41
          Align = alBottom
          TabOrder = 1
          object lbl2: TLabel
            Left = 30
            Top = 15
            Width = 131
            Height = 13
            AutoSize = False
            Caption = #20165#25903#25345'xls'#12289'xlsx'#12289'txt'#25991#20214
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clRed
            Font.Height = -11
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object btnImportCustFile: TButton
            Left = 428
            Top = 12
            Width = 80
            Height = 25
            Caption = #23548#20837#21644#23545#27604
            TabOrder = 3
            OnClick = btnImportCustFileClick
          end
          object btnAddCustFile: TButton
            Left = 190
            Top = 12
            Width = 80
            Height = 25
            Caption = #22686#21152
            TabOrder = 0
            OnClick = btnAddCustFileClick
          end
          object btnDelCustFile: TButton
            Left = 269
            Top = 12
            Width = 80
            Height = 25
            Caption = #21024#38500
            TabOrder = 1
            OnClick = btnDelCustFileClick
          end
          object btnClearCustFile: TButton
            Left = 349
            Top = 12
            Width = 80
            Height = 25
            Caption = #28165#31354
            TabOrder = 2
            OnClick = btnClearCustFileClick
          end
          object Button2: TButton
            Left = 514
            Top = 12
            Width = 103
            Height = 25
            Caption = #21305#37197#20108#32500#30721
            TabOrder = 4
            OnClick = Button2Click
          end
        end
        object gridCustFile: TDBGridEh
          Left = 2
          Top = 15
          Width = 932
          Height = 175
          Align = alClient
          DataSource = dsCustFile
          DynProps = <>
          TabOrder = 0
          OnMouseDown = gridCustFileMouseDown
          object RowDetailData: TRowDetailPanelControlEh
          end
        end
      end
      object grp2: TGroupBox
        Left = 0
        Top = 275
        Width = 936
        Height = 334
        Align = alClient
        Caption = #25805#20316#26085#24535
        TabOrder = 1
        object mmoCustFile: TMemo
          Left = 2
          Top = 15
          Width = 932
          Height = 298
          Align = alClient
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          ReadOnly = True
          ScrollBars = ssBoth
          TabOrder = 0
        end
        object StatusBar1: TStatusBar
          Left = 2
          Top = 313
          Width = 932
          Height = 19
          Panels = <
            item
              Width = 500
            end
            item
              Width = 50
            end>
        end
        object ProgressBar1: TProgressBar
          Left = 584
          Top = 272
          Width = 150
          Height = 17
          TabOrder = 2
        end
      end
      object GroupBox8: TGroupBox
        Left = 0
        Top = 0
        Width = 936
        Height = 42
        Align = alTop
        Caption = #25991#20214#21015#34920
        TabOrder = 2
        object Label31: TLabel
          Left = 14
          Top = 15
          Width = 48
          Height = 13
          Caption = #21697#26816#25991#20214
        end
        object edtTxt: TEdit
          Left = 70
          Top = 11
          Width = 549
          Height = 21
          TabOrder = 0
        end
        object btnTxt: TButton
          Left = 622
          Top = 9
          Width = 67
          Height = 25
          Caption = #27983#35272
          TabOrder = 1
          OnClick = btnTxtClick
        end
      end
    end
    object TabSheet4: TTabSheet
      Caption = #30456#21516'/'#24046#24322
      ImageIndex = 3
      object GroupBox1: TGroupBox
        Left = 0
        Top = 0
        Width = 936
        Height = 85
        Align = alTop
        Caption = #23545#27604#25991#20214
        TabOrder = 0
        object Label23: TLabel
          Left = 24
          Top = 25
          Width = 39
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = 'A'#25991#20214
        end
        object Label24: TLabel
          Left = 24
          Top = 56
          Width = 39
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = 'B'#25991#20214
        end
        object edtA: TEdit
          Left = 69
          Top = 21
          Width = 276
          Height = 21
          ReadOnly = True
          TabOrder = 2
        end
        object edtB: TEdit
          Left = 69
          Top = 52
          Width = 276
          Height = 21
          ReadOnly = True
          TabOrder = 5
        end
        object btnImportA: TButton
          Left = 351
          Top = 17
          Width = 75
          Height = 25
          Caption = #23548#20837#25991#20214
          TabOrder = 0
          OnClick = btnImportAClick
        end
        object btnImportB: TButton
          Left = 351
          Top = 48
          Width = 75
          Height = 25
          Caption = #23548#20837#25991#20214
          TabOrder = 3
          OnClick = btnImportBClick
        end
        object btnClearA: TButton
          Left = 432
          Top = 17
          Width = 41
          Height = 25
          Caption = #28165#31354
          TabOrder = 1
          OnClick = btnClearAClick
        end
        object btnClearB: TButton
          Left = 432
          Top = 48
          Width = 41
          Height = 25
          Caption = #28165#31354
          TabOrder = 4
          OnClick = btnClearBClick
        end
      end
      object Panel3: TPanel
        Left = 0
        Top = 85
        Width = 936
        Height = 42
        Align = alTop
        TabOrder = 1
        object btnPublic: TButton
          Left = 287
          Top = 6
          Width = 217
          Height = 25
          Caption = #24471#21040'A'#25991#20214#20013#21547#26377'B'#25991#20214#30340#20849#26377#25968#25454
          TabOrder = 1
          OnClick = btnPublicClick
        end
        object btnRepeat: TButton
          Left = 64
          Top = 6
          Width = 217
          Height = 25
          Caption = #21435#38500'A'#25991#20214#20013#21547#26377'B'#25991#20214#30340#37325#22797#25968#25454
          TabOrder = 0
          OnClick = btnRepeatClick
        end
      end
      object GroupBox2: TGroupBox
        Left = 0
        Top = 127
        Width = 936
        Height = 482
        Align = alClient
        Caption = #25805#20316#26085#24535
        TabOrder = 2
        object mmoLog: TMemo
          Left = 2
          Top = 15
          Width = 932
          Height = 465
          Align = alClient
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          ReadOnly = True
          ScrollBars = ssBoth
          TabOrder = 0
        end
      end
    end
    object TabMulti: TTabSheet
      Caption = #22810#25991#20214#21435#37325#22797
      ImageIndex = 6
      object grp4: TGroupBox
        Left = 0
        Top = 0
        Width = 936
        Height = 233
        Align = alTop
        Caption = #25991#20214#21015#34920
        TabOrder = 0
        object pnl2: TPanel
          Left = 2
          Top = 176
          Width = 932
          Height = 55
          Align = alBottom
          TabOrder = 1
          object lbl7: TLabel
            Left = 16
            Top = 40
            Width = 385
            Height = 13
            AutoSize = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clRed
            Font.Height = -11
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object lbl8: TLabel
            Left = 160
            Top = 38
            Width = 107
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #20108#32500#30721#20301#25968
            Visible = False
          end
          object lblFileCount: TLabel
            Left = 445
            Top = 40
            Width = 334
            Height = 13
            AutoSize = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlue
            Font.Height = -11
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object btnImportMulti: TButton
            Left = 341
            Top = 9
            Width = 82
            Height = 25
            Caption = #23548#20837
            TabOrder = 4
            OnClick = btnImportMultiClick
          end
          object btnAddMulti: TButton
            Left = 14
            Top = 9
            Width = 82
            Height = 25
            Caption = #28155#21152#25991#20214
            TabOrder = 0
            OnClick = btnAddMultiClick
          end
          object btnDelMulti: TButton
            Left = 177
            Top = 9
            Width = 82
            Height = 25
            Caption = #21024#38500
            TabOrder = 2
            OnClick = btnDelMultiClick
          end
          object btnClearMulti: TButton
            Left = 259
            Top = 9
            Width = 82
            Height = 25
            Caption = #28165#31354
            TabOrder = 3
            OnClick = btnClearMultiClick
          end
          object btnRepeatMulti: TButton
            Left = 424
            Top = 9
            Width = 82
            Height = 25
            Caption = #22788#29702#37325#22797
            TabOrder = 5
            OnClick = btnRepeatMultiClick
          end
          object btnNotRepeatMulti: TButton
            Left = 505
            Top = 9
            Width = 82
            Height = 25
            Caption = #22788#29702#19981#37325#22797
            TabOrder = 6
            OnClick = btnNotRepeatMultiClick
          end
          object btnAddMultiDir: TButton
            Left = 95
            Top = 9
            Width = 82
            Height = 25
            Caption = #28155#21152#25991#20214#22841
            TabOrder = 1
            OnClick = btnAddMultiDirClick
          end
          object edtCodeLen: TEdit
            Left = 273
            Top = 30
            Width = 48
            Height = 21
            ReadOnly = True
            TabOrder = 7
            Text = '100'
            Visible = False
          end
          object btnImportMulti2: TButton
            Left = 593
            Top = 9
            Width = 186
            Height = 25
            Caption = #23548#20837#25253#38169#21333#20987#36825#37324'('#34917#22238#36710#31526#21495')'
            TabOrder = 8
            OnClick = btnImportMulti2Click
          end
        end
        object gridMulti: TDBGridEh
          Left = 2
          Top = 15
          Width = 932
          Height = 161
          Align = alClient
          DataSource = dsMulti
          DynProps = <>
          TabOrder = 0
          OnMouseDown = gridMultiMouseDown
          object RowDetailData: TRowDetailPanelControlEh
          end
        end
      end
      object grp5: TGroupBox
        Left = 0
        Top = 233
        Width = 936
        Height = 376
        Align = alClient
        Caption = #25805#20316#26085#24535
        TabOrder = 1
        object mmoMulti: TMemo
          Left = 2
          Top = 15
          Width = 932
          Height = 359
          Align = alClient
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          ReadOnly = True
          ScrollBars = ssBoth
          TabOrder = 0
        end
      end
    end
    object TabCheck: TTabSheet
      Caption = #21305#37197#39564#35777#30721
      ImageIndex = 5
      object grp3: TGroupBox
        Left = 0
        Top = 0
        Width = 936
        Height = 89
        Align = alTop
        Caption = #23545#27604#25991#20214
        TabOrder = 0
        object lbl1: TLabel
          Left = 22
          Top = 29
          Width = 39
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #23376#25991#20214
        end
        object lbl3: TLabel
          Left = 24
          Top = 55
          Width = 39
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #29238#25991#20214
        end
        object Label25: TLabel
          Left = 22
          Top = 57
          Width = 102
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #23458#25143#25991#20214#23548#20837#24180#26376
          Visible = False
        end
        object Label27: TLabel
          Left = 207
          Top = 57
          Width = 150
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #26684#24335#65306'YYYYMM'#65288#22914'201612'#65289
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          Visible = False
        end
        object edtCheckA: TEdit
          Left = 69
          Top = 21
          Width = 276
          Height = 21
          ReadOnly = True
          TabOrder = 2
        end
        object edtCheckB: TEdit
          Left = 69
          Top = 62
          Width = 276
          Height = 21
          NumbersOnly = True
          ReadOnly = True
          TabOrder = 6
        end
        object btnImportCheckA: TButton
          Left = 351
          Top = 17
          Width = 75
          Height = 25
          Caption = #23548#20837#25991#20214
          TabOrder = 0
          OnClick = btnImportCheckAClick
        end
        object btnImportCheckB: TButton
          Left = 351
          Top = 48
          Width = 75
          Height = 25
          Caption = #23548#20837#25991#20214
          TabOrder = 4
          OnClick = btnImportCheckBClick
        end
        object btnClearCheckA: TButton
          Left = 432
          Top = 17
          Width = 41
          Height = 25
          Caption = #28165#31354
          TabOrder = 1
          OnClick = btnClearCheckAClick
        end
        object btnClearCheckB: TButton
          Left = 432
          Top = 48
          Width = 41
          Height = 25
          Caption = #28165#31354
          TabOrder = 5
          OnClick = btnClearCheckBClick
        end
        object chk1: TCheckBox
          Left = 479
          Top = 56
          Width = 146
          Height = 17
          Caption = #39318#34892#26159#21542#21253#21547#21015#26631#39064
          Checked = True
          State = cbChecked
          TabOrder = 8
          Visible = False
        end
        object chkAuto: TCheckBox
          Left = 479
          Top = 21
          Width = 114
          Height = 17
          Caption = #33258#21160#21305#37197#29238#25991#20214
          TabOrder = 3
          OnClick = chkAutoClick
        end
        object edtCustNy: TEdit
          Left = 127
          Top = 52
          Width = 74
          Height = 21
          TabOrder = 7
          Visible = False
        end
      end
      object Panel4: TPanel
        Left = 0
        Top = 395
        Width = 936
        Height = 82
        Align = alBottom
        TabOrder = 3
        object lbl6: TLabel
          Left = 238
          Top = 34
          Width = 50
          Height = 13
          Alignment = taRightJustify
          AutoSize = False
          Caption = #23376#25209#27425#21495
        end
        object chk3: TCheckBox
          Left = 489
          Top = 28
          Width = 97
          Height = 17
          Caption = #29983#25104#34920#26684
          TabOrder = 3
          OnClick = chk3Click
        end
        object chk2: TCheckBox
          Left = 400
          Top = 28
          Width = 75
          Height = 17
          Caption = #29983#25104#25991#26412
          Checked = True
          State = cbChecked
          TabOrder = 2
          OnClick = chk2Click
        end
        object edtBatchNo: TEdit
          Left = 294
          Top = 28
          Width = 100
          Height = 21
          TabOrder = 1
        end
        object btnCheck: TButton
          Left = 70
          Top = 20
          Width = 131
          Height = 44
          Caption = #21305#37197#39564#35777#30721
          TabOrder = 0
          OnClick = btnCheckClick
        end
      end
      object GroupBox3: TGroupBox
        Left = 0
        Top = 477
        Width = 936
        Height = 132
        Align = alBottom
        Caption = #25805#20316#26085#24535
        TabOrder = 4
        object mmoLog1: TMemo
          Left = 2
          Top = 15
          Width = 932
          Height = 115
          Align = alClient
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          ReadOnly = True
          ScrollBars = ssBoth
          TabOrder = 0
        end
      end
      object GroupBox4: TGroupBox
        Left = 0
        Top = 89
        Width = 329
        Height = 306
        Align = alLeft
        Caption = #23376#25991#20214#26684#24335
        TabOrder = 1
        object CheckListBoxA: TCheckListBox
          Left = 2
          Top = 15
          Width = 286
          Height = 289
          OnClickCheck = CheckListBoxAClickCheck
          Align = alClient
          ItemHeight = 13
          Items.Strings = (
            #24207#21495
            #26465#30721#20869#23481
            #39564#35777#30721
            #23376#25209#27425#21495)
          TabOrder = 0
        end
        object Panel5: TPanel
          Left = 288
          Top = 15
          Width = 39
          Height = 289
          Align = alRight
          TabOrder = 1
          object btnUpA: TSpeedButton
            Left = 3
            Top = 27
            Width = 31
            Height = 33
            Caption = #8593
            OnClick = btnUpAClick
          end
          object btnDownA: TSpeedButton
            Left = 3
            Top = 66
            Width = 31
            Height = 33
            Caption = #8595
            OnClick = btnDownAClick
          end
        end
      end
      object GroupBox5: TGroupBox
        Left = 538
        Top = 89
        Width = 398
        Height = 306
        Align = alRight
        Caption = #29238#25991#20214#26684#24335
        TabOrder = 2
        object CheckListBoxB: TCheckListBox
          Left = 2
          Top = 15
          Width = 355
          Height = 289
          OnClickCheck = CheckListBoxBClickCheck
          Align = alClient
          ItemHeight = 13
          Items.Strings = (
            #24207#21495
            #26465#30721#20869#23481
            #39564#35777#30721
            #23376#25209#27425#21495)
          TabOrder = 0
        end
        object Panel6: TPanel
          Left = 357
          Top = 15
          Width = 39
          Height = 289
          Align = alRight
          TabOrder = 1
          object btnUpB: TSpeedButton
            Left = 3
            Top = 27
            Width = 31
            Height = 33
            Caption = #8593
            OnClick = btnUpBClick
          end
          object btnDownB: TSpeedButton
            Left = 3
            Top = 66
            Width = 31
            Height = 33
            Caption = #8595
            OnClick = btnDownBClick
          end
        end
      end
    end
    object tabServer: TTabSheet
      Caption = #26381#21153#22120#37197#32622
      ImageIndex = 7
      TabVisible = False
      object Label17: TLabel
        Left = 28
        Top = 28
        Width = 114
        Height = 13
        Alignment = taRightJustify
        AutoSize = False
        Caption = #26381#21153#22120'IP'#65306
      end
      object Label18: TLabel
        Left = 28
        Top = 55
        Width = 114
        Height = 13
        Alignment = taRightJustify
        AutoSize = False
        Caption = #25968#25454#24211#30331#24405#21517#65306
      end
      object Label19: TLabel
        Left = 28
        Top = 82
        Width = 114
        Height = 13
        Alignment = taRightJustify
        AutoSize = False
        Caption = #25968#25454#24211#23494#30721#65306
      end
      object Label26: TLabel
        Left = 28
        Top = 103
        Width = 114
        Height = 13
        Alignment = taRightJustify
        AutoSize = False
        Caption = #25968#25454#24211#65306
      end
      object hostIP: TEdit
        Left = 140
        Top = 20
        Width = 83
        Height = 21
        TabOrder = 0
        Text = '127.0.0.1'
      end
      object DBUser: TEdit
        Left = 140
        Top = 47
        Width = 83
        Height = 21
        TabOrder = 1
        Text = 'sa'
      end
      object DBPass: TEdit
        Left = 140
        Top = 74
        Width = 83
        Height = 21
        TabOrder = 2
        Text = '123456'
      end
      object DB: TEdit
        Left = 140
        Top = 101
        Width = 83
        Height = 21
        TabOrder = 3
        Text = 'DB_QRCode'
      end
      object TButton
        Left = 140
        Top = 128
        Width = 83
        Height = 25
        Caption = #36830#25509
        TabOrder = 4
        OnClick = btnConnectClick
      end
    end
    object TabSheet6: TTabSheet
      Caption = #23545#27604#20108#32500#30721
      ImageIndex = 8
      TabVisible = False
      object GroupBox6: TGroupBox
        Left = 0
        Top = 0
        Width = 936
        Height = 457
        Align = alTop
        Caption = #23545#27604#25991#20214
        TabOrder = 0
        Visible = False
        object grpA1: TGroupBox
          Left = 2
          Top = 15
          Width = 932
          Height = 76
          Align = alTop
          Caption = '2'
          TabOrder = 0
          object lbl4: TLabel
            Left = -15
            Top = 13
            Width = 80
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #24453#21305#37197#25991#20214
          end
          object lbl5: TLabel
            Left = 26
            Top = 41
            Width = 39
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #28304#25991#20214
          end
          object lbl9: TLabel
            Left = 411
            Top = 13
            Width = 80
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #24453#21305#37197#25991#20214
          end
          object lbl10: TLabel
            Left = 452
            Top = 41
            Width = 39
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #28304#25991#20214
          end
          object edtPath1: TEdit
            Left = 73
            Top = 13
            Width = 121
            Height = 21
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 0
          end
          object edtSPath1: TEdit
            Left = 73
            Top = 41
            Width = 121
            Height = 21
            NumbersOnly = True
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 1
          end
          object btnDb1: TButton
            Tag = 1
            Left = 206
            Top = 13
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 2
            OnClick = btnDb1Click
          end
          object btnDbS1: TButton
            Tag = 1
            Left = 206
            Top = 41
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 3
            OnClick = btnDbS1Click
          end
          object btnClear1: TButton
            Tag = 1
            Left = 287
            Top = 13
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 4
            OnClick = btnClear1Click
          end
          object btnSClear1: TButton
            Tag = 1
            Left = 287
            Top = 41
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 5
            OnClick = btnSClear1Click
          end
          object Check1: TCheckBox
            Tag = 1
            Left = 334
            Top = 13
            Width = 78
            Height = 17
            Caption = #21305#37197#36887#21495
            TabOrder = 6
            OnClick = Check1Click
          end
          object CheckS1: TCheckBox
            Tag = 1
            Left = 334
            Top = 42
            Width = 78
            Height = 17
            Caption = #21305#37197#31354#26684
            TabOrder = 7
            OnClick = CheckS1Click
          end
          object edtPath2: TEdit
            Left = 497
            Top = 13
            Width = 121
            Height = 21
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 8
          end
          object edtSPath2: TEdit
            Left = 497
            Top = 41
            Width = 121
            Height = 21
            NumbersOnly = True
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 9
          end
          object btnDb2: TButton
            Tag = 2
            Left = 624
            Top = 13
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 10
            OnClick = btnDb1Click
          end
          object btnDbS2: TButton
            Tag = 2
            Left = 624
            Top = 41
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 11
            OnClick = btnDbS1Click
          end
          object btnClear2: TButton
            Tag = 2
            Left = 702
            Top = 13
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 12
            OnClick = btnClear1Click
          end
          object btnSClear2: TButton
            Tag = 2
            Left = 702
            Top = 41
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 13
            OnClick = btnSClear1Click
          end
          object Check2: TCheckBox
            Tag = 2
            Left = 749
            Top = 14
            Width = 146
            Height = 17
            Caption = #21305#37197#36887#21495
            TabOrder = 14
            OnClick = Check1Click
          end
          object CheckS2: TCheckBox
            Tag = 2
            Left = 749
            Top = 42
            Width = 114
            Height = 17
            Caption = #21305#37197#31354#26684
            TabOrder = 15
            OnClick = CheckS1Click
          end
        end
        object grpA2: TGroupBox
          Left = 2
          Top = 91
          Width = 932
          Height = 76
          Align = alTop
          Caption = '4'
          TabOrder = 1
          object lbl11: TLabel
            Left = -15
            Top = 20
            Width = 80
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #24453#21305#37197#25991#20214
          end
          object lbl12: TLabel
            Left = 26
            Top = 47
            Width = 39
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #28304#25991#20214
          end
          object lbl13: TLabel
            Left = 411
            Top = 20
            Width = 80
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #24453#21305#37197#25991#20214
          end
          object lbl14: TLabel
            Left = 452
            Top = 47
            Width = 39
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #28304#25991#20214
          end
          object edtPath3: TEdit
            Left = 73
            Top = 20
            Width = 121
            Height = 21
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 0
          end
          object edtSPath3: TEdit
            Left = 73
            Top = 47
            Width = 121
            Height = 21
            NumbersOnly = True
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 1
          end
          object btnDb3: TButton
            Tag = 3
            Left = 206
            Top = 20
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 2
            OnClick = btnDb1Click
          end
          object btnDbS3: TButton
            Tag = 3
            Left = 206
            Top = 47
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 3
            OnClick = btnDbS1Click
          end
          object btnSClear3: TButton
            Tag = 3
            Left = 287
            Top = 47
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 4
            OnClick = btnSClear1Click
          end
          object btnClear3: TButton
            Tag = 3
            Left = 287
            Top = 20
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 5
            OnClick = btnClear1Click
          end
          object Check3: TCheckBox
            Tag = 3
            Left = 334
            Top = 20
            Width = 78
            Height = 17
            Caption = #21305#37197#36887#21495
            TabOrder = 6
            OnClick = Check1Click
          end
          object CheckS3: TCheckBox
            Tag = 3
            Left = 334
            Top = 47
            Width = 78
            Height = 17
            Caption = #21305#37197#31354#26684
            TabOrder = 7
            OnClick = CheckS1Click
          end
          object edtPath4: TEdit
            Left = 497
            Top = 20
            Width = 121
            Height = 21
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 8
          end
          object edtSPath4: TEdit
            Left = 497
            Top = 47
            Width = 121
            Height = 21
            NumbersOnly = True
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 9
          end
          object btnDb4: TButton
            Tag = 4
            Left = 624
            Top = 20
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 10
            OnClick = btnDb1Click
          end
          object btnDbS4: TButton
            Tag = 4
            Left = 624
            Top = 47
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 11
            OnClick = btnDbS1Click
          end
          object btnSClear4: TButton
            Tag = 4
            Left = 702
            Top = 47
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 12
            OnClick = btnSClear1Click
          end
          object btnClear4: TButton
            Tag = 4
            Left = 702
            Top = 20
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 13
            OnClick = btnClear1Click
          end
          object Check4: TCheckBox
            Tag = 4
            Left = 749
            Top = 20
            Width = 146
            Height = 17
            Caption = #21305#37197#36887#21495
            TabOrder = 14
            OnClick = Check1Click
          end
          object CheckS4: TCheckBox
            Tag = 4
            Left = 749
            Top = 47
            Width = 114
            Height = 17
            Caption = #21305#37197#31354#26684
            TabOrder = 15
            OnClick = CheckS1Click
          end
        end
        object grpA3: TGroupBox
          Left = 2
          Top = 167
          Width = 932
          Height = 76
          Align = alTop
          Caption = '6'
          TabOrder = 2
          object lbl15: TLabel
            Left = -15
            Top = 15
            Width = 80
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #24453#21305#37197#25991#20214
          end
          object lbl16: TLabel
            Left = 26
            Top = 58
            Width = 39
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #28304#25991#20214
          end
          object lbl17: TLabel
            Left = 411
            Top = 15
            Width = 80
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #24453#21305#37197#25991#20214
          end
          object lbl18: TLabel
            Left = 452
            Top = 58
            Width = 39
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #28304#25991#20214
          end
          object edtPath5: TEdit
            Left = 73
            Top = 15
            Width = 121
            Height = 21
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 0
          end
          object edtSPath5: TEdit
            Left = 73
            Top = 50
            Width = 121
            Height = 21
            NumbersOnly = True
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 1
          end
          object btnDb5: TButton
            Tag = 5
            Left = 206
            Top = 15
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 2
            OnClick = btnDb1Click
          end
          object btnDbS5: TButton
            Tag = 5
            Left = 206
            Top = 46
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 3
            OnClick = btnDbS1Click
          end
          object btnSClear5: TButton
            Tag = 5
            Left = 287
            Top = 46
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 4
            OnClick = btnSClear1Click
          end
          object btnClear5: TButton
            Tag = 5
            Left = 287
            Top = 15
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 5
            OnClick = btnClear1Click
          end
          object Check5: TCheckBox
            Tag = 5
            Left = 334
            Top = 15
            Width = 78
            Height = 17
            Caption = #21305#37197#36887#21495
            TabOrder = 6
            OnClick = Check1Click
          end
          object CheckS5: TCheckBox
            Tag = 5
            Left = 334
            Top = 49
            Width = 78
            Height = 17
            Caption = #21305#37197#31354#26684
            TabOrder = 7
            OnClick = CheckS1Click
          end
          object edtPath6: TEdit
            Left = 497
            Top = 15
            Width = 121
            Height = 21
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 8
          end
          object edtSPath6: TEdit
            Left = 497
            Top = 50
            Width = 121
            Height = 21
            NumbersOnly = True
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 9
          end
          object btnDb6: TButton
            Tag = 6
            Left = 624
            Top = 15
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 10
            OnClick = btnDb1Click
          end
          object btnDbS6: TButton
            Tag = 6
            Left = 624
            Top = 46
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 11
            OnClick = btnDbS1Click
          end
          object btnSClear6: TButton
            Tag = 6
            Left = 702
            Top = 46
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 12
            OnClick = btnSClear1Click
          end
          object btnClear6: TButton
            Tag = 6
            Left = 702
            Top = 15
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 13
            OnClick = btnClear1Click
          end
          object Check6: TCheckBox
            Tag = 6
            Left = 749
            Top = 17
            Width = 146
            Height = 17
            Caption = #21305#37197#36887#21495
            TabOrder = 14
            OnClick = Check1Click
          end
          object CheckS6: TCheckBox
            Tag = 6
            Left = 749
            Top = 49
            Width = 114
            Height = 17
            Caption = #21305#37197#31354#26684
            TabOrder = 15
            OnClick = CheckS1Click
          end
        end
        object grpA4: TGroupBox
          Left = 2
          Top = 243
          Width = 932
          Height = 76
          Align = alTop
          Caption = '8'
          TabOrder = 3
          object lbl19: TLabel
            Left = -15
            Top = 17
            Width = 80
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #24453#21305#37197#25991#20214
          end
          object lbl20: TLabel
            Left = 26
            Top = 52
            Width = 39
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #28304#25991#20214
          end
          object lbl21: TLabel
            Left = 411
            Top = 17
            Width = 80
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #24453#21305#37197#25991#20214
          end
          object lbl22: TLabel
            Left = 452
            Top = 52
            Width = 39
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #28304#25991#20214
          end
          object edtPath7: TEdit
            Left = 73
            Top = 17
            Width = 121
            Height = 21
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 0
          end
          object edtSPath7: TEdit
            Left = 73
            Top = 52
            Width = 121
            Height = 21
            NumbersOnly = True
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 1
          end
          object btnDb7: TButton
            Tag = 7
            Left = 206
            Top = 17
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 2
            OnClick = btnDb1Click
          end
          object btnDbS7: TButton
            Tag = 7
            Left = 206
            Top = 52
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 3
            OnClick = btnDbS1Click
          end
          object btnSClear7: TButton
            Tag = 7
            Left = 287
            Top = 52
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 4
            OnClick = btnSClear1Click
          end
          object btnClear7: TButton
            Tag = 7
            Left = 287
            Top = 17
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 5
            OnClick = btnClear1Click
          end
          object Check7: TCheckBox
            Tag = 7
            Left = 334
            Top = 17
            Width = 78
            Height = 17
            Caption = #21305#37197#36887#21495
            TabOrder = 6
            OnClick = Check1Click
          end
          object CheckS7: TCheckBox
            Tag = 7
            Left = 334
            Top = 52
            Width = 78
            Height = 17
            Caption = #21305#37197#31354#26684
            TabOrder = 7
            OnClick = CheckS1Click
          end
          object edtPath8: TEdit
            Left = 497
            Top = 17
            Width = 121
            Height = 21
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 8
          end
          object edtSPath8: TEdit
            Left = 497
            Top = 52
            Width = 121
            Height = 21
            NumbersOnly = True
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 9
          end
          object btnDb8: TButton
            Tag = 8
            Left = 624
            Top = 17
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 10
            OnClick = btnDb1Click
          end
          object btnDbS8: TButton
            Tag = 8
            Left = 624
            Top = 52
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 11
            OnClick = btnDbS1Click
          end
          object btnSClear8: TButton
            Tag = 8
            Left = 702
            Top = 52
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 12
            OnClick = btnSClear1Click
          end
          object btnClear8: TButton
            Tag = 8
            Left = 702
            Top = 18
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 13
            OnClick = btnClear1Click
          end
          object Check8: TCheckBox
            Tag = 8
            Left = 749
            Top = 19
            Width = 146
            Height = 17
            Caption = #21305#37197#36887#21495
            TabOrder = 14
            OnClick = Check1Click
          end
          object CheckS8: TCheckBox
            Tag = 8
            Left = 749
            Top = 52
            Width = 114
            Height = 17
            Caption = #21305#37197#31354#26684
            TabOrder = 15
            OnClick = CheckS1Click
          end
        end
        object grpA5: TGroupBox
          Left = 2
          Top = 319
          Width = 932
          Height = 82
          Align = alTop
          Caption = '10'
          TabOrder = 4
          object lbl23: TLabel
            Left = -15
            Top = 15
            Width = 80
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #24453#21305#37197#25991#20214
          end
          object lbl24: TLabel
            Left = 26
            Top = 50
            Width = 39
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #28304#25991#20214
          end
          object lbl25: TLabel
            Left = 411
            Top = 16
            Width = 80
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #24453#21305#37197#25991#20214
          end
          object lbl26: TLabel
            Left = 452
            Top = 50
            Width = 39
            Height = 13
            Alignment = taRightJustify
            AutoSize = False
            Caption = #28304#25991#20214
          end
          object edtPath9: TEdit
            Left = 73
            Top = 15
            Width = 121
            Height = 21
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 0
          end
          object edtSPath9: TEdit
            Left = 73
            Top = 46
            Width = 121
            Height = 21
            NumbersOnly = True
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 1
          end
          object btnDb9: TButton
            Tag = 9
            Left = 206
            Top = 15
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 2
            OnClick = btnDb1Click
          end
          object btnDbS9: TButton
            Tag = 9
            Left = 206
            Top = 42
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 3
            OnClick = btnDbS1Click
          end
          object btnSClear9: TButton
            Tag = 9
            Left = 287
            Top = 42
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 4
            OnClick = btnSClear1Click
          end
          object btnClear9: TButton
            Tag = 9
            Left = 287
            Top = 15
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 5
            OnClick = btnClear1Click
          end
          object Check9: TCheckBox
            Tag = 9
            Left = 334
            Top = 15
            Width = 78
            Height = 17
            Caption = #21305#37197#36887#21495
            TabOrder = 6
            OnClick = Check1Click
          end
          object CheckS9: TCheckBox
            Tag = 9
            Left = 334
            Top = 47
            Width = 78
            Height = 17
            Caption = #21305#37197#31354#26684
            TabOrder = 7
            OnClick = CheckS1Click
          end
          object edtPath10: TEdit
            Left = 497
            Top = 15
            Width = 121
            Height = 21
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 8
          end
          object edtSPath10: TEdit
            Left = 497
            Top = 46
            Width = 121
            Height = 21
            NumbersOnly = True
            ParentShowHint = False
            ReadOnly = True
            ShowHint = True
            TabOrder = 9
          end
          object btnDb10: TButton
            Tag = 10
            Left = 624
            Top = 15
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 10
            OnClick = btnDb1Click
          end
          object btnDbS10: TButton
            Tag = 10
            Left = 624
            Top = 42
            Width = 75
            Height = 25
            Caption = #23548#20837#25991#20214
            TabOrder = 11
            OnClick = btnDbS1Click
          end
          object btnClear10: TButton
            Tag = 10
            Left = 702
            Top = 15
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 12
            OnClick = btnClear1Click
          end
          object btnSClear10: TButton
            Tag = 10
            Left = 702
            Top = 42
            Width = 41
            Height = 25
            Caption = #28165#31354
            TabOrder = 13
            OnClick = btnSClear1Click
          end
          object Check10: TCheckBox
            Tag = 10
            Left = 749
            Top = 16
            Width = 146
            Height = 17
            Caption = #21305#37197#36887#21495
            TabOrder = 14
            OnClick = Check1Click
          end
          object CheckS10: TCheckBox
            Tag = 10
            Left = 749
            Top = 48
            Width = 114
            Height = 17
            Caption = #21305#37197#31354#26684
            TabOrder = 15
            OnClick = CheckS1Click
          end
        end
        object btnCheck2: TButton
          Left = 406
          Top = 407
          Width = 131
          Height = 44
          Caption = #21305#37197#20108#32500#30721
          TabOrder = 5
          OnClick = btnCheck2Click
        end
        object btnClear: TButton
          Left = 269
          Top = 407
          Width = 131
          Height = 44
          Caption = #20840#28165
          TabOrder = 6
          OnClick = btnClearClick
        end
      end
      object GroupBox7: TGroupBox
        Left = 0
        Top = 457
        Width = 936
        Height = 152
        Align = alClient
        Caption = #25805#20316#26085#24535
        TabOrder = 1
        object mmoLog2: TMemo
          Left = 2
          Top = 15
          Width = 932
          Height = 135
          Align = alClient
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          ReadOnly = True
          ScrollBars = ssBoth
          TabOrder = 0
        end
      end
    end
  end
  object ADOQuery1: TADOQuery
    Connection = ADOConnection1
    CommandTimeout = 0
    Parameters = <>
    Left = 912
    Top = 528
  end
  object ADOConnection1: TADOConnection
    CommandTimeout = 0
    Connected = True
    ConnectionString = 
      'Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;U' +
      'ser ID=sa;Initial Catalog=DB_RQCode;Data Source=127.0.0.1'
    ConnectionTimeout = 15000
    LoginPrompt = False
    Provider = 'SQLOLEDB.1'
    Left = 832
    Top = 528
  end
  object OpenDialogA: TOpenDialog
    Left = 568
    Top = 520
  end
  object OpenDialogB: TOpenDialog
    Left = 712
    Top = 528
  end
  object XLS: TXLSReadWriteII5
    ComponentVersion = '5.20.62'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 464
    Top = 544
  end
  object qryCustFile: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    AfterOpen = qryCustFileAfterOpen
    CommandTimeout = 0
    Parameters = <>
    SQL.Strings = (
      'select * from ##tmpCustFile')
    Left = 80
    Top = 560
    object qryCustFileselected: TBooleanField
      FieldName = 'selected'
    end
    object qryCustFilefileName: TStringField
      FieldName = 'fileName'
      Size = 500
    end
    object qryCustFileisContainFirst: TBooleanField
      FieldName = 'isContainFirst'
    end
  end
  object dsCustFile: TDataSource
    DataSet = qryCustFile
    Left = 80
    Top = 496
  end
  object conOA: TADOConnection
    CommandTimeout = 0
    Connected = True
    ConnectionString = 
      'Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;U' +
      'ser ID=sa;Initial Catalog=DB_RQCode;Data Source=127.0.0.1'
    ConnectionTimeout = 15000
    LoginPrompt = False
    Provider = 'SQLOLEDB.1'
    Left = 192
    Top = 528
  end
  object dsMulti: TDataSource
    DataSet = qryMulti
    Left = 344
    Top = 536
  end
  object qryMulti: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    AfterOpen = qryMultiAfterOpen
    CommandTimeout = 0
    Parameters = <>
    SQL.Strings = (
      'select * from ##tmpCustFile')
    Left = 312
    Top = 496
    object BooleanField1: TBooleanField
      FieldName = 'selected'
    end
    object StringField1: TStringField
      FieldName = 'fileName'
      Size = 500
    end
    object BooleanField2: TBooleanField
      FieldName = 'isContainFirst'
      Visible = False
    end
  end
  object adoCheck: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      '')
    Left = 564
    Top = 432
  end
  object adoCheck3: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select top 10  A.'#26465#30721#20869#23481',A.'#39564#35777#30721' from [dbo].[1] A')
    Left = 640
    Top = 262
  end
  object XLS2: TXLSReadWriteII5
    ComponentVersion = '5.20.62'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 704
    Top = 264
  end
end
