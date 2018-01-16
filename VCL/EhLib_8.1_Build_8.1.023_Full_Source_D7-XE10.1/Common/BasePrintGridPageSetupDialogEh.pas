{*******************************************************}
{                                                       }
{                        EhLib v8.1                     }
{                                                       }
{   Copyright (c) 2015-2015 by Dmitry V. Bolshakov      }
{                                                       }
{*******************************************************}

{$I EhLib.Inc}

unit BasePrintGridPageSetupDialogEh;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, Dialogs, ExtCtrls, StdCtrls, ComCtrls,
{$IFDEF EH_LIB_17} System.UITypes, {$ENDIF}
  Mask, Printers, Math, PrintUtilsEh, DBCtrlsEh, Buttons;

type
  TfSpreadGridsPrintPageSetupDialogEh = class(TForm)
    pcPageSetup: TPageControl;
    tsPage: TTabSheet;
    tsMargins: TTabSheet;
    tsHeaderFooter: TTabSheet;
    bOk: TButton;
    bCancel: TButton;
    Label1: TLabel;
    Bevel1: TBevel;
    Label2: TLabel;
    Bevel2: TBevel;
    Panel1: TPanel;
    rbPortrait: TRadioButton;
    rbLandscape: TRadioButton;
    Image1: TImage;
    Image2: TImage;
    Panel2: TPanel;
    rdAdjust: TRadioButton;
    rbFit: TRadioButton;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    neScale: TDBNumberEditEh;
    neFitToWide: TDBNumberEditEh;
    neFitToTall: TDBNumberEditEh;
    neLeftMargin: TDBNumberEditEh;
    neTopMargin: TDBNumberEditEh;
    neBottomMargin: TDBNumberEditEh;
    neRightMargin: TDBNumberEditEh;
    neHeader: TDBNumberEditEh;
    neFooter: TDBNumberEditEh;
    reHeaderLeft: TDBRichEditEh;
    reHeaderCenter: TDBRichEditEh;
    reHeaderRight: TDBRichEditEh;
    reFooterLeft: TDBRichEditEh;
    reFooterCenter: TDBRichEditEh;
    reFooterRight: TDBRichEditEh;
    Label6: TLabel;
    Label7: TLabel;
    spFont: TSpeedButton;
    spInsertPageNo: TSpeedButton;
    spInsertPages: TSpeedButton;
    spInsertDate: TSpeedButton;
    spInsertShortDate: TSpeedButton;
    spInsertLongDate: TSpeedButton;
    spInsertTime: TSpeedButton;
    Bevel3: TBevel;
    Bevel4: TBevel;
    FontDialog1: TFontDialog;
    procedure spInsertPageNoClick(Sender: TObject);
    procedure spInsertPagesClick(Sender: TObject);
    procedure spInsertDateClick(Sender: TObject);
    procedure spInsertShortDateClick(Sender: TObject);
    procedure spInsertLongDateClick(Sender: TObject);
    procedure spInsertTimeClick(Sender: TObject);
    procedure spFontClick(Sender: TObject);
    procedure Image1Click(Sender: TObject);
    procedure Image2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fSpreadGridsPrintPageSetupDialogEh: TfSpreadGridsPrintPageSetupDialogEh;

function ShowSpreadGridPageSetupDialog(PrintComponent: TBasePrintServiceComponentEh): Boolean;

implementation

{$R *.dfm}

type
  TPrintControlComponentCrack = class(TBasePrintServiceComponentEh);

function ShowSpreadGridPageSetupDialog(PrintComponent: TBasePrintServiceComponentEh): Boolean;
var
  f: TfSpreadGridsPrintPageSetupDialogEh;
  ppr: TPrintControlComponentCrack;

  procedure ResetRichtextAlignment(RichEdit: TDBRichEditEh; Alignment: TAlignment);
  begin
    RichEdit.SelectAll;
    RichEdit.Paragraph.Alignment := Alignment;
  end;

begin
  Result := False;
  f := TfSpreadGridsPrintPageSetupDialogEh.Create(Application);
  f.pcPageSetup.ActivePageIndex := 0;
  ppr := TPrintControlComponentCrack(PrintComponent);
  if ppr.Orientation = poPortrait
    then f.rbPortrait.Checked := True
    else f.rbLandscape.Checked := True;

  if ppr.ScalingMode = smAdjustToScaleEh
    then f.rdAdjust.Checked := True
    else f.rbFit.Checked := True;

  f.neScale.Value := ppr.Scale;
  f.neFitToWide.Value := ppr.FitToPagesWide;
  f.neFitToTall.Value := ppr.FitToPagesTall;

  f.neLeftMargin.Value := Round(ppr.PageMargins.CurRegionalMetricLeft*10) / 10;
  f.neRightMargin.Value := Round(ppr.PageMargins.CurRegionalMetricRight*10) / 10;
  f.neTopMargin.Value := Round(ppr.PageMargins.CurRegionalMetricTop*10) / 10;
  f.neBottomMargin.Value := Round(ppr.PageMargins.CurRegionalMetricBottom*10) / 10;
  f.neHeader.Value := Round(ppr.PageMargins.CurRegionalMetricHeader*10) / 10;
  f.neFooter.Value := Round(ppr.PageMargins.CurRegionalMetricFooter*10) / 10;

  f.reHeaderLeft.RtfText := ppr.PageHeader.LeftText;
  ResetRichtextAlignment(f.reHeaderLeft, taLeftJustify);
  f.reHeaderCenter.RtfText := ppr.PageHeader.CenterText;
  ResetRichtextAlignment(f.reHeaderCenter, taCenter);
  f.reHeaderRight.RtfText := ppr.PageHeader.RightText;
  ResetRichtextAlignment(f.reHeaderRight, taRightJustify);

  f.reFooterLeft.RtfText := ppr.PageFooter.LeftText;
  ResetRichtextAlignment(f.reFooterLeft, taLeftJustify);
  f.reFooterCenter.RtfText := ppr.PageFooter.CenterText;
  ResetRichtextAlignment(f.reFooterCenter, taCenter);
  f.reFooterRight.RtfText := ppr.PageFooter.RightText;
  ResetRichtextAlignment(f.reFooterRight, taRightJustify);

  if f.ShowModal = mrOk then
  begin

    if f.rbPortrait.Checked
      then ppr.Orientation := poPortrait
      else ppr.Orientation := poLandscape;

    if f.rdAdjust.Checked = True
      then ppr.ScalingMode := smAdjustToScaleEh
      else ppr.ScalingMode := smFitToPagesEh;

    ppr.Scale := f.neScale.Value;
    ppr.FitToPagesWide := f.neFitToWide.Value;
    ppr.FitToPagesTall := f.neFitToTall.Value;

    ppr.PageMargins.CurRegionalMetricLeft := f.neLeftMargin.Value;
    ppr.PageMargins.CurRegionalMetricRight := f.neRightMargin.Value;
    ppr.PageMargins.CurRegionalMetricTop := f.neTopMargin.Value;
    ppr.PageMargins.CurRegionalMetricBottom := f.neBottomMargin.Value;
    ppr.PageMargins.CurRegionalMetricHeader := f.neHeader.Value;
    ppr.PageMargins.CurRegionalMetricFooter := f.neFooter.Value;

    if f.reHeaderLeft.Lines.Text <> ''
      then ppr.PageHeader.LeftText := f.reHeaderLeft.RtfText
      else ppr.PageHeader.LeftText := '';

    if f.reHeaderCenter.Lines.Text <> ''
      then ppr.PageHeader.CenterText := f.reHeaderCenter.RtfText
      else ppr.PageHeader.CenterText := '';

    if f.reHeaderRight.Lines.Text <> ''
      then ppr.PageHeader.RightText := f.reHeaderRight.RtfText
      else ppr.PageHeader.RightText := '';

    if f.reFooterLeft.Lines.Text <> ''
      then ppr.PageFooter.LeftText := f.reFooterLeft.RtfText
      else ppr.PageFooter.LeftText := '';

    if f.reFooterCenter.Lines.Text <> ''
      then ppr.PageFooter.CenterText := f.reFooterCenter.RtfText
      else ppr.PageFooter.CenterText := '';

    if f.reFooterRight.Lines.Text <> ''
      then ppr.PageFooter.RightText := f.reFooterRight.RtfText
      else ppr.PageFooter.RightText := '';

    Result := True;
  end;

  f.Free;
end;

procedure TfSpreadGridsPrintPageSetupDialogEh.spInsertPageNoClick(
  Sender: TObject);
begin
  if ActiveControl is TDBRichEditEh then
    (ActiveControl as TDBRichEditEh).SelText := '&[Page]';
end;

procedure TfSpreadGridsPrintPageSetupDialogEh.spInsertPagesClick(
  Sender: TObject);
begin
  if ActiveControl is TDBRichEditEh then
    (ActiveControl as TDBRichEditEh).SelText := '&[Pages]';
end;

procedure TfSpreadGridsPrintPageSetupDialogEh.Image1Click(Sender: TObject);
begin
  rbPortrait.Checked := True;
end;

procedure TfSpreadGridsPrintPageSetupDialogEh.Image2Click(Sender: TObject);
begin
  rbLandscape.Checked := True;
end;

procedure TfSpreadGridsPrintPageSetupDialogEh.spFontClick(Sender: TObject);
var
  re: TDBRichEditEh;
begin
  if ActiveControl is TDBRichEditEh then
  begin
    re := ActiveControl as TDBRichEditEh;
    FontDialog1.Font.Assign(re.SelAttributes);
    if FontDialog1.Execute then
      re.SelAttributes.Assign(FontDialog1.Font);
  end;
end;

procedure TfSpreadGridsPrintPageSetupDialogEh.spInsertDateClick(
  Sender: TObject);
begin
  if ActiveControl is TDBRichEditEh then
    (ActiveControl as TDBRichEditEh).SelText := '&[Date]';
end;

procedure TfSpreadGridsPrintPageSetupDialogEh.spInsertShortDateClick(
  Sender: TObject);
begin
  if ActiveControl is TDBRichEditEh then
    (ActiveControl as TDBRichEditEh).SelText := '&[ShortDate]';
end;

procedure TfSpreadGridsPrintPageSetupDialogEh.spInsertLongDateClick(
  Sender: TObject);
begin
  if ActiveControl is TDBRichEditEh then
    (ActiveControl as TDBRichEditEh).SelText := '&[LongDate]';
end;

procedure TfSpreadGridsPrintPageSetupDialogEh.spInsertTimeClick(
  Sender: TObject);
begin
  if ActiveControl is TDBRichEditEh then
    (ActiveControl as TDBRichEditEh).SelText := '&[Time]';
end;

end.
