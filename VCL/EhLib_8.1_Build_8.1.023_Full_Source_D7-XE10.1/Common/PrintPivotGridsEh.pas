{*******************************************************}
{                                                       }
{                      EhLib v8.1                       }
{              TPrintPivotGridEh component              }
{                    (Build 8.1.002)                    }
{                                                       }
{   Copyright (c) 2015-2015 by Dmitry V. Bolshakov      }
{                                                       }
{*******************************************************}

unit PrintPivotGridsEh;

interface

{$I EhLib.Inc}

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls,
  EhLibVCL, Types,
{$IFDEF EH_LIB_17} System.Contnrs, {$ENDIF}
  StdCtrls, ImgList, Forms, DB, Variants,
  PivotGridsEh, PrintUtilsEh, PrntsEh, PrViewEh, GridsEh, ToolCtrlsEh;

type

{ TPivotGridPrintServiceEh }

  TPivotGridPrintServiceEh = class(TCustomPivotGridPrintServiceEh)
  private
//    FGridControl: TCustomControl;

  protected
    function CheckDrawLine(ACol, ARow: Integer; BorderType: TGridCellBorderTypeEh; var Color: TColor; var Width: Integer): Boolean; override;
    function GetControlCanvas: TCanvas; override;

    procedure PrintCellData(ACol, ARow: Integer; ARect: TRect); override;
    procedure SetColRowSize; override;

    procedure DrawFieldNameCell(ACol, ARow: Integer; ARect: TRect; PivotCel: TPivotCellEh); virtual;
    procedure DrawDataCell(ACol, ARow: Integer; ARect: TRect; PivotCel: TPivotCellEh); virtual;
    procedure DrawAxisValueCell(ACol, ARow: Integer; ARect: TRect; PivotCel: TPivotCellEh); virtual;
    procedure DrawAxisValueCellData(ACol, ARow: Integer; ARect: TRect; PivotCel: TPivotCellEh; ShowGroupingSign, ShowValue: Boolean); virtual;

//    property GridControl: TCustomControl read FGridControl write FGridControl;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure PrintTo(VPrinter: TVirtualPrinter); override;

  published
//    property Grid: TCustomPivotGridEh read FGrid write FGrid;

    property Scale;
    property FitToPagesWide;
    property FitToPagesTall;
    property ScalingMode;
    property Orientation;
    property ColorSchema;
    property PageFooter;
    property PageHeader;
    property PageMargins;
    property TextBeforeContent;
    property TextAfterContent;

  end;

implementation

uses EhLibConsts;

type
  TGridCrack = class(TCustomPivotGridEh);

var
  FDrawCellParams: TPivotCellDrawParamsEh;

procedure InitUnit;
begin
  FDrawCellParams := TPivotCellDrawParamsEh.Create;
end;

procedure FinalyUnit;
begin
  FreeAndNil(FDrawCellParams);
end;

{ TPivotGridPrintServiceEh }

constructor TPivotGridPrintServiceEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TPivotGridPrintServiceEh.Destroy;
begin
  inherited Destroy;
end;

function TPivotGridPrintServiceEh.GetControlCanvas: TCanvas;
begin
  Result := TGridCrack(Grid).Canvas;
end;

//function TPrintPivotGridEh.CreatePrintParams: TBasePrintParamsEh;
//begin
//  Result := TPrintParamsEh.Create;
//end;
//
//function TPrintPivotGridEh.GetPrintParams: TPrintParamsEh;
//begin
//  Result := TPrintParamsEh(inherited PrintParams);
//end;
//
//procedure TPrintPivotGridEh.SetPrintParams(const Value: TPrintParamsEh);
//begin
//  inherited PrintParams := Value;
//end;

function TPivotGridPrintServiceEh.CheckDrawLine(ACol, ARow: Integer;
  BorderType: TGridCellBorderTypeEh; var Color: TColor;
  var Width: Integer): Boolean;
var
  IsExtend: Boolean;
begin
  TGridCrack(Grid).CheckDrawCellBorder(ACol, ARow, BorderType, Result, Color, IsExtend);
  Width := PenW;
end;

procedure TPivotGridPrintServiceEh.SetColRowSize;
var
  i: Integer;
  Grid: TGridCrack;
begin
  Grid := TGridCrack(Self.Grid);
  SetLength(ColWidths, Grid.FullColCount);
  for i := 0 to Length(ColWidths)-1 do
    ColWidths[i] := Round(Grid.ColWidths[i] * ScaleX);

  SetLength(RowHeights, Grid.FullRowCount);
  for i := 0 to Length(RowHeights)-1 do
    RowHeights[i] := Round(Grid.RowHeights[i] * ScaleY);
end;

procedure TPivotGridPrintServiceEh.PrintCellData(ACol, ARow: Integer; ARect: TRect);
var
  PivotCel: TPivotCellEh;
  Grid: TGridCrack;
//  IsCellSelectionFilledByTheme: Boolean;
begin
  Grid := TGridCrack(Self.Grid);
  Canvas.Font := Grid.Font;
//  IsCellSelectionFilledByTheme := False;

//  if (gdFixed in AState) or (ACol >= ColCount) or (ARow >= RowCount) then
//    Canvas.Brush.Color := FixedColor
//  else
//    Canvas.Brush.Color := Color;

  if (ACol < Grid.FVisPivotColCount) and (ARow < Grid.FVisPivotRowCount) then
  begin
    PivotCel := Grid.VisPivotGridArray[ACol, ARow];
    if PivotCel.CelType in [sctFieldNameForRowEh, sctFieldNameForColEh] then
      DrawFieldNameCell(ACol, ARow, ARect, PivotCel)
    else if PivotCel.CelType in [sctDataEh, sctHorzAggregateData, sctVertAggregateData] then
      DrawDataCell(ACol, ARow, ARect, PivotCel)
    else if PivotCel.CelType in [sctAxisValueEh, sctValuesColCaptionEh] then
      DrawAxisValueCell(ACol, ARow, ARect, PivotCel)
    else
    begin
//      Canvas.FillRect(ARect);
    end;
  end else
  begin
//    Canvas.FillRect(ARect);
  end;

end;

procedure TPivotGridPrintServiceEh.DrawAxisValueCell(ACol, ARow: Integer;
  ARect: TRect; PivotCel: TPivotCellEh);
var
  ShowValue: Boolean;
  ShowGroupingSign: Boolean;
  Grid: TGridCrack;

  function NextRowCellIsLevelSum(): Boolean;
  var
    NextCel: TPivotCellEh;
  begin
    if ARow < Grid.FullRowCount-1 then
    begin
      NextCel := Grid.VisPivotGridArray[ACol, ARow+1];
      Result := NextCel.VertAggrLevelCol = NextCel.VertAggrLevelRow;
    end else
      Result := False;
  end;

begin
  if CheckCellAreaDrawn(ACol, ARow) then Exit;
  Grid := TGridCrack(Self.Grid);

  if (PivotCel.VertAggrLevelRow > 0) and (PivotCel.VertAggrLevelRow >= PivotCel.VertAggrLevelCol)
   or
     (PivotCel.HorzAggrLevelCol > 0) and (PivotCel.HorzAggrLevelCol >= PivotCel.HorzAggrLevelRow)
  then
    Canvas.Font := TGridCrack(Grid).GridCellParams.AxisAggregateFont
  else
    Canvas.Font := TGridCrack(Grid).GridCellParams.AxisFont;

  if //PivotCell.ShowValue and
     (PivotCel.MasterRow >= 0) and
     (PivotCel.DrawDownLine = False) and
     (ARect.Top <= TGridCrack(Grid).VertAxis.FixedBoundary)
  then
  begin
    if NextRowCellIsLevelSum() then
    begin
      ShowGroupingSign := True;
      ShowValue := True;
    end else
      Exit;
  end else
  begin
    ShowGroupingSign := PivotCel.ShowGroupingSign;
    ShowValue := PivotCel.ShowValue;
  end;

  DrawAxisValueCellData(ACol, ARow, ARect, PivotCel, ShowGroupingSign, ShowValue);
end;

procedure TPivotGridPrintServiceEh.DrawAxisValueCellData(ACol, ARow: Integer;
  ARect: TRect; PivotCel: TPivotCellEh; ShowGroupingSign, ShowValue: Boolean);
var
  s: String;
  TreeSignRect: TRect;
  TreeElement: TTreeElementEh;
  Al: TAlignment;
  Grid: TGridCrack;
begin
  Grid := TGridCrack(Self.Grid);
  Al := taLeftJustify;
//  if UseRightToLeftAlignment then ChangeBiDiModeAlignment(Al);

  if not Grid.DataSourceIsEmpty and
     Grid.GrandTotalRowVisible and
    (PivotCel.VertAggrLevelRow = Grid.ActualRowFlds.Count) and
    (PivotCel.VertAggrLevelCol = Grid.ActualRowFlds.Count-1)
  then
    s := SGrandTotalEh
  else if not Grid.DataSourceIsEmpty and
          Grid.FShowGrandTotalCols and
          (ACol = Grid.FullColCount-1 * Grid.ActualValueFields.Count) and
          (ARow = Grid.ActualColFlds.Count)
  then
    s := SGrandTotalEh
  else
  begin
    if ShowValue = True
      then s := VarToStr(PivotCel.Value)
      else s := '';

    if PivotCel.ShowValue and
      (PivotCel.PivotField <> nil) and
      (PivotCel.VertAggrLevelRow < Grid.ActualRowFlds.Count) and
      (PivotCel.HorzAggrLevelCol < Grid.ActualColFlds.Count)
    then
      s := PivotCel.PivotField.ValueAsDispayText(PivotCel.Value);

    if ((PivotCel.VertAggrLevelRow > 0) or (PivotCel.HorzAggrLevelCol > 0)) and (s <> '') then
      s := s + ' ' + STotalEh;
  end;

  if ShowGroupingSign then
  begin
    TreeSignRect := ARect;
    TreeSignRect.Right := TreeSignRect.Left + Trunc(18 * ScaleX);
    TreeElement := tehMinus;
    if PivotCel.RowsTreeNode <> nil then
    begin
      if PivotCel.RowsTreeNode.Expanded
        then TreeElement := tehMinus
        else TreeElement := tehPlus;
    end else if PivotCel.ColsTreeNode <> nil then
    begin
      if PivotCel.ColsTreeNode.Expanded
        then TreeElement := tehMinus
        else TreeElement := tehPlus;
    end;
    Canvas.Pen.Color := Grid.GridLineColors.GetDarkColor;
    Canvas.Brush.Color := Grid.GridCellParams.ActualAxisColor;
    DrawTreeElement(Canvas, TreeSignRect, TreeElement, False, ScaleX, ScaleY, False, False, tvgsClassicEh);
    ARect.Left := TreeSignRect.Right;
  end;

  if ShowValue then
    WriteTextEh(Canvas, ARect, False, PenW, PenW, s, Al, tlTop, False, False, 0, 0, False, False);
end;

procedure TPivotGridPrintServiceEh.DrawDataCell(ACol, ARow: Integer; ARect: TRect;
  PivotCel: TPivotCellEh);
var
  s: String;
  Al: TAlignment;
  IsCellFilled: Boolean;
  ACellRect: TRect;
  AState: TGridDrawState;
  Grid: TGridCrack;
begin
  Grid := TGridCrack(Self.Grid);
  Al := taRightJustify;

  Canvas.Brush.Color := Grid.GridCellParams.ActualDataColor;
  Canvas.Font := Grid.GridCellParams.DataFont;
  IsCellFilled := True;

  Grid.FillDrawCellParams(FDrawCellParams, ACol, ARow, ARect, AState, PivotCel);

  s := FDrawCellParams.DisplayValue;
  Canvas.Font := FDrawCellParams.Font;

  ACellRect := ARect;

  WriteTextEh(Canvas, ARect, not IsCellFilled, PenW, PenW, s,
    Al, tlTop, False, False, 0, 0, False, False);
end;


procedure TPivotGridPrintServiceEh.DrawFieldNameCell(ACol, ARow: Integer; ARect: TRect;
  PivotCel: TPivotCellEh);
var
  s: String;
  Al: TAlignment;
  Grid: TGridCrack;
begin
  Grid := TGridCrack(Self.Grid);
  Canvas.Font := Grid.GridCellParams.FieldNameFont;

  if PivotCel.ShowValue = True then
    s := VarToStr(PivotCel.Value)
  else
    s := '';

  Al := taLeftJustify;
  if PivotCel.ShowValue then
    WriteTextEh(Canvas, ARect, False, PenW, PenW, s, Al, tlTop, False, False, 0, 0,
      False, False);
end;

procedure TPivotGridPrintServiceEh.PrintTo(VPrinter: TVirtualPrinter);
begin
  if Grid = nil then
    raise Exception.Create('TPrintDBVertGridEh.Grid property is not assigned');
  inherited PrintTo(VPrinter);
end;

initialization
  InitUnit;
finalization
  FinalyUnit;
end.
