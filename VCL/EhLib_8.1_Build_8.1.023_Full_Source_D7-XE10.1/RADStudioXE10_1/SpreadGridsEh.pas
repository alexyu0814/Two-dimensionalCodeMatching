{*******************************************************}
{                                                       }
{                        EhLib v8.1                     }
{                                                       }
{                      SpreadGridsEh                    }
{                      Build 8.1.003                    }
{                                                       }
{   Copyright (c) 2015-2015 by Dmitry V. Bolshakov      }
{                                                       }
{*******************************************************}

{$I EhLib.Inc}

unit SpreadGridsEh;

interface

uses
{$IFDEF EH_LIB_17} System.Generics.Collections, {$ENDIF}
  Windows, SysUtils, Messages, Classes, Controls, Forms, StdCtrls, TypInfo,
{$IFDEF EH_LIB_5} Contnrs, {$ENDIF}
{$IFDEF EH_LIB_6} Variants, Types, {$ENDIF}
{$IFDEF EH_LIB_7} Themes, UxTheme, {$ENDIF}
{$IFDEF EH_LIB_17} System.UITypes, {$ENDIF}
{$IFDEF CIL}
  EhLibVCLNET,
  WinUtils,
{$ELSE}
  {$IFDEF FPC}
  EhLibLCL,
  {$ELSE}
  EhLibVCL,
  {$ENDIF}
{$ENDIF}
  GridsEh, ToolCtrlsEh, Graphics;

type

{ TSpreadGridCellEh }

  TSpreadGridCellEh = class(TPersistent)
  private
    FValue: Variant;
    FMasterMergeColOffset: Integer;
    FMasterMergeRowOffset: Integer;
    FMergeColCount: Integer;
    FMergeRowCount: Integer;
  public
    procedure Clear; virtual;
    property Value: Variant read FValue write FValue;
    property MasterMergeColOffset: Integer read FMasterMergeColOffset;
    property MasterMergeRowOffset: Integer read FMasterMergeRowOffset;
    property MergeColCount: Integer read FMergeColCount;
    property MergeRowCount: Integer read FMergeRowCount;
  end;

  TSpreadGridArray = array of array of TSpreadGridCellEh;

{ TCustomSpreadGridEh }

  TCustomSpreadGridEh = class(TCustomGridEh)
  private
    FSpreadArray: TSpreadGridArray;
//    FDrawenCellArr: array of Boolean;
    FDrawenCellArr: array of TGridCoord;
//    FDrawColCount, FDrawRowCount: Integer;
    function GetVisibleColCount: Integer;
    function GetVisibleRowCount: Integer;
    function GetCell(ACol, ARow: Integer): TSpreadGridCellEh;
  protected
    function NeedBufferedPaint: Boolean; override;
    function GetCellDisplayText(ACol, ARow: Integer): String; virtual;
    function CheckCellAreaDrawn(ACol, ARow: Integer): Boolean; virtual;
    function CreateSpreadGridCell(ACol, ARow: Integer): TSpreadGridCellEh; virtual;

    procedure Paint; override;
    procedure SetCellDrawn(ACol, ARow: Integer);
    procedure SetSpreadArraySize(AColCount, ARowCount: Integer);
    procedure SetCellCanvasParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); virtual;
    procedure DrawCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); override;
    procedure DrawCellArea(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); override;
    procedure DrawMasterCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; SpreadCell: TSpreadGridCellEh); virtual;
    procedure DrawMasterForMergedCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); virtual;

    property SpreadArray: TSpreadGridArray read FSpreadArray;
    property VisibleColCount: Integer read GetVisibleColCount;
    property VisibleRowCount: Integer read GetVisibleRowCount;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure SetGridSize(ANewColCount, ANewRowCount: Integer); virtual;
    procedure MergeCells(BaseCellCol, BaseCellRow, MergeColCount, MergeRowCount: Integer);
    property Cell[ACol, ARow: Integer]: TSpreadGridCellEh read GetCell;
  end;

implementation

{ TSpreadGridCellEh }

procedure TSpreadGridCellEh.Clear;
begin
  FValue := Variants.Null;
  FMasterMergeColOffset := 0;
  FMasterMergeRowOffset := 0;
  FMergeColCount := 0;
  FMergeRowCount := 0;
end;

{ TCustomSpreadGridEh }

constructor TCustomSpreadGridEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  SetGridSize(ColCount, RowCount);
  VertScrollBar.SmoothStep := True;
end;

destructor TCustomSpreadGridEh.Destroy;
begin
  SetSpreadArraySize(0, 0);
  inherited Destroy;
end;

function TCustomSpreadGridEh.CreateSpreadGridCell(ACol, ARow: Integer): TSpreadGridCellEh;
begin
  Result := TSpreadGridCellEh.Create;
end;

function TCustomSpreadGridEh.GetCell(ACol, ARow: Integer): TSpreadGridCellEh;
begin
  Result := FSpreadArray[ACol, ARow];
end;

function TCustomSpreadGridEh.GetCellDisplayText(ACol, ARow: Integer): String;
begin
  Result := VarToStr(SpreadArray[ACol, ARow].FValue);
//  Result := IntToStr(ACol) + ':' + IntToStr(ARow);
end;

function TCustomSpreadGridEh.GetVisibleColCount: Integer;
begin
  Result := HorzAxis.RolLastVisCel - HorzAxis.RolStartVisCel + 1;
end;

function TCustomSpreadGridEh.GetVisibleRowCount: Integer;
begin
  Result := VertAxis.RolLastVisCel - VertAxis.RolStartVisCel + 1;
end;

procedure TCustomSpreadGridEh.SetCellCanvasParams(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState);
begin
  Canvas.Font.Color := StyleServices.GetSystemColor(clWindowText);
  Canvas.Brush.Color := StyleServices.GetSystemColor(clMoneyGreen);
  if (ACol >= FixedColCount-FrozenColCount) and (ARow >= FixedRowCount-FrozenRowCount) then
    Canvas.Brush.Color := StyleServices.GetSystemColor(clWindow);

  if gdCurrent in State then
  begin
    Canvas.Font.Color := StyleServices.GetSystemColor(clHighlightText);
    if gdFocused in State then
      Canvas.Brush.Color := StyleServices.GetSystemColor(clHighlight)
    else
      Canvas.Brush.Color := StyleServices.GetSystemColor(clBtnShadow);
  end;

end;

procedure TCustomSpreadGridEh.DrawCell(ACol, ARow: Integer; ARect: TRect;
  State: TGridDrawState);
var
  AFillRect: TRect;
begin
  SetCellCanvasParams(ACol, ARow, ARect, State);

  AFillRect := ARect;
  Canvas.FillRect(AFillRect);
  Canvas.TextRect(AFillRect, AFillRect.Left+2, AFillRect.Top+2, GetCellDisplayText(ACol, ARow));
end;

function TCustomSpreadGridEh.CheckCellAreaDrawn(ACol, ARow: Integer): Boolean;
var
//  CheckCol, CheckRow: Integer;}
  i: Integer;
begin
{
  CheckCol := ACol;
  if CheckCol >= ColCount then
    CheckCol := CheckCol - ColCount + FixedColCount + VisibleColCount
  else if CheckCol >= FixedColCount then
    CheckCol := CheckCol - HorzAxis.RolStartVisCel;

  CheckRow := ARow;
  if CheckRow >= RowCount then
    CheckRow := CheckRow - RowCount + FixedRowCount + VisibleRowCount
  else if CheckRow >= FixedRowCount then
    CheckRow := CheckRow - VertAxis.RolStartVisCel;

  Result := FDrawenCellArr[FDrawColCount * CheckRow + CheckCol];}

  Result := False;
  for i := 0 to Length(FDrawenCellArr)-1 do
    if (FDrawenCellArr[i].X = ACol) and (FDrawenCellArr[i].Y = ARow) then
    begin
      Result := True;
      Break;
    end;
end;

procedure TCustomSpreadGridEh.SetCellDrawn(ACol, ARow: Integer);
var
//  DrawnCol, DrawnRow: Integer;
  NewPos: Integer;
begin
{  DrawnCol := ACol;
  if DrawnCol >= ColCount then
    DrawnCol := DrawnCol - ColCount + FixedColCount + VisibleColCount
  else if DrawnCol >= FixedColCount then
    DrawnCol := DrawnCol - HorzAxis.RolStartVisCel;

  DrawnRow := ARow;
  if DrawnRow >= RowCount then
    DrawnRow := DrawnRow - RowCount + FixedRowCount + VisibleRowCount
  else if DrawnRow >= FixedRowCount then
    DrawnRow := DrawnRow - VertAxis.RolStartVisCel;

  FDrawenCellArr[FDrawColCount * DrawnRow + DrawnCol] := True;}

  NewPos := Length(FDrawenCellArr);
  SetLength(FDrawenCellArr, NewPos+1);
  FDrawenCellArr[NewPos].X := ACol;
  FDrawenCellArr[NewPos].Y := ARow;
end;

procedure TCustomSpreadGridEh.DrawCellArea(ACol, ARow: Integer; ARect: TRect;
  State: TGridDrawState);
var
  SprCell: TSpreadGridCellEh;
begin
  if CheckCellAreaDrawn(ACol, ARow) then Exit;

  SprCell := FSpreadArray[ACol, ARow];

  if ((SprCell.FMergeColCount > 0) or (SprCell.FMergeRowCount > 0)) then
  begin
    DrawMasterCell(ACol, ARow, ARect, State, SprCell);
  end else if ((SprCell.FMasterMergeColOffset > 0) or (SprCell.FMasterMergeRowOffset > 0)) then
  begin
    DrawMasterForMergedCell(ACol, ARow, ARect, State);
  end else
    inherited DrawCellArea(ACol, ARow, ARect, State);
end;

procedure TCustomSpreadGridEh.DrawMasterCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; SpreadCell: TSpreadGridCellEh);
var
  i, j: Integer;
  MasterRect: TRect;
begin
  MasterRect := ARect;
  for i := ACol+1 to ACol + SpreadCell.FMergeColCount do
    MasterRect.Right := MasterRect.Right + ColWidths[i];
  for j := ARow+1 to ARow + SpreadCell.FMergeRowCount do
    MasterRect.Bottom := MasterRect.Bottom + RowHeights[j];

  for i := ACol to ACol + SpreadCell.FMergeColCount do
  begin
    for j := ARow to ARow + SpreadCell.FMergeRowCount do
      SetCellDrawn(i, j);
  end;
  DrawBordersForCellArea(ACol, ARow, MasterRect, State);
  DrawCell(ACol, ARow, MasterRect, State);
//  SetCellDrawn(ACol, ARow);
end;

procedure TCustomSpreadGridEh.DrawMasterForMergedCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState);
var
  SprCell, MasterCell: TSpreadGridCellEh;
  i: Integer;
  MasterCellPos: TGridCoord;
begin
  SprCell := FSpreadArray[ACol, ARow];
  MasterCellPos := GridCoord(ACol - SprCell.FMasterMergeColOffset, ARow - SprCell.FMasterMergeRowOffset);
  MasterCell := FSpreadArray[MasterCellPos.X, MasterCellPos.Y];
  for i := ACol - 1 downto MasterCellPos.X do
    ARect.Left := ARect.Left - ColWidths[i];
  for i := ARow - 1 downto MasterCellPos.Y do
    ARect.Top := ARect.Top - RowHeights[i];
  ARect.Right := ARect.Left + ColWidths[MasterCellPos.X];
  ARect.Bottom := ARect.Top + RowHeights[MasterCellPos.Y];
  DrawMasterCell(MasterCellPos.X, MasterCellPos.Y, ARect, State, MasterCell);
end;

procedure TCustomSpreadGridEh.MergeCells(BaseCellCol, BaseCellRow,
  MergeColCount, MergeRowCount: Integer);
var
  BaseCell, InMergeCell: TSpreadGridCellEh;
  i, j: Integer;
begin
  BaseCell := FSpreadArray[BaseCellCol, BaseCellRow];
  BaseCell.FMergeColCount := MergeColCount;
  BaseCell.FMergeRowCount := MergeRowCount;

  for i := 0 to MergeColCount do
  begin
    for j := 0 to MergeRowCount do
    begin
      if (i > 0) or (j > 0) then
      begin
        InMergeCell := FSpreadArray[BaseCellCol+i, BaseCellRow+j];
        InMergeCell.FMasterMergeColOffset := i;
        InMergeCell.FMasterMergeRowOffset := j;
      end;
    end;
  end;
end;

function TCustomSpreadGridEh.NeedBufferedPaint: Boolean;
begin
  Result := True;
end;

procedure TCustomSpreadGridEh.Paint;
{var
  i: Integer;}
begin
{  FDrawColCount := FixedColCount + (HorzAxis.RolLastVisCel - HorzAxis.RolStartVisCel + 1) + ContraColCount;
  FDrawRowCount := FixedRowCount + (VertAxis.RolLastVisCel - VertAxis.RolStartVisCel + 1)  + ContraRowCount;
  SetLength(FDrawenCellArr, FDrawColCount * FDrawRowCount);
  for i:= 0 to Length(FDrawenCellArr)-1 do
    FDrawenCellArr[i] := False;}
  SetLength(FDrawenCellArr, 0);

  inherited Paint;
end;

procedure TCustomSpreadGridEh.SetGridSize(ANewColCount, ANewRowCount: Integer);
begin
  ColCount := ANewColCount;
  RowCount := ANewRowCount;
  SetSpreadArraySize(ColCount, RowCount);
end;

procedure TCustomSpreadGridEh.SetSpreadArraySize(AColCount, ARowCount: Integer);
var
  c,r: Integer;
  OldCC, OldRC: Integer;
begin
  if (Length(FSpreadArray) > 0) then
  begin
    OldRC := Length(FSpreadArray[0]);
    OldCC  := Length(FSpreadArray);
    if ARowCount > OldRC then
    begin
      for c := 0 to OldCC-1 do
      begin
        SetLength(FSpreadArray[c], ARowCount);
//        SetLength(VisSpinGridArray[c], ARowCount);
        for r := OldRC to ARowCount-1 do
          FSpreadArray[c, r] := CreateSpreadGridCell(c, r);
      end;
    end else if ARowCount < OldRC then
    begin
      for c := 0 to OldCC-1 do
      begin
        for r := ARowCount to OldRC-1 do
        begin
          FSpreadArray[c, r].Free;
          FSpreadArray[c, r] := nil;
        end;
        SetLength(FSpreadArray[c], ARowCount);
//        SetLength(VisSpinGridArray[c], ARowCount);
      end;
    end;
  end;

  if AColCount > Length(FSpreadArray) then
  begin
    OldCC  := Length(FSpreadArray);
    SetLength(FSpreadArray, AColCount);
//    SetLength(VisSpinGridArray, AColCount);
    for c := OldCC to AColCount-1 do
    begin
      SetLength(FSpreadArray[c], ARowCount);
//      SetLength(VisSpinGridArray[c], ARowCount);
      for r := 0 to ARowCount-1 do
        FSpreadArray[c, r] := CreateSpreadGridCell(c, r);
    end;
  end else if AColCount < Length(FSpreadArray) then
  begin
    OldCC  := Length(FSpreadArray);
    for c := AColCount to OldCC-1 do
    begin
      for r := 0 to ARowCount-1 do
      begin
        FSpreadArray[c, r].Free;
        FSpreadArray[c, r] := nil;
      end;
    end;
    SetLength(FSpreadArray, AColCount);
//    SetLength(VisSpinGridArray, AColCount);
  end;

{?
  for c := 0 to AColCount-1 do
  begin
    for r := 0 to ARowCount-1 do
      FSpreadArray[c, r].Clear;
  end;
}
end;

end.
