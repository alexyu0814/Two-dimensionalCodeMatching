{*******************************************************}
{                                                       }
{                        EhLib v8.1                     }
{                                                       }
{                CalendarPicker Component               }
{                      Build 8.1.005                    }
{                                                       }
{   Copyright (c) 2014-2015 by Dmitry V. Bolshakov      }
{                                                       }
{*******************************************************}

{$I EhLib.Inc}

unit PlannerCalendarPickerEh;

interface

uses
{$IFDEF EH_LIB_17} System.Generics.Collections, {$ENDIF}
  Windows, SysUtils, Messages, Classes, Controls, Forms, StdCtrls, TypInfo,
  DateUtils, ExtCtrls, Buttons, Dialogs,
  Contnrs, Variants, Types, Themes, UxTheme,
{$IFDEF EH_LIB_17} System.UITypes, {$ENDIF}
{$IFDEF CIL}
  EhLibVCLNET,
  WinUtils,
{$ELSE}
  {$IFDEF FPC}
  EhLibLCL, LMessages, LCLType,
  {$ELSE}
  EhLibVCL,
  {$ENDIF}
{$ENDIF}
  PlannerDataEh, SpreadGridsEh, PlannersEh,
  GridsEh, ToolCtrlsEh, Graphics;

type
  TCustomCalendarMonthViewEh = class;
  TPlannerCalendarPickerEh = class;

{ TCalPickCellEh }

  TCalPickCellEh = class(TSpreadGridCellEh)
  private
    FHasPlanEvents: Boolean;
  public
    procedure Clear; override;
    property HasPlanEvents: Boolean read FHasPlanEvents write FHasPlanEvents;
  end;

  TCalendarMonthPickerButtonStyleEh = (cmpsPriorMonthEh, cmpsNextMonthEh);

{ TCalendarMonthPickerButtonEh }

  TCalendarMonthPickerButtonEh = class(TSpeedButtonEh)
  private
    FStyle: TCalendarMonthPickerButtonStyleEh;
    function GetMonthPicker: TCustomCalendarMonthViewEh;
    procedure SetStyle(const Value: TCalendarMonthPickerButtonStyleEh);

    procedure CMMouseEnter(var Message: TMessage); message CM_MOUSEENTER;
    procedure CMMouseLeave(var Message: TMessage); message CM_MOUSELEAVE;
  protected
    FMouseInControl: Boolean;
    procedure Paint; override;
  public
    constructor Create(AOwner: TComponent); override;
    procedure EditButtonDown(ButtonNum: Integer; var AutoRepeat: Boolean); override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;

    property MonthPicker: TCustomCalendarMonthViewEh read GetMonthPicker;
    property Style: TCalendarMonthPickerButtonStyleEh read FStyle write SetStyle;
  end;

  TCalendarMonthViewOptionEh = (clvFillTopDatesRowEh, clvFillBottomDatesRowEh,
    cmvShowPriorMonthButtonEh, cmvShowNextMonthButtonEh);
  TCalendarMonthViewOptionsEh = set of TCalendarMonthViewOptionEh;

{ TCustomCalendarMonthViewEh }

  TCustomCalendarMonthViewEh = class(TCustomSpreadGridEh, ISimpleChangeNotificationEh)
  private
    FBackButton: TCalendarMonthPickerButtonEh;
    FDate: TDateTime;
    FFirstWeekDayNum: Integer;
    FForwardButton: TCalendarMonthPickerButtonEh;
    FMaxDate: TDateTime;
    FMinDate: TDateTime;
    FOptions: TCalendarMonthViewOptionsEh;
    FStartDataRow: Integer;
    FStartDate: TDateTime;

    function AdjustDateToStartForGrid(ADateTime: TDateTime): TDateTime;
    function GetCalendarPicker: TPlannerCalendarPickerEh;
    function GetCell(ACol, ARow: Integer): TCalPickCellEh;
    function GetCurrentMonth: Integer;
    function GetEndDate: TDateTime;
    function GetMonth: TDateTime;
    function GetPlannerDataSource: TPlannerDataSourceEh;
    function GetPlannerControl: TPlannerControlEh;

    procedure ISimpleChangeNotificationEh.Change = PlannerDataSourceChange;
    procedure PlannerDataSourceChange(Sender: TObject); virtual;
    procedure PlannerDataSourceChanged;
    procedure SetDate(const Value: TDateTime);
    procedure SetMonth(const Value: TDateTime);
    procedure SetOptions(const Value: TCalendarMonthViewOptionsEh);

  protected
    FCellPosUpdating: Boolean;

    function CellToDate(ACol, ARow: Integer): TDateTime;
    function CreateSpreadGridCell(ACol, ARow: Integer): TSpreadGridCellEh; override;
    function DateToCell(ADateTime: TDateTime): TGridCoord;
    function IsCellCelected(ACol, ARow: Integer): Boolean;
    function IsDateCelected(ADateTime: TDateTime): Boolean;
    function SelectCell(ACol, ARow: Longint): Boolean; override;

    procedure BuildGrid;
    procedure CellMouseDown(const Cell: TGridCoord; Button: TMouseButton; Shift: TShiftState; const ACellRect: TRect; const GridMousePos, CellMousePos: TPoint); override;
    procedure CheckDrawCellBorder(ACol, ARow: Integer; BorderType: TGridCellBorderTypeEh; var IsDraw: Boolean; var BorderColor: TColor; var IsExtent: Boolean); override;
    procedure CurrentCellMoved(OldCurrent: TGridCoord); override;
    procedure DateChanged; virtual;
    procedure DrawBottomOutBoundaryData(ARect: TRect); override;
    procedure DrawCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); override;
    procedure DrawDayCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); virtual;
    procedure DrawEmptyCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); virtual;
    procedure DrawLeftOutBoundaryData(ARect: TRect); override;
    procedure DrawMonthNameRow(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); virtual;
    procedure DrawRightOutBoundaryData(ARect: TRect); override;
    procedure DrawTopOutBoundaryData(ARect: TRect); override;
    procedure DrawWeekDayNamesCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); virtual;
    procedure DrawWeekNum(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); virtual;
    procedure GetColRowByDate(ADate: TDateTime; var ACol, ARow: Integer); virtual;
    procedure KeyDown(var Key: Word; Shift: TShiftState); override;
    procedure LayoutChanged; virtual;
    procedure MonthChanged; virtual;
    procedure Resize; override;
    procedure UpdateCellPos; virtual;

    property Cell[ACol, ARow: Integer]: TCalPickCellEh read GetCell;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    function GetDateAtCell(ACol, ARow: Integer): TDateTime; virtual;
    function SuggestedWidth: Integer; virtual;
    function SuggestedHeight: Integer; virtual;

    procedure NextMonth; virtual;
    procedure PriorMonth; virtual;

    property CalendarPicker: TPlannerCalendarPickerEh read GetCalendarPicker;
    property Date: TDateTime read FDate write SetDate;
    property StartDate: TDateTime read FStartDate;
    property EndDate: TDateTime read GetEndDate;
    property MaxDate: TDateTime read FMaxDate write FMaxDate;
    property MinDate: TDateTime read FMinDate write FMinDate;
    property Month: TDateTime read GetMonth write SetMonth;
    property Options: TCalendarMonthViewOptionsEh read FOptions write SetOptions;
    property PlannerDataSource: TPlannerDataSourceEh read GetPlannerDataSource; //;FPlannerDataSource write SetPlannerDataSource;
    property PlannerControl: TPlannerControlEh read GetPlannerControl; //  write SetPlannerControl;
  end;

{ TCalendarMonthViewEh }

  TCalendarMonthViewEh = class(TCustomCalendarMonthViewEh)

  end;

{ TCalendarPickerLineParamsEh }

  TCalendarPickerLineParamsEh = class(TPersistent)
  private
    FCalendarPicker: TPlannerCalendarPickerEh;
    FBrightColor: TColor;
    FDarkColor: TColor;

    procedure SetBrightColor(const Value: TColor);
    procedure SetDarkColor(const Value: TColor);

  public
    constructor Create(ACalendarPicker: TPlannerCalendarPickerEh);
    destructor Destroy; override;

    function GetActualDarkColor: TColor; virtual;
    function GetActualBrightColor: TColor; virtual;

  published
    property DarkColor: TColor read FDarkColor write SetDarkColor default clDefault;
    property BrightColor: TColor read FBrightColor write SetBrightColor default clDefault;

  end;


{ TPlannerCalendarPickerEh }

  TPlannerCalendarPickerEh = class(TCustomControl, IPlannerControlChangeReceiverEh)
  private
    FBorderStyle: TBorderStyle;
    FCalsInCol: Integer;
    FCalsInRow: Integer;
    FCalViews: TList;
    FDate: TDateTime;
    FHolidaysFont: TFont;
    FHolidaysFontStored: Boolean;
    FLineParams: TCalendarPickerLineParamsEh;
    FPlannerControl: TPlannerControlEh;
    FStartMonth: TDateTime;

    function GetBorderSize: Integer;
    function GetCalViewCount: Integer;
    function GetCalViews(Index: Integer): TCalendarMonthViewEh;
    function GetStartMonth: TDateTime;
    function GetStopMonth: TDateTime;

    procedure IPlannerControlChangeReceiverEh.Change = PlannerDataSourceChange;
    procedure IPlannerControlChangeReceiverEh.GetViewPeriod = GetViewPeriod;
    {$IFDEF FPC}
    procedure SetBorderStyle(const Value: TBorderStyle); reintroduce;
    {$ELSE}
    procedure SetBorderStyle(const Value: TBorderStyle);
    {$ENDIF}
    procedure SetDate(const Value: TDateTime);
    procedure SetPlannerControl(const Value: TPlannerControlEh);
    procedure SetStartMonth(const Value: TDateTime);
    procedure PlannerDataSourceChange(Sender: TObject); virtual;

    {$IFDEF FPC}
    {$ELSE}
    procedure CMCtl3DChanged(var Message: TMessage); message CM_CTL3DCHANGED;
    {$ENDIF}
    procedure CMFontChanged(var Message: TMessage); message CM_FONTCHANGED;

    procedure WMEraseBkgnd(var Message: TWmEraseBkgnd); message WM_ERASEBKGND;

    procedure SetHolidaysFont(const Value: TFont);
    procedure SetHolidaysFontStored(const Value: Boolean);
    procedure SetLineParams(const Value: TCalendarPickerLineParamsEh);

  protected
    FViewPeriodStart: TDateTime;
    FViewPeriodFinish: TDateTime;

    function CreateCalendarView: TCalendarMonthViewEh;

    procedure CalViewMouseWheelDown(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: Boolean); virtual;
    procedure CalViewMouseWheelUp(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: Boolean); virtual;
    procedure CreateParams(var Params: TCreateParams); override;
    procedure CreateWnd; override;
    procedure EnsureDataForViewPeriod; virtual;
    procedure GetViewPeriod(var AStartDate, AEndDate: TDateTime); virtual;
    procedure HolidaysFontChanged(Sender: TObject); virtual;
    procedure InvalidateControls; virtual;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure Paint; override;
    procedure PlannerDataSourceChanged; virtual;
    procedure RefreshDefaultHolidaysFont; virtual;
    procedure ResetCalendarViews;
    procedure Resize; override;
    procedure SetCalViewsCount(Count: Integer);
    procedure SetDateInView; virtual;
    procedure UpdateControls; virtual;
    procedure UpdateDates;
    procedure UpdateStartMonth; virtual;

    property CalViewCount: Integer read GetCalViewCount;
    property CalViews[Index: Integer]: TCalendarMonthViewEh read GetCalViews;

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    function DefaultBrightLineColor: TColor; virtual;
    function DefaultDarkLineColor: TColor; virtual;
    function HasFocusControl: Boolean; virtual;
    function SuggestedWidth(CalendarViewCount: Integer): Integer; virtual;
    function SuggestedHeight(CalendarViewCount: Integer): Integer; virtual;

    procedure Scroll(MonthCount: Integer); virtual;
    procedure NextMonth; virtual;
    procedure PriorMonth; virtual;
    procedure NextDay; virtual;
    procedure PriorDay; virtual;
    procedure NextWeek; virtual;
    procedure PriorWeek; virtual;
    procedure InteractiveSetDate(const Value: TDateTime); virtual;

    property Date: TDateTime read FDate write SetDate;
    property StartMonth: TDateTime read GetStartMonth write SetStartMonth;
    property StopMonth: TDateTime read GetStopMonth;
  published
    property PlannerControl: TPlannerControlEh read FPlannerControl write SetPlannerControl;

    property Align;
    property Anchors;
    property BiDiMode;
    property BorderStyle: TBorderStyle read FBorderStyle write SetBorderStyle default bsSingle;
    property Color;
    property Constraints;
    {$IFDEF FPC}
    {$ELSE}
    property Ctl3D default False;
    {$ENDIF}
    property DragCursor;
    property DragKind;
    property DragMode;
    property Enabled;
    property Font;
    property HolidaysFont: TFont read FHolidaysFont write SetHolidaysFont;
    property HolidaysFontStored: Boolean read FHolidaysFontStored write SetHolidaysFontStored default False;
    property LineParams: TCalendarPickerLineParamsEh read FLineParams write SetLineParams;
//    property FooterColor;
    {$IFDEF FPC}
    {$ELSE}
    property ImeMode;
    property ImeName;
    {$ENDIF}
    property ParentBiDiMode;
    {$IFDEF FPC}
    {$ELSE}
    property ParentCtl3D;
    {$ENDIF}
    property ParentColor;
    property ParentFont;
    property ParentShowHint;
    property PopupMenu;
    property ShowHint;
    property TabOrder;
    property TabStop;
{$IFDEF EH_LIB_13}
    property Touch;
{$ENDIF}
    property Visible;

{$IFDEF EH_LIB_5}
    property OnContextPopup;
{$ENDIF}
    property OnDblClick;
    property OnDragDrop;
    property OnDragOver;
    property OnEndDock;
    property OnEndDrag;
    property OnEnter;
    property OnExit;
{$IFDEF EH_LIB_13}
    property OnGesture;
{$ENDIF}
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
    property OnMouseDown;
    property OnMouseMove;
    property OnMouseUp;
    property OnStartDock;
    property OnStartDrag;

  end;

implementation

{ TCalPickCellEh }

procedure TCalPickCellEh.Clear;
begin
  FHasPlanEvents := False;
end;

{ TCustomCalendarMonthPickerEh }

constructor TCustomCalendarMonthViewEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  inherited Options := [goDrawFocusSelectedEh];
  FixedColCount := 1;
  ColCount := 7 + FixedColCount;

  FixedRowCount := 2;
  RowCount := 8;
  FStartDataRow := 2;

  SetGridSize(ColCount, RowCount);
  HorzScrollBar.Visible := False;
  VertScrollBar.Visible := False;
  OutBoundaryData.LeftIndent := 10;
  OutBoundaryData.TopIndent := 2;
  OutBoundaryData.RightIndent := 10;
  OutBoundaryData.BottomIndent := 2;

  FFirstWeekDayNum := 2;
  FDate := Today;
  FStartDate := AdjustDateToStartForGrid(FDate);
  UpdateCellPos;

  FBackButton := TCalendarMonthPickerButtonEh.Create(Self);
  FBackButton.Parent := Self;
  FBackButton.Style := cmpsPriorMonthEh;

  FForwardButton := TCalendarMonthPickerButtonEh.Create(Self);
  FForwardButton.Parent := Self;
  FBackButton.Style := cmpsNextMonthEh;
  FOptions := [cmvShowPriorMonthButtonEh, cmvShowNextMonthButtonEh];
end;

destructor TCustomCalendarMonthViewEh.Destroy;
begin

  inherited Destroy;
end;

function TCustomCalendarMonthViewEh.CreateSpreadGridCell(ACol,
  ARow: Integer): TSpreadGridCellEh;
begin
  Result := TCalPickCellEh.Create;
end;

procedure TCustomCalendarMonthViewEh.CheckDrawCellBorder(ACol, ARow: Integer;
  BorderType: TGridCellBorderTypeEh; var IsDraw: Boolean;
  var BorderColor: TColor; var IsExtent: Boolean);
begin
  inherited CheckDrawCellBorder(ACol, ARow, BorderType, IsDraw, BorderColor, IsExtent);
//  Exit;
//  if (ACol = 0) and (ARow = 0) and (BorderType = cbtRightEh) then
//    IsDraw := False
//  else if (ACol = 0) {and (ARow > 0)} and (BorderType = cbtBottomEh) then
//    IsDraw := False;
end;

procedure TCustomCalendarMonthViewEh.DrawCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState);
begin
  if (ACol = 0) and (ARow = 0) then
    DrawMonthNameRow(ACol, ARow,  ARect, State)
  else if (ARow = 1) and (ACol = 0) then
    DrawEmptyCell(ACol, ARow,  ARect, State)
  else if (ACol = 0) then
    DrawWeekNum(ACol, ARow,  ARect, State)
  else if (ARow = 1) then
    DrawWeekDayNamesCell(ACol, ARow,  ARect, State)
  else
    DrawDayCell(ACol, ARow,  ARect, State);
end;

procedure TCustomCalendarMonthViewEh.DrawDayCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState);
var
  s: String;
  AddDats: Integer;
  DrawDate: TDateTime;
  ShowData: Boolean;
begin
  AddDats := 7 * (ARow-FStartDataRow) + (ACol-1);
  DrawDate := FStartDate + AddDats;
  s := FormatDateTime('D', DrawDate);
  Canvas.Brush.Color := StyleServices.GetSystemColor(Color);

  Canvas.Font := Font;
  Canvas.Font.Color := StyleServices.GetSystemColor(Font.Color);
{  if FStartDate + AddDats = Date then
    Canvas.Font.Style := [fsBold];}

  if (PlannerControl <> nil) and
     (PlannerControl.CurWorkingTimeCalendar <> nil) and
     not PlannerControl.CurWorkingTimeCalendar.IsWorkday(DrawDate)
  then
    Canvas.Font := CalendarPicker.HolidaysFont;

  ShowData := True;
  if GetCurrentMonth <> MonthOf(DrawDate) then
  begin
    if GetCurrentMonth < MonthOf(DrawDate) then
      ShowData := clvFillTopDatesRowEh in Options
    else if GetCurrentMonth > MonthOf(DrawDate) then
      ShowData := clvFillBottomDatesRowEh in Options;

    if Canvas.Font.Color = StyleServices.GetSystemColor(clRed) then
    begin
      Canvas.Font.Color := ChangeColorSaturation(Canvas.Font.Color, 64);
    end else
      Canvas.Font.Color := StyleServices.GetSystemColor(clGrayText);
  end;

  if (PlannerDataSource <> nil) and
     ( (PlannerDataSource.LoadedStartDate <> 0) or
       (PlannerDataSource.LoadedFinishDate <> 0)) then
  begin
    if (DrawDate < PlannerDataSource.LoadedStartDate) or
       (DrawDate >= PlannerDataSource.LoadedFinishDate)
    then
      Canvas.Font.Color := StyleServices.GetSystemColor(clGrayText);
  end;

  if ShowData and IsDateCelected(DrawDate) then
  begin
    if CalendarPicker.HasFocusControl then
    begin
      Canvas.Brush.Color := StyleServices.GetSystemColor(clHighlight);
      Canvas.Font.Color := StyleServices.GetSystemColor(clHighlightText);
    end else
    begin
      Canvas.Brush.Color := StyleServices.GetSystemColor(clBtnFace);
      Canvas.Font.Color := StyleServices.GetSystemColor(clWindowText);
    end;
  end;

  if Cell[ACol, ARow].HasPlanEvents then
    Canvas.Font.Style := Canvas.Font.Style + [fsBold];

  if ShowData then
    WriteTextEh(Canvas, ARect, True, 0, 0, s,
      taRightJustify, tlTop, False, False, 0, 0, False, True)
  else
    Canvas.FillRect(ARect);
end;

procedure TCustomCalendarMonthViewEh.DrawEmptyCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState);
begin
  Canvas.Brush.Color := StyleServices.GetSystemColor(Color);
  Canvas.FillRect(ARect);
end;

procedure TCustomCalendarMonthViewEh.DrawMonthNameRow(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState);
var
  s: String;
begin
  Canvas.Font := Font;
  Canvas.Font.Color := StyleServices.GetSystemColor(Font.Color);
  Canvas.Brush.Color := StyleServices.GetSystemColor(Color);
  s := FormatDateTime('mmmm', FStartDate + 7) + ' ' + FormatDateTime('yyyy', FStartDate + 7);
  WriteTextEh(Canvas, ARect, True, 0, 0, s,
    taCenter, tlCenter, False, False, 0, 0, False, True);
end;

procedure TCustomCalendarMonthViewEh.DrawBottomOutBoundaryData(ARect: TRect);
begin
  Canvas.Brush.Color := StyleServices.GetSystemColor(Color);
  Canvas.FillRect(ARect);
end;

procedure TCustomCalendarMonthViewEh.DrawLeftOutBoundaryData(ARect: TRect);
begin
  Canvas.Brush.Color := StyleServices.GetSystemColor(Color);
  Canvas.FillRect(ARect);
end;

procedure TCustomCalendarMonthViewEh.DrawRightOutBoundaryData(ARect: TRect);
begin
  Canvas.Brush.Color := StyleServices.GetSystemColor(Color);
  Canvas.FillRect(ARect);
end;

procedure TCustomCalendarMonthViewEh.DrawTopOutBoundaryData(ARect: TRect);
begin
  Canvas.Brush.Color := StyleServices.GetSystemColor(Color);
  Canvas.FillRect(ARect);
end;

procedure TCustomCalendarMonthViewEh.DrawWeekDayNamesCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState);
var
  DName: String;
  DNum: Integer;
  LineRect: TRect;
begin
  Canvas.Pen.Color := StyleServices.GetSystemColor(Color);
  DrawPolyline([Point(ARect.Left, ARect.Bottom-1), Point(ARect.Right, ARect.Bottom-1)]);
  Dec(ARect.Bottom);

  LineRect := ARect;
  if (ACol = 1) or (ACol = ColCount-1) then
  begin
    if ACol = 1
      then LineRect.Right := LineRect.Left + 3
      else LineRect.Left := LineRect.Right - 3;
    Canvas.Pen.Color := StyleServices.GetSystemColor(Color);
    DrawPolyline([Point(LineRect.Left, LineRect.Bottom-1), Point(LineRect.Right, LineRect.Bottom-1)]);
    if ACol = 1 then
    begin
      LineRect.Right := ARect.Right;
      LineRect.Left := LineRect.Left + 3;
    end else
    begin
      LineRect.Left := ARect.Left;
      LineRect.Right := LineRect.Right - 3;
    end;
  end;
  Canvas.Pen.Color := CalendarPicker.LineParams.GetActualBrightColor; // clBtnFace;  //GridLineColors.GetBrightColor;
  DrawPolyline([Point(LineRect.Left, LineRect.Bottom-1), Point(LineRect.Right, LineRect.Bottom-1)]);
  Dec(ARect.Bottom);

  DNum := (ACol-1) + FFirstWeekDayNum;
  if DNum > 7 then DNum := DNum - 7;

  DName := FormatSettings.ShortDayNames[DNum];

  Canvas.Font := Font;
  Canvas.Font.Color := StyleServices.GetSystemColor(Font.Color);
  Canvas.Brush.Color := StyleServices.GetSystemColor(Color);
  if DNum in [1, 7] then
  begin
    Canvas.Font := CalendarPicker.HolidaysFont;
    Canvas.Font.Color := StyleServices.GetSystemColor(CalendarPicker.HolidaysFont.Color);
  end;

  WriteTextEh(Canvas, ARect, True, 0, 0, DName,
    taRightJustify, tlCenter, False, False, 0, 0, False, True);
end;

procedure TCustomCalendarMonthViewEh.DrawWeekNum(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState);
var
  WeekNoDate: TDateTime;
  s: String;
  LineRect: TRect;
begin
//  Canvas.Pen.Color := CalendarPicker.LineParams.GetActualBrightColor; // clBtnFace;  //GridLineColors.GetBrightColor;
  Canvas.Pen.Color := StyleServices.GetSystemColor(Color);
  Canvas.Font := Font;
  Canvas.Font.Size := Canvas.Font.Size - 1;
  Canvas.Font.Color := StyleServices.GetSystemColor(Font.Color);

  DrawPolyline([Point(ARect.Right-1, ARect.Top), Point(ARect.Right-1, ARect.Bottom)]);
  Dec(ARect.Right);

  LineRect := ARect;
  if (ARow = FStartDataRow) or (ARow = RowCount-1) then
  begin
    if ARow = FStartDataRow
      then LineRect.Bottom := LineRect.Top+3
      else LineRect.Top := LineRect.Bottom-3;
    Canvas.Pen.Color := StyleServices.GetSystemColor(Color);
    DrawPolyline([Point(LineRect.Right-1, LineRect.Top), Point(LineRect.Right-1, LineRect.Bottom)]);
    if ARow = FStartDataRow then
    begin
      LineRect.Bottom := ARect.Bottom;
      LineRect.Top := LineRect.Top+3;
    end else
    begin
      LineRect.Top := ARect.Top;
      LineRect.Bottom := LineRect.Bottom-3;
    end;
  end;
  Canvas.Pen.Color := CalendarPicker.LineParams.GetActualBrightColor;
  DrawPolyline([Point(LineRect.Right-1, LineRect.Top), Point(LineRect.Right-1, LineRect.Bottom)]);
  Dec(ARect.Right);

  Canvas.Brush.Color := StyleServices.GetSystemColor(Color);
  WeekNoDate := FStartDate + 7 * (ARow-2);
  s := '';
  s := s + IntToStr(WeekOfTheYear(WeekNoDate));
  WriteTextEh(Canvas, ARect, True, 0, 0, s,
    taRightJustify, tlTop, False, False, 0, 0, False, True);
end;

function TCustomCalendarMonthViewEh.AdjustDateToStartForGrid(ADateTime: TDateTime): TDateTime;
var
  y, m, d: Word;
begin
  DecodeDate(ADateTime, y, m, d);
  Result := StartOfTheWeek(EncodeDate(y, m, 1));
end;

procedure TCustomCalendarMonthViewEh.BuildGrid;
var
  i: Integer;
  OutBoundWidth: Integer;
begin
  if not HandleAllocated then Exit;
  Canvas.Font := Font;
  DefaultRowHeight := Canvas.TextHeight('Wg') + Canvas.TextHeight('Wg') div 4;
  for i := 0 to RowCount-1 do
    RowHeights[i] := DefaultRowHeight;
  RowHeights[0] := RowHeights[0] * 2;
  RowHeights[1] := RowHeights[1] + 2;

  DefaultColWidth := Trunc(Canvas.TextWidth('Wg') * 1.5);
  for i := 0 to ColCount-1 do
    ColWidths[i] := DefaultColWidth;
  ColWidths[0] := Canvas.TextWidth('Wg') + 2;

  MergeCells(0,0, ColCount-1, 0);

  OutBoundWidth := ClientWidth -
    ((HorzAxis.FixedBoundary-HorzAxis.GridClientStart) + HorzAxis.RolLen + HorzAxis.ContraLen);
  OutBoundaryData.LeftIndent := OutBoundWidth div 2;
  OutBoundaryData.RightIndent := OutBoundWidth div 2 + OutBoundWidth mod 2;

  FBackButton.SetBounds(HorzAxis.GridClientStart, 10, FBackButton.Width, FBackButton.Height);
  FForwardButton.SetBounds(HorzAxis.GridClientStop-FForwardButton.Width, 10, FForwardButton.Width, FForwardButton.Height);
end;

function TCustomCalendarMonthViewEh.GetCalendarPicker: TPlannerCalendarPickerEh;
begin
  Result := Owner as TPlannerCalendarPickerEh;
end;

function TCustomCalendarMonthViewEh.GetCell(ACol,
  ARow: Integer): TCalPickCellEh;
begin
  Result := TCalPickCellEh(inherited Cell[ACol, ARow]);
end;

function TCustomCalendarMonthViewEh.GetCurrentMonth: Integer;
var
  y, m, d: Word;
begin
  DecodeDate(FStartDate + 7, y, m, d);
  Result := m;
end;

function TCustomCalendarMonthViewEh.GetDateAtCell(ACol,
  ARow: Integer): TDateTime;
begin
  Result := 0;
end;

function TCustomCalendarMonthViewEh.GetEndDate: TDateTime;
begin
  Result := FStartDate + 6 * 7;
end;

procedure TCustomCalendarMonthViewEh.NextMonth;
begin
  if CalendarPicker <> nil
    then CalendarPicker.Scroll(1)
    else Month := IncMonth(Month, 1);
end;

procedure TCustomCalendarMonthViewEh.PriorMonth;
begin
  if CalendarPicker <> nil
    then CalendarPicker.Scroll(-1)
    else Month := IncMonth(Month, -1);
end;

procedure TCustomCalendarMonthViewEh.Resize;
begin
  inherited Resize;
  BuildGrid;
end;

function TCustomCalendarMonthViewEh.GetMonth: TDateTime;
var
  y, m, d: Word;
begin
  DecodeDate(FStartDate + 7, y, m, d);
  Result := EncodeDate(y, m, 1);
end;

function TCustomCalendarMonthViewEh.GetPlannerControl: TPlannerControlEh;
begin
  if (CalendarPicker <> nil) and (CalendarPicker.PlannerControl <> nil)
    then Result := CalendarPicker.PlannerControl
    else Result := nil;
end;

function TCustomCalendarMonthViewEh.GetPlannerDataSource: TPlannerDataSourceEh;
begin
  if (CalendarPicker <> nil) and
     (CalendarPicker.PlannerControl <> nil) and
     (CalendarPicker.PlannerControl.PlannerDataSource <> nil)
  then
    Result := CalendarPicker.PlannerControl.PlannerDataSource
  else
    Result := nil;
end;

function TCustomCalendarMonthViewEh.IsCellCelected(ACol,
  ARow: Integer): Boolean;
begin
  Result := False;
  if PlannerControl <> nil then
  begin
    if PlannerControl.CoveragePeriodType = pcpDayEh  then
    begin
      Result := (ACol = Col) and (ARow = Row);
    end else if PlannerControl.CoveragePeriodType = pcpWeekEh then
    begin
      Result := (ARow = Row);
    end else if PlannerControl.CoveragePeriodType = pcpMonthEh then
    begin
      Result := True;
    end;
  end;
end;

function TCustomCalendarMonthViewEh.IsDateCelected(
  ADateTime: TDateTime): Boolean;
var
  AFromDate, AToDate: TDateTime;
begin
  Result := False;
  if PlannerControl <> nil then
  begin
    PlannerControl.CoveragePeriod(AFromDate, AToDate);
    if (ADateTime >= AFromDate) and (ADateTime < AToDate) then
      Result := True;
  end;
end;

procedure TCustomCalendarMonthViewEh.KeyDown(var Key: Word;
  Shift: TShiftState);
begin
//  inherited KeyDown(Key, Shift);
  case Key of
    VK_LEFT:
      begin
        CalendarPicker.PriorDay;
        CalendarPicker.UpdateControls;
      end;
    VK_RIGHT:
      begin
        CalendarPicker.NextDay;
        CalendarPicker.UpdateControls;
      end;
    VK_UP:
      begin
        CalendarPicker.PriorWeek;
        CalendarPicker.UpdateControls;
      end;
    VK_DOWN:
      begin
        CalendarPicker.NextWeek;
        CalendarPicker.UpdateControls;
      end;
    VK_NEXT:
      begin
        CalendarPicker.NextMonth;
        CalendarPicker.UpdateControls;
      end;
    VK_PRIOR:
      begin
        CalendarPicker.PriorMonth;
        CalendarPicker.UpdateControls;
      end;
  end;
end;

procedure TCustomCalendarMonthViewEh.SetMonth(const Value: TDateTime);
var
  y, m, d: Word;
  ny, nm, nd: Word;
  FirstMonthDay: TDateTime;
  LastMDay: Integer;
begin
  DecodeDate(Value, ny, nm, nd);
  FirstMonthDay := EncodeDate(ny, nm, 1);
  if FirstMonthDay <> Month then
  begin
    FStartDate := AdjustDateToStartForGrid(FirstMonthDay);

    DecodeDate(FDate, y, m, d);
    LastMDay := MonthDays[IsLeapYear(ny)][nm];
    if LastMDay < d then
      d := LastMDay;
    FDate := EncodeDate(ny, nm, d);
    UpdateCellPos;
    DateChanged;
    MonthChanged;
    Invalidate;
  end;
end;

procedure TCustomCalendarMonthViewEh.SetOptions(
  const Value: TCalendarMonthViewOptionsEh);
begin
  if FOptions <> Value then
  begin
    FOptions := Value;
    LayoutChanged;
    Invalidate;
  end;
end;

function TCustomCalendarMonthViewEh.SelectCell(ACol, ARow: Integer): Boolean;
var
  CelDate: TDateTime;
begin
  CelDate := CellToDate(ACol, ARow);
  if GetCurrentMonth = MonthOf(CelDate)
    then Result := True
    else Result := False;
end;

procedure TCustomCalendarMonthViewEh.SetDate(const Value: TDateTime);
var
  AMonth: TDateTime;
begin
  if FDate <> Value then
  begin
    AMonth := Month;
    FDate := Value;
    FStartDate := AdjustDateToStartForGrid(FDate);
    UpdateCellPos;
    DateChanged;
    if AMonth <> Month then
      MonthChanged;
    Invalidate;
  end;
end;

procedure TCustomCalendarMonthViewEh.UpdateCellPos;
var
  CellPos: TGridCoord;
begin
  if FCellPosUpdating then Exit;
  FCellPosUpdating := True;
  CellPos := DateToCell(Date);
  FocusCell(CellPos.X, CellPos.Y, True);
  FCellPosUpdating := False;
end;

procedure TCustomCalendarMonthViewEh.DateChanged;
begin
//  if PlannerControl <> nil then
//    PlannerControl.CurrentTime := Date;
end;

procedure TCustomCalendarMonthViewEh.MonthCHanged;
begin
  PlannerDataSourceChanged;
end;

function TCustomCalendarMonthViewEh.DateToCell(
  ADateTime: TDateTime): TGridCoord;
var
  DaysBn: Integer;
begin
  if FStartDate <= ADateTime then
  begin
    DaysBn := DaysBetween(FStartDate, ADateTime);
    Result.Y := DaysBn div 7 + FixedRowCount;
    Result.X := DaysBn mod 7 + FixedColCount;
  end else
  begin
    Result.Y := -1;
    Result.X := -1;
  end;
end;

procedure TCustomCalendarMonthViewEh.GetColRowByDate(ADate: TDateTime;
  var ACol, ARow: Integer);
var
  AGridCoord: TGridCoord;
begin
  AGridCoord := DateToCell(ADate);
  ACol := AGridCoord.X;
  ARow := AGridCoord.Y;
end;

procedure TCustomCalendarMonthViewEh.CellMouseDown(const Cell: TGridCoord;
  Button: TMouseButton; Shift: TShiftState; const ACellRect: TRect;
  const GridMousePos, CellMousePos: TPoint);
begin
  inherited;
end;

function TCustomCalendarMonthViewEh.CellToDate(ACol, ARow: Integer): TDateTime;
begin
  Result := FStartDate + (ARow - FixedRowCount) * 7 + (ACol - FixedColCount);
end;

procedure TCustomCalendarMonthViewEh.CurrentCellMoved(OldCurrent: TGridCoord);
begin
  inherited CurrentCellMoved(OldCurrent);
  if not FCellPosUpdating then
  begin
    if CalendarPicker <> nil then
      CalendarPicker.InteractiveSetDate(CellToDate(Col, Row));
  end;
end;

procedure TCustomCalendarMonthViewEh.PlannerDataSourceChanged;
var
  i: Integer;
  ACol, ARow: Integer;
  PlanItem: TPlannerDataItemEh;
  ADate: TDateTime;
begin
  for ACol := 0 to ColCount-1 do
    for ARow := 0 to RowCount-1 do
      Cell[ACol, ARow].HasPlanEvents := False;

  if PlannerDataSource <> nil then
    for i := 0 to PlannerDataSource.Count - 1 do
    begin
      PlanItem := PlannerDataSource[i];
      if ((PlanItem.EndTime >= FStartDate) and (PlanItem.StartTime < EndDate)) or
         ((PlanItem.StartTime >= FStartDate) and (PlanItem.StartTime < EndDate)) or
         ((PlanItem.EndTime >= FStartDate) and (PlanItem.EndTime < EndDate))
      then
      begin
        if PlanItem.EndTime < FStartDate
          then ADate := FStartDate
          else ADate := DateOf(PlanItem.StartTime);
        while True do
        begin
          GetColRowByDate(ADate, ACol, ARow);
          if (ACol >= 0) and (ARow >= 0) then
            Cell[ACol, ARow].HasPlanEvents := True;
          ADate := IncDay(ADate);
          if (ADate >= EndDate) or (ADate > PlanItem.EndTime) then
            Break;
        end;
      end;
    end;
  Invalidate;
end;

function TCustomCalendarMonthViewEh.SuggestedHeight: Integer;
begin
  Result := VertAxis.FixedBoundary +
           VertAxis.RolLen +
           VertAxis.ContraLen +
//           OutBoundaryData.TopIndent +
           OutBoundaryData.BottomIndent;
  if HandleAllocated then
    Result := Result + (Height - ClientHeight);
end;

function TCustomCalendarMonthViewEh.SuggestedWidth: Integer;
begin
  BuildGrid;
  Result := HorzAxis.FixedBoundary - HorzAxis.GridClientStart +
           HorzAxis.RolLen +
           HorzAxis.ContraLen + 10;
  if HandleAllocated then
    Result := Result + (Width - ClientWidth);
end;

procedure TCustomCalendarMonthViewEh.PlannerDataSourceChange(Sender: TObject);
begin
  if Sender is TPlannerDataSourceEh then
  begin
    if (csDestroying in ComponentState) or
       (csDestroying in PlannerDataSource.ComponentState)
    then
      Exit;

    PlannerDataSourceChanged;
  end else if Sender is TPlannerControlEh then
  begin
    Date := TPlannerControlEh(Sender).CurrentTime;
    Invalidate;
  end;
end;

procedure TCustomCalendarMonthViewEh.LayoutChanged;
begin
  Invalidate;
  FBackButton.Visible := cmvShowPriorMonthButtonEh in Options;
  FForwardButton.Visible := cmvShowNextMonthButtonEh in Options;
end;

{ TCalendarMonthPickerButtonEh }

constructor TCalendarMonthPickerButtonEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

procedure TCalendarMonthPickerButtonEh.EditButtonDown(ButtonNum: Integer;
  var AutoRepeat: Boolean);
begin
  inherited EditButtonDown(ButtonNum, AutoRepeat);
  AutoRepeat := True;
  if Style = cmpsPriorMonthEh then
    MonthPicker.NextMonth
  else
    MonthPicker.PriorMonth;
end;

procedure TCalendarMonthPickerButtonEh.CMMouseEnter(var Message: TMessage);
begin
  inherited;
  FMouseInControl := True;
  Invalidate;
end;

procedure TCalendarMonthPickerButtonEh.CMMouseLeave(var Message: TMessage);
begin
  inherited;
  FMouseInControl := False;
  Invalidate;
end;

procedure TCalendarMonthPickerButtonEh.MouseDown(Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  inherited MouseDown(Button, Shift, X, Y);
end;

function TCalendarMonthPickerButtonEh.GetMonthPicker: TCustomCalendarMonthViewEh;
begin
  Result := TCustomCalendarMonthViewEh(Owner);
end;

procedure TCalendarMonthPickerButtonEh.Paint;
var
  ArrawRect: TRect;
  CliRect: TRect;
  BF: Integer;
begin
  CliRect := Rect(0, 0, Width, Height);
  if FMouseInControl
    then Canvas.Brush.Color := StyleServices.GetSystemColor(clMenu)
    else Canvas.Brush.Color := StyleServices.GetSystemColor(MonthPicker.Color);
  Canvas.FillRect(CliRect);
  ArrawRect := Rect(0,0,4,7);
  ArrawRect := CenteredRect(CliRect, ArrawRect);
  Canvas.Pen.Color := StyleServices.GetSystemColor(clWindowText);
  if Style = cmpsPriorMonthEh then
  begin
    BF := 1
  end else
  begin
    BF := -1;
    ArrawRect.Left := ArrawRect.Right;
  end;

  Canvas.Polyline([Point(ArrawRect.Left, ArrawRect.Top), Point(ArrawRect.Left, ArrawRect.Bottom)]);
  Canvas.Polyline([Point(ArrawRect.Left+1*BF, ArrawRect.Top+1), Point(ArrawRect.Left+1*BF, ArrawRect.Bottom-1)]);
  Canvas.Polyline([Point(ArrawRect.Left+2*BF, ArrawRect.Top+2), Point(ArrawRect.Left+2*BF, ArrawRect.Bottom-2)]);
  Canvas.Polyline([Point(ArrawRect.Left+3*BF, ArrawRect.Top+3), Point(ArrawRect.Left+3*BF, ArrawRect.Bottom-3)]);
end;

procedure TCalendarMonthPickerButtonEh.SetStyle(
  const Value: TCalendarMonthPickerButtonStyleEh);
begin
  if FStyle <> Value then
  begin
    FStyle := Value;
    Invalidate;
  end;
end;

{ TPlannerCalendarPickerEh }

constructor TPlannerCalendarPickerEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FCalViews := TList.Create;
  SetCalViewsCount(1);
  Color := clWindow;
  ParentColor := False;
  FBorderStyle := bsSingle;
  {$IFDEF FPC}
  {$ELSE}
  Ctl3D := False;
  ParentCtl3D := False;
  {$ENDIF}
  FHolidaysFont := TFont.Create;
  FHolidaysFont.OnChange := HolidaysFontChanged;
  FLineParams := TCalendarPickerLineParamsEh.Create(Self);
end;

destructor TPlannerCalendarPickerEh.Destroy;
begin
  Destroying;
  PlannerControl := nil;
  SetCalViewsCount(0);
  FreeAndNil(FCalViews);
  FreeAndNil(FHolidaysFont);
  FreeAndNil(FLineParams);
  inherited Destroy;
end;

procedure TPlannerCalendarPickerEh.CalViewMouseWheelDown(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
begin
  Handled := True;
  if PlannerControl = nil then Exit;
  if PlannerControl.CoveragePeriodType = pcpDayEh  then
    InteractiveSetDate(Date + 1)
  else if PlannerControl.CoveragePeriodType = pcpWeekEh then
    InteractiveSetDate(IncWeek(Date, 1))
  else if PlannerControl.CoveragePeriodType = pcpMonthEh then
    InteractiveSetDate(IncMonth(Date, 1));
end;

procedure TPlannerCalendarPickerEh.CalViewMouseWheelUp(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
begin
  Handled := True;
  if PlannerControl = nil then Exit;
  if PlannerControl.CoveragePeriodType = pcpDayEh  then
    InteractiveSetDate(Date - 1)
  else if PlannerControl.CoveragePeriodType = pcpWeekEh then
    InteractiveSetDate(IncWeek(Date, -1))
  else if PlannerControl.CoveragePeriodType = pcpMonthEh then
    InteractiveSetDate(IncMonth(Date, -1));
end;

{$IFDEF FPC}
{$ELSE}
procedure TPlannerCalendarPickerEh.CMCtl3DChanged(var Message: TMessage);
begin
  if NewStyleControls and (BorderStyle = bsSingle) then
    RecreateWnd;
  inherited;
end;
{$ENDIF}

procedure TPlannerCalendarPickerEh.CMFontChanged(var Message: TMessage);
begin
  inherited;
  ResetCalendarViews;
  RefreshDefaultHolidaysFont;
end;

procedure TPlannerCalendarPickerEh.WMEraseBkgnd(var Message: TWmEraseBkgnd);
begin
  Message.Result := 1;
//  inherited;
end;

function TPlannerCalendarPickerEh.GetCalViewCount: Integer;
begin
  Result := FCalViews.Count;
end;

function TPlannerCalendarPickerEh.GetCalViews(Index: Integer): TCalendarMonthViewEh;
begin
  Result := TCalendarMonthViewEh(FCalViews[Index]);
end;

procedure TPlannerCalendarPickerEh.Paint;
var
  Rect: TRect;
begin
  Rect := GetClientRect;
  Canvas.Brush.Color := StyleServices.GetSystemColor(Color);
  Canvas.FillRect(Rect);
//  inherited Paint;
end;

procedure FixDate(Year, Month: Word; var Day: Word);
var
  dinm: Word;
begin
  dinm := DaysInAMonth(Year, Month);
  if dinm < Day then
    Day := dinm;
end;

procedure TPlannerCalendarPickerEh.NextDay;
begin
  InteractiveSetDate(Date+1);
end;

procedure TPlannerCalendarPickerEh.PriorDay;
begin
  InteractiveSetDate(Date-1);
end;

procedure TPlannerCalendarPickerEh.NextWeek;
begin
  InteractiveSetDate(IncWeek(Date));
end;

procedure TPlannerCalendarPickerEh.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) then
  begin
    if (AComponent is TPlannerControlEh) and (AComponent = PlannerControl) then
      PlannerControl := nil;
  end;
end;

procedure TPlannerCalendarPickerEh.PriorWeek;
begin
  InteractiveSetDate(IncWeek(Date,-1));
end;

procedure TPlannerCalendarPickerEh.NextMonth;
begin
  InteractiveSetDate(IncMonth(Date));
end;

procedure TPlannerCalendarPickerEh.PriorMonth;
begin
  InteractiveSetDate(IncMonth(Date,-1));
end;

procedure TPlannerCalendarPickerEh.Scroll(MonthCount: Integer);
begin
  StartMonth := IncMonth(StartMonth, MonthCount);
end;

function TPlannerCalendarPickerEh.CreateCalendarView: TCalendarMonthViewEh;
begin
  Result := TCalendarMonthViewEh.Create(Self);
  Result.Visible := False;
  Result.Parent := Self;
  Result.BorderStyle := bsNone;
//  Result.BorderStyle := bsSingle;
  Result.ParentColor := True;
  Result.OnMouseWheelDown := CalViewMouseWheelDown;
  Result.OnMouseWheelUp := CalViewMouseWheelUp;
end;

procedure TPlannerCalendarPickerEh.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  Params.Style := Params.Style or WS_CLIPCHILDREN;
  {$IFDEF FPC}
  {$ELSE}
  if BorderStyle = bsSingle then
    if NewStyleControls and Ctl3D
      then Params.ExStyle := Params.ExStyle or WS_EX_CLIENTEDGE
      else Params.Style := Params.Style or WS_BORDER;
  {$ENDIF}
end;

procedure TPlannerCalendarPickerEh.CreateWnd;
begin
  inherited CreateWnd;
  CalViews[0].HandleNeeded;
  ResetCalendarViews;
end;

procedure TPlannerCalendarPickerEh.ResetCalendarViews;
var
  DefCalWidth, DefCalHeight: Integer;
  i: Integer;
//  j: Integer;
  iPos, jPos: Integer;
  xOffset, yOffset: Integer;
  CalOps: TCalendarMonthViewOptionsEh;
  CalView: TCalendarMonthViewEh;
begin
  if not HandleAllocated then Exit;

  DefCalWidth := CalViews[0].SuggestedWidth;
  DefCalHeight := CalViews[0].SuggestedHeight;

  FCalsInRow := ClientWidth div DefCalWidth;
  if FCalsInRow = 0 then FCalsInRow := 1;

  FCalsInCol := ClientHeight div DefCalHeight;
  if FCalsInCol = 0 then FCalsInCol := 1;

  if ClientWidth > DefCalWidth
    then xOffset := ClientWidth mod DefCalWidth div 2
    else xOffset := 0;

  if ClientHeight > DefCalHeight
    then yOffset := ClientHeight mod DefCalHeight div 2
    else yOffset := 0;

  SetCalViewsCount(FCalsInRow*FCalsInCol);
  iPos := 0;
  jPos := 0;
//  j := 0;
  for i := 0 to CalViewCount-1 do
  begin
    CalView := CalViews[i];
    CalView.Visible := True;
    CalView.SetBounds(iPos+xOffset, jPos+yOffset, DefCalWidth, DefCalHeight);
    CalOps := [];
    if i = 0 then
      CalOps := CalOps + [cmvShowPriorMonthButtonEh];
    if (i = FCalsInRow-1) and (jPos = 0) then
      CalOps := CalOps + [cmvShowNextMonthButtonEh];
    CalView.Options := CalOps;
//    if j < FCalsInCol-1
//      then CalView.OutBoundaryData.BottomIndent := 0
//      else CalView.OutBoundaryData.BottomIndent := 10;

    Inc(iPos, DefCalWidth);
    if iPos + DefCalWidth > ClientWidth then
    begin
      Inc(jPos, DefCalHeight);
//      j := j + 1;
      iPos := 0;
    end;
  end;
  UpdateDates;
end;

procedure TPlannerCalendarPickerEh.Resize;
begin
  inherited Resize;
  ResetCalendarViews;
end;

procedure TPlannerCalendarPickerEh.SetBorderStyle(const Value: TBorderStyle);
begin
  if FBorderStyle <> Value then
  begin
    FBorderStyle := Value;
    {$IFDEF FPC}
    RecreateWnd(Self);
    {$ELSE}
    RecreateWnd;
    {$ENDIF}
  end;
end;

procedure TPlannerCalendarPickerEh.SetCalViewsCount(Count: Integer);
var
  i: Integer;
begin
  if Count < CalViewCount then
  begin
    for i := Count to CalViewCount-1 do
      CalViews[i].Free;
    FCalViews.Count := Count;
  end else if Count > CalViewCount then
  begin
    for i := 0 to Count-CalViewCount-1 do
      FCalViews.Add(CreateCalendarView);
//    ViewPeriodChanged;
  end;
end;

procedure TPlannerCalendarPickerEh.SetDate(const Value: TDateTime);
begin
  if FDate <> Value then
  begin
    FDate := Value;
    UpdateStartMonth;
    InvalidateControls;
  end;
end;

procedure TPlannerCalendarPickerEh.SetPlannerControl(const Value: TPlannerControlEh);
begin
  if FPlannerControl <> Value then
  begin
    if Assigned(FPlannerControl) then
      FPlannerControl.UnRegisterChanges(Self);
    FPlannerControl := Value;
    if Assigned(FPlannerControl) then
    begin
      FPlannerControl.RegisterChanges(Self);
      FPlannerControl.FreeNotification(Self);
    end;
    PlannerDataSourceChanged;
    EnsureDataForViewPeriod;
  end;
end;

function TPlannerCalendarPickerEh.GetBorderSize: Integer;
var
  Params: TCreateParams;
  R: TRect;
begin
  CreateParams(Params);
  SetRect(R, 0, 0, 0, 0);
  AdjustWindowRectEx(R, Params.Style, False, Params.ExStyle);
  Result := R.Bottom - R.Top;
end;

function TPlannerCalendarPickerEh.SuggestedHeight(CalendarViewCount: Integer): Integer;
begin
  if HandleAllocated then
  begin
    Result := CalViews[0].SuggestedHeight * CalendarViewCount;
    Result := Result + GetBorderSize;
  end else
    Result := 100;
end;

function TPlannerCalendarPickerEh.SuggestedWidth(CalendarViewCount: Integer): Integer;
begin
  if HandleAllocated then
  begin
    Result := CalViews[0].SuggestedWidth * CalendarViewCount;
    Result := Result + GetBorderSize;
  end else
    Result := 100;
end;

procedure TPlannerCalendarPickerEh.PlannerDataSourceChange(Sender: TObject);
begin
  PlannerDataSourceChanged;
end;

procedure TPlannerCalendarPickerEh.PlannerDataSourceChanged;
begin
  if PlannerControl <> nil then
  begin
    if Date <> PlannerControl.CurrentTime then
      Date := PlannerControl.CurrentTime
    else
      UpdateDates;
    InvalidateControls;
  end;
end;

procedure TPlannerCalendarPickerEh.SetDateInView;
var
  Year, Month, Day, CurDay: Word;
  NewDate: TDateTime;
begin
  if Date < StartMonth then
  begin
    DecodeDate(StartMonth, Year, Month, Day);
    CurDay := DayOf(Date);
    FixDate(Year, Month, CurDay);
    NewDate := EncodeDate(Year, Month, CurDay);
    InteractiveSetDate(NewDate);
  end else if Date >= StopMonth then
  begin
    DecodeDate(IncMonth(StopMonth, -1), Year, Month, Day);
    CurDay := DayOf(Date);
    FixDate(Year, Month, CurDay);
    NewDate := EncodeDate(Year, Month, CurDay);
    InteractiveSetDate(NewDate);
  end;
end;

procedure TPlannerCalendarPickerEh.UpdateDates;
var
  i: Integer;
  ADate: TDateTime;
  AViewPeriodStart, AViewPeriodFinish: TDateTime;
begin
  ADate := StartMonth;
  if CalViewCount = 0 then Exit;
  CalViews[0].Month := StartMonth;
  CalViews[0].PlannerDataSourceChanged;
  for i := 1 to CalViewCount-1 do
  begin
    ADate := IncMonth(ADate);
    CalViews[i].Month := ADate;
    CalViews[i].PlannerDataSourceChanged;
  end;
  GetViewPeriod(AViewPeriodStart, AViewPeriodFinish);
  if (AViewPeriodStart <> FViewPeriodStart) or (AViewPeriodFinish <> FViewPeriodFinish) then
  begin
    FViewPeriodStart := AViewPeriodStart;
    FViewPeriodFinish := AViewPeriodFinish;
    EnsureDataForViewPeriod;
  end;
end;

procedure TPlannerCalendarPickerEh.EnsureDataForViewPeriod;
begin
  if not (csLoading in ComponentState) and
     not (csDestroying in ComponentState) and
     (PlannerControl <> nil)
  then
    PlannerControl.EnsureDataForViewPeriod;
end;

procedure TPlannerCalendarPickerEh.UpdateStartMonth;
begin
  if Date < StartMonth then
    StartMonth := FDate
  else if Date >= StopMonth then
    StartMonth := IncMonth(FDate, -(CalViewCount-1));
end;

function TPlannerCalendarPickerEh.GetStartMonth: TDateTime;
begin
  Result := FStartMonth;
end;

function TPlannerCalendarPickerEh.GetStopMonth: TDateTime;
begin
  Result := IncMonth(StartMonth, CalViewCount);
end;

procedure TPlannerCalendarPickerEh.InteractiveSetDate(const Value: TDateTime);
begin
  if PlannerControl <> nil then
    PlannerControl.CurrentTime := Value;
end;

procedure TPlannerCalendarPickerEh.InvalidateControls;
var
  i: Integer;
begin
  for i := 0 to CalViewCount-1 do
    CalViews[i].Invalidate;
end;

procedure TPlannerCalendarPickerEh.UpdateControls;
var
  i: Integer;
begin
  for i := 0 to CalViewCount-1 do
    CalViews[i].Update;
end;

procedure TPlannerCalendarPickerEh.SetStartMonth(const Value: TDateTime);
var
  Year, Month, Day: Word;
  ANewStartMonth: TDateTime;
begin
  DecodeDate(Value, Year, Month, Day);
  ANewStartMonth := EncodeDate(Year, Month, 1);
  if FStartMonth <> ANewStartMonth then
  begin
    FStartMonth := ANewStartMonth;
    SetDateInView;
    UpdateDates;
  end;
end;

procedure TPlannerCalendarPickerEh.HolidaysFontChanged(Sender: TObject);
begin
  ResetCalendarViews;
  FHolidaysFontStored := True;
end;

procedure TPlannerCalendarPickerEh.RefreshDefaultHolidaysFont;
var
  Save: TNotifyEvent;
begin
  if HolidaysFontStored then Exit;
  Save := FHolidaysFont.OnChange;
  FHolidaysFont.OnChange := nil;
  try
    FHolidaysFont.Assign(Font);
    FHolidaysFont.Color := clRed;
  finally
    FHolidaysFont.OnChange := Save;
  end;
end;

procedure TPlannerCalendarPickerEh.SetHolidaysFont(const Value: TFont);
begin
  FHolidaysFont.Assign(Value);
end;

procedure TPlannerCalendarPickerEh.SetHolidaysFontStored(const Value: Boolean);
begin
  if FHolidaysFontStored <> Value then
  begin
    FHolidaysFontStored := Value;
    RefreshDefaultHolidaysFont;
  end;
end;

procedure TPlannerCalendarPickerEh.SetLineParams(
  const Value: TCalendarPickerLineParamsEh);
begin
  FLineParams.Assign(Value);
end;

function TPlannerCalendarPickerEh.HasFocusControl: Boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to CalViewCount-1 do
  begin
    Result := CalViews[i].IsActiveControl;
    if Result then Break;
  end;
end;

procedure TPlannerCalendarPickerEh.GetViewPeriod(var AStartDate, AEndDate: TDateTime);
begin
  AStartDate := CalViews[0].Month;
  AEndDate := IncMonth(CalViews[CalViewCount-1].Month);
end;

{ TCalendarPickerLineParamsEh }

constructor TCalendarPickerLineParamsEh.Create(
  ACalendarPicker: TPlannerCalendarPickerEh);
begin
  inherited Create;
  FCalendarPicker := ACalendarPicker;
  FBrightColor := clDefault;
  FDarkColor := clDefault;
end;

destructor TCalendarPickerLineParamsEh.Destroy;
begin
  inherited Destroy;
end;

function TCalendarPickerLineParamsEh.GetActualBrightColor: TColor;
begin
  if BrightColor = clDefault
    then Result := FCalendarPicker.DefaultBrightLineColor
    else Result := BrightColor;
end;

function TCalendarPickerLineParamsEh.GetActualDarkColor: TColor;
begin
  if BrightColor = clDefault
    then Result := FCalendarPicker.DefaultDarkLineColor
    else Result := DarkColor;
end;

procedure TCalendarPickerLineParamsEh.SetBrightColor(const Value: TColor);
begin
  if FBrightColor <> Value then
  begin
    FBrightColor := Value;
    FCalendarPicker.InvalidateControls;
  end;
end;

procedure TCalendarPickerLineParamsEh.SetDarkColor(const Value: TColor);
begin
  if FDarkColor <> Value then
  begin
    FDarkColor := Value;
    FCalendarPicker.InvalidateControls;
  end;
end;

function TPlannerCalendarPickerEh.DefaultBrightLineColor: TColor;
begin
  Result := StyleServices.GetSystemColor(clBtnFace);
end;

function TPlannerCalendarPickerEh.DefaultDarkLineColor: TColor;
begin
  Result := StyleServices.GetSystemColor(clGray);
end;

end.
