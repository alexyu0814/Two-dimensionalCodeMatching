{*******************************************************}
{                                                       }
{                        EhLib v8.1                     }
{                                                       }
{                     Planner Component                 }
{                      Build 8.1.008                    }
{                                                       }
{   Copyright (c) 2014-2016 by Dmitry V. Bolshakov      }
{                                                       }
{*******************************************************}

{$I EhLib.Inc}

unit PlannersEh;

interface

uses
{$IFDEF EH_LIB_17} System.Generics.Collections, System.UITypes, {$ENDIF}
  Windows, SysUtils, Messages, Classes, Controls, Forms, StdCtrls, TypInfo,
  DateUtils, ExtCtrls, Buttons, Dialogs, ImgList, GraphUtil, 
  Contnrs, Variants, Types, Themes, UxTheme,
{$IFDEF CIL}
  EhLibVCLNET,
  WinUtils,
{$ELSE}
  {$IFDEF FPC}
  EhLibLCL, LMessages, LCLType, Win32Extra,
  {$ELSE}
  EhLibVCL, PrintUtilsEh,
  {$ENDIF}
{$ENDIF}
  PlannerDataEh, SpreadGridsEh,
//  PrintPlannersEh,
  GridsEh, ToolCtrlsEh, Graphics;

type
  TCustomPlannerViewEh = class;
  TTimeSpanDisplayItemEh = class;
  TPlannerControlEh = class;
  TPlannerAxisTimelineViewEh = class;

  TDrawSpanItemDrawStateEh = set of (sidsSelectedEh, sidsFocusedEh);

  IPlannerControlChangeReceiverEh = interface
    ['{532A2D57-0ADB-49A3-8D8F-A300CA7C8D5B}']
    procedure Change(Sender: TObject);
    procedure GetViewPeriod(var AStartDate, AEndDate: TDateTime);
  end;

{ TDrawSpanItemArgsEh }

  TDrawSpanItemArgsEh = class(TPersistent)
  private
    FDrawState: TDrawSpanItemDrawStateEh;
    FText: String;
    FAlignment: TAlignment;
  published
    property DrawState: TDrawSpanItemDrawStateEh read FDrawState write FDrawState;
    property Text: String read FText write FText;
    property Alignment: TAlignment read FAlignment write FAlignment;
  end;

  TDrawSpanItemEventEh = procedure (PlannerControl: TPlannerControlEh;
    PlannerView: TCustomPlannerViewEh; SpanItem: TTimeSpanDisplayItemEh;
    Rect: TRect; DrawArgs: TDrawSpanItemArgsEh; var Processed: Boolean) of object;

  TTimeSpanBoundRectRelPosEh = (brrlWindowClientEh, brrlGridRolAreaEh);

  TPlannerResourceViewEh = record
    Resource: TPlannerResourceEh;
    GridOffset: Integer;
    GridStartAxisBar: Integer;
  end;

  TSpanInteractiveChange = (sichSpanLeftSizingEh, sichSpanRightSizingEh,
    sichSpanTopSizingEh, sichSpanButtomSizingEh, sichSpanMovingEh);
  TSpanInteractiveChanges = set of TSpanInteractiveChange;

  TTimeOrientationEh = (toHorizontalEh, toVerticalEh);
  TPropFillStyleEh = (fsDefaultEh, fsSolidEh, fsVerticalGradientEh, fsHorizontalGradientEh);

{ TDrawElementParamsEh }

  TDrawElementParamsEh = class(TPersistent)
  private
    FAltColor: TColor;
    FBorderColor: TColor;
    FColor: TColor;
    FFillStyle: TPropFillStyleEh;
    FFont: TFont;
    FFontStored: Boolean;
    FHue: TColor;

    function IsFontStored: Boolean;

    procedure SetAltColor(const Value: TColor);
    procedure SetBorderColor(const Value: TColor);
    procedure SetColor(const Value: TColor);
    procedure SetFillStyle(const Value: TPropFillStyleEh);
    procedure SetFont(const Value: TFont);
    procedure SetFontStored(const Value: Boolean);
    procedure SetHue(const Value: TColor);

  protected
    function DefaultFont: TFont; virtual;

    procedure NotifyChanges; virtual;
    procedure FontChanged(Sender: TObject);
    procedure RefreshDefaultFont;
    procedure AssignFontDefaultProps; virtual;

    function GetDefaultColor: TColor; virtual;
    function GetDefaultAltColor: TColor; virtual;
    function GetDefaultFillStyle: TPropFillStyleEh; virtual;
    function GetDefaultBorderColor: TColor; virtual;
    function GetDefaultHue: TColor; virtual;

    property Color: TColor read FColor write SetColor default clDefault;
    property Font: TFont read FFont write SetFont stored IsFontStored;
    property FontStored: Boolean read FFontStored write SetFontStored default False;
    property AltColor: TColor read FAltColor write SetAltColor default clDefault;
    property FillStyle: TPropFillStyleEh read FFillStyle write SetFillStyle default fsDefaultEh;
    property BorderColor: TColor read FBorderColor write SetBorderColor default clDefault;
    property Hue: TColor read FHue write SetHue default clDefault;

  public
    constructor Create;
    destructor Destroy; override;

    function GetActualColor: TColor; virtual;
    function GetActualAltColor: TColor; virtual;
    function GetActualFillStyle: TPropFillStyleEh; virtual;
    function GetActualBorderColor: TColor; virtual;
    function GetActualHue: TColor; virtual;
  end;

{ TDataBarsAreaEh }

  TDataBarsAreaEh = class(TPersistent)
  private
    FBarSize: Integer;
    FColor: TColor;
    FPlannerView: TCustomPlannerViewEh;

    procedure SetColor(const Value: TColor);
    procedure SetBarSize(const Value: Integer);

  protected
    function DefaultColor: TColor; virtual;
    function DefaultBarSize: Integer; virtual;

    procedure NotifyChanges; virtual;

    property PlannerView: TCustomPlannerViewEh read FPlannerView;

    function GetActualColor: TColor; virtual;
    function GetActualBarSize: Integer; virtual;

    property Color: TColor read FColor write SetColor default clDefault;
    property BarSize: Integer read FBarSize write SetBarSize default 0;

  public
    constructor Create(APlannerView: TCustomPlannerViewEh);
    destructor Destroy; override;
  end;

{ TDataBarsVertAreaEh }

  TDataBarsVertAreaEh = class(TDataBarsAreaEh)
  private
    function GetRowHeight: Integer;
    procedure SetRowHeight(const Value: Integer);
  protected
    function GetActualRowHeight: Integer; virtual;
  published
    property Color;
    property RowHeight: Integer read GetRowHeight write SetRowHeight default 0;
  end;

{ TDataBarsHorzAreaEh }

  TDataBarsHorzAreaEh = class(TDataBarsAreaEh)
  private
    function GetColWidth: Integer;
    procedure SetColWidth(const Value: Integer);
  protected
    function GetActualColWidth: Integer; virtual;
  published
    property Color;
    property ColWidth: Integer read GetColWidth write SetColWidth default 0;
  end;

{ TPlannerViewDrawElementEh }

  TPlannerViewDrawElementEh = class(TPersistent)
  private
    FSize: Integer;
    FColor: TColor;
    FFont: TFont;
    FVisible: Boolean;
    FVisibleStored: Boolean;
    FFontStored: Boolean;
    FPlannerView: TCustomPlannerViewEh;

    function GetVisible: Boolean;
    function IsFontStored: Boolean;

    procedure SetColor(const Value: TColor);
    procedure SetFont(const Value: TFont);
    procedure SetFontStored(const Value: Boolean);
    procedure SetSize(const Value: Integer);
    procedure SetVisible(const Value: Boolean);
    procedure SetVisibleStored(const Value: Boolean);

  protected
    function DefaultFont: TFont; virtual;
    function DefaultVisible: Boolean; virtual;
    function DefaultColor: TColor; virtual;
    function DefaultSize: Integer; virtual;
    function IsVisibleStored: Boolean; virtual;

    procedure NotifyChanges; virtual;
    procedure FontChanged(Sender: TObject); virtual;
    procedure RefreshDefaultFont; virtual;
    procedure AssignFontDefaultProps; virtual;

    property PlannerView: TCustomPlannerViewEh read FPlannerView;

    function GetActualColor: TColor; virtual;
    function GetActualSize: Integer; virtual;

    property Color: TColor read FColor write SetColor default clDefault;
    property Font: TFont read FFont write SetFont stored IsFontStored;
    property FontStored: Boolean read FFontStored write SetFontStored default False;
    property Size: Integer read FSize write SetSize default 0;
    property Visible: Boolean read GetVisible write SetVisible stored IsVisibleStored;
    property VisibleStored: Boolean read IsVisibleStored write SetVisibleStored stored False;

  public
    constructor Create(APlannerView: TCustomPlannerViewEh);
    destructor Destroy; override;
  end;

{ THoursBarAreaEh }

  THoursBarAreaEh = class(TPlannerViewDrawElementEh)
  protected
    function DefaultFont: TFont; override;
    function DefaultColor: TColor; override;
    function DefaultSize: Integer; override;
    function DefaultVisible: Boolean; override;
    procedure AssignFontDefaultProps; override;
  end;

{ THoursVertBarAreaEh }

  THoursVertBarAreaEh = class(THoursBarAreaEh)
  private
    function GetWidth: Integer;
    procedure SetWidth(const Value: Integer);
  published
    property Color;
    property Font;
    property FontStored;
    property Width: Integer read GetWidth write SetWidth default 0;
    property Visible;
    property VisibleStored;
  end;

{ THoursHorzBarAreaEh }

  THoursHorzBarAreaEh = class(THoursBarAreaEh)
  private
    function GetHeight: Integer;
    procedure SetHeight(const Value: Integer);
  published
    property Color;
    property Font;
    property FontStored;
    property Height: Integer read GetHeight write SetHeight default 0;
    property Visible;
    property VisibleStored;
  end;

{ TWeekBarAreaEh }

  TWeekBarAreaEh = class(THoursVertBarAreaEh)
  protected
    procedure AssignFontDefaultProps; override;
  end;

{ TDayNameAreaEh }

  TDayNameAreaEh = class(TPlannerViewDrawElementEh)
  protected
    function DefaultColor: TColor; override;
    function DefaultFont: TFont; override;
    function DefaultVisible: Boolean; override;
    function DefaultSize: Integer; override;
  end;

{ TDayNameVertAreaEh }

  TDayNameVertAreaEh = class(TDayNameAreaEh)
  private
    function GetHeight: Integer;
    procedure SetHeight(const Value: Integer);
  public
    function GetActualHeight: Integer; virtual;
  published
    property Color;
    property Font;
    property FontStored;
    property Height: Integer read GetHeight write SetHeight default 0;
    property Visible;
    property VisibleStored;
  end;

{ TDayNameHorzAreaEh }

  TDayNameHorzAreaEh = class(TDayNameAreaEh)
  private
    function GetWidth: Integer;
    procedure SetWidth(const Value: Integer);
  public
    function GetActualWidth: Integer; virtual;
  published
    property Color;
    property Font;
    property FontStored;
    property Width: Integer read GetWidth write SetWidth default 0;
    property Visible;
    property VisibleStored;
  end;

{ TResourceCaptionAreaEh }

  TResourceCaptionAreaEh = class(TPlannerViewDrawElementEh)
  private
    FEnhancedChanges: Boolean;
    function GetVisible: Boolean;

    procedure SetVisible(const Value: Boolean);
  protected
    function DefaultFont: TFont; override;
    function DefaultVisible: Boolean; override;
    function DefaultColor: TColor; override;
    function DefaultSize: Integer; override;

    procedure NotifyChanges; override;

    property Visible: Boolean read GetVisible write SetVisible stored IsVisibleStored;
  end;

{ TResourceVertCaptionAreaEh }

  TResourceVertCaptionAreaEh = class(TResourceCaptionAreaEh)
  private
    function GetHeight: Integer;
    procedure SetHeight(const Value: Integer);

  published
    property Color;
    property Font;
    property FontStored;
    property Height: Integer read GetHeight write SetHeight default 0;
    property Visible;
    property VisibleStored;
  end;

{ TResourceHorzCaptionAreaEh }

  TResourceHorzCaptionAreaEh = class(TResourceCaptionAreaEh)
  private
    function GetWidth: Integer;
    procedure SetWidth(const Value: Integer);

  published
    property Color;
    property Font;
    property FontStored;
    property Width: Integer read GetWidth write SetWidth default 0;
    property Visible;
    property VisibleStored;
  end;

{ TDatesBarAreaEh }

  TDatesBarAreaEh = class(TPlannerViewDrawElementEh)
  private
    function GetPlannerView: TPlannerAxisTimelineViewEh;

  protected
    function DefaultVisible: Boolean; override;
    function DefaultSize: Integer; override;

    property PlannerView: TPlannerAxisTimelineViewEh read GetPlannerView;
  end;

{ TDatesColAreaEh }

  TDatesColAreaEh = class(TDatesBarAreaEh)
  private
    function GetWidth: Integer;
    procedure SetWidth(const Value: Integer);

  published
    property Color;
    property Font;
    property FontStored;
    property Width: Integer read GetWidth write SetWidth default 0;
    property Visible;
    property VisibleStored;
  end;

{ TDatesRowAreaEh }

  TDatesRowAreaEh = class(TDatesBarAreaEh)
  private
    function GetHeight: Integer;
    procedure SetHeight(const Value: Integer);

  published
    property Color;
    property Font;
    property FontStored;
    property Height: Integer read GetHeight write SetHeight default 0;
    property Visible;
    property VisibleStored;
  end;

{ TInGridControlEh }

  TInGridControlEh = class(TPersistent)
  private
    FHorzLocating: TTimeSpanBoundRectRelPosEh;
    FVertLocating: TTimeSpanBoundRectRelPosEh;
    FGrid: TCustomPlannerViewEh;
    FBoundRect: TRect;
    FVisible: Boolean;

  protected
    procedure DblClick; virtual;

  public
    constructor Create(AGrid: TCustomPlannerViewEh);
    destructor Destroy; override;

    procedure Assign(Source: TPersistent); override;
    procedure GetInGridDrawRect(var ARect: TRect);

    property Visible: Boolean read FVisible write FVisible;
    property BoundRect: TRect read FBoundRect write FBoundRect;
    property PlannerView: TCustomPlannerViewEh read FGrid;
    property HorzLocating: TTimeSpanBoundRectRelPosEh read FHorzLocating;
    property VertLocating: TTimeSpanBoundRectRelPosEh read FVertLocating;
  end;

{ TTimeSpanDisplayItemEh }

  TTimeSpanDisplayItemEh = class(TInGridControlEh)
  private
    FEndTime: TDateTime;
    FGridColNum: Integer;
    FInCellCols: Integer;
    FInCellFromCol: Integer;
    FInCellFromRow: Integer;
    FInCellRows: Integer;
    FInCellToCol: Integer;
    FInCellToRow: Integer;
    FPlanItem: TPlannerDataItemEh;
    FStartGridRollPos: Integer;
    FStartTime: TDateTime;
    FStopGridRolPos: Integer;

  protected
    FDrawBackOutInfo: Boolean;
    FDrawForwardOutInfo: Boolean;
    FAlignment: TAlignment;
    FAllowedInteractiveChanges: TSpanInteractiveChanges;
    FTimeOrientation: TTimeOrientationEh;

  public
    constructor Create(AGrid: TCustomPlannerViewEh; APlanItem: TPlannerDataItemEh); reintroduce;
    destructor Destroy; override;

    procedure Assign(Source: TPersistent); override;
    procedure DblClick; override;

    property DrawBackOutInfo: Boolean read FDrawBackOutInfo;
    property DrawForwardOutInfo: Boolean read FDrawForwardOutInfo;

    property AllowedInteractiveChanges: TSpanInteractiveChanges read FAllowedInteractiveChanges;
    property EndTime: TDateTime read FEndTime write FEndTime;
    property GridColNum: Integer read FGridColNum;
    property InCellCols: Integer read FInCellCols;
    property InCellFromCol: Integer read FInCellFromCol;
    property InCellFromRow: Integer read FInCellFromRow;
    property InCellRows: Integer read FInCellRows;
    property InCellToCol: Integer read FInCellToCol;
    property InCellToRow: Integer read FInCellToRow;
    property PlanItem: TPlannerDataItemEh read FPlanItem;
    property StartGridRollPos: Integer read FStartGridRollPos;
    property StartTime: TDateTime read FStartTime write FStartTime;
    property StopGridRolPosl: Integer read FStopGridRolPos;
    property TimeOrientation: TTimeOrientationEh read FTimeOrientation;
    property Alignment: TAlignment read FAlignment write FAlignment;
  end;

{ TDummyTimeSpanDisplayItemEh }

  TDummyTimeSpanDisplayItemEh = class(TTimeSpanDisplayItemEh)
  private
    FStartTime: TDateTime;
    FEndTime: TDateTime;
    procedure SetEndTime(const Value: TDateTime);
    procedure SetStartTime(const Value: TDateTime);
  public
    procedure Assign(Source: TPersistent); override;

    property StartTime: TDateTime read FStartTime write SetStartTime;
    property EndTime: TDateTime read FEndTime write SetEndTime;
  end;

  TPlannerGridControlTypeEh = (pgctNextEventEh, pgctPriorEventEh);

{ TPlannerGridControlEh }

  TPlannerGridControlEh = class(TCustomControl)
  private
    FGrid: TCustomPlannerViewEh;
    FControlType: TPlannerGridControlTypeEh;
    FClickEnabled: Boolean;

  protected
    procedure Paint; override;

  public

    constructor Create(AGrid: TCustomPlannerViewEh); reintroduce;
    destructor Destroy; override;

    procedure Click; override;
    procedure Realign;

    property ControlType: TPlannerGridControlTypeEh read FControlType;
    property ClickEnabled: Boolean read FClickEnabled write FClickEnabled;
  end;

  TMoveHintWindow = class(THintWindow)
  private
    procedure WMEraseBkgnd(var Message: TWmEraseBkgnd); message WM_ERASEBKGND;
  end;

{ TPlannerGridLineParamsEh }

  TPlannerGridLineParamsEh = class(TGridLineColorsEh)
  private
    FPaleColor: TColor;
    procedure SetPaleColor(const Value: TColor);
    function GetPlannerView: TCustomPlannerViewEh;
  protected
    property PlannerView: TCustomPlannerViewEh read GetPlannerView;
  public
    constructor Create(AGrid: TCustomGridEh);

    function GetDarkColor: TColor; override;
    function GetBrightColor: TColor; override;
    function GetPaleColor: TColor; virtual;
    function DefaultPaleColor: TColor; virtual;
  published
    property DarkColor;
    property BrightColor;
    property PaleColor: TColor read FPaleColor write SetPaleColor default clDefault;
  end;

{ TPlannerViewCellDrawArgsEh }

  TPlannerViewCellDrawArgsEh = class(TObject)
  public
    Value: Variant;
    Text: String;
    Alignment: TAlignment;
    Layout: TTextLayout;
    WordWrap: Boolean;
    Orientation: TTextOrientationEh;
    HorzMargin: Integer;
    VertMargin: Integer;
    FontColor: TColor;
    FontName: TFontName;
    FontSize: Integer;
    FontStyle: TFontStyles;
    BackColor: TColor;
  end;

  TPlannerViewSpanHintParamsEh = class(TObject)
  public
    HintPos: TPoint;
    HintMaxWidth: Integer;
    HintColor: TColor;
    HintFont: TFont;
    CursorRect: TRect;
    ReshowTimeout: Integer;
    HideTimeout: Integer;
    HintStr: string;
  end;

  TPlannerDateRangeModeEh = (pdrmDayEh, pdrmWeekEh, pdrmMonthEh,
    pdrmEventsListEh, pdrmVertTimeBandEh, pdrmHorzTimeBandEh);
  TFirstWeekDayEh = (fwdSundayEh, fwdMondayEh);
  TDayNameFormatEh = (dnfLongFormatEh, dnfShortFormatEh, dnfNonEh);
  TRectangleSideEh = (rsLeftEh, rsRightEh, rsTopEh, rsBottomEh);
  TPlannerStateEh = (psNormalEh, psSpanLeftSizingEh, psSpanRightSizingEh,
    psSpanTopSizingEh, psSpanButtomSizingEh, psSpanMovingEh, psSpanTestMovingEh);
  TTimeUnitEh = (tuMSecEh, tuSecEh, tiMinEh, tuHourEh, tuDayEh, tuWeekEh, tuMonth, tuYear);
  TPlannerCoveragePeriodTypeEh = (pcpDayEh, pcpWeekEh, pcpMonthEh);
  TPlannerViewCellTypeEh = (pctTopLeftCellEh, pctDataCellEh, pctAlldayDataCellEh,
    pctResourceCaptionCellEh, pctInterResourceSpaceEh, pctDayNameCellEh,
    pctDateBarEh, pctDateCellEh, pctTimeCellEh, pctWeekNoCell);

  TPlannerViewDrawCellEventEh = procedure (PlannerView: TCustomPlannerViewEh;
    ACol, ARow: Integer; ARect: TRect; State: TGridDrawState;
    CellType: TPlannerViewCellTypeEh; ALocalCol, ALocalRow: Integer;
    var Processed: Boolean) of object;

  TPlannerViewSpanItemHintShowEventEh = procedure(PlannerControl: TPlannerControlEh;
    PlannerView: TCustomPlannerViewEh; CursorPos: TPoint; SpanRect: TRect;
    InSpanCursorPos: TPoint; SpanItem: TTimeSpanDisplayItemEh;
    Params: TPlannerViewSpanHintParamsEh; var Processed: Boolean) of object;

  TPlannerGridOprionEh = (pgoAddEventOnDoubleClickEh);
  TPlannerGridOprionsEh = set of TPlannerGridOprionEh;

{ TCustomPlannerViewEh }

  TCustomPlannerViewEh = class(TCustomSpreadGridEh)
  private
    FActiveMode: Boolean;
    FCurrentTime: TDateTime;
    FFirstWeekDay: TFirstWeekDayEh;
    FFirstWeekDayNum: Integer;
    FOptions: TPlannerGridOprionsEh;
    FRangeMode: TPlannerDateRangeModeEh;
    FResourceCellFillColor: TColor;
    FSelectedPlanItem: TPlannerDataItemEh;
    FSpanItems: TObjectList;
//    FPlannerDataSource: TPlannerDataSourceEh;
    FHoursBarArea: THoursBarAreaEh;
    FResourceCaptionArea: TResourceCaptionAreaEh;
    FDayNameArea: TDayNameAreaEh;
    FDataBarsArea: TDataBarsAreaEh;
    FOnDrawCell: TPlannerViewDrawCellEventEh;
    FOnSelectionChanged: TNotifyEvent;
    FOnSpanItemHintShow: TPlannerViewSpanItemHintShowEventEh;
    FHintFont: TFont;

    function CheckStartSpanMove(MousePos: TPoint): Boolean;
    function GetGridIndex: Integer;
    function GetSpanItems(Index: Integer): TTimeSpanDisplayItemEh;
    function GetSpanItemsCount: Integer;
    function GridCoordToDataCoord(const AGridCoord: TGridCoord): TGridCoord;
    function GetGridLineParams: TPlannerGridLineParamsEh;
    function GetPlannerDataSource: TPlannerDataSourceEh;

    procedure ClearSpanItems;
    procedure DrawSpanItem(SpanItem: TTimeSpanDisplayItemEh; DrawRect: TRect);
    procedure SetActiveMode(const Value: Boolean);
    procedure SetCurrentTime(const Value: TDateTime);
    procedure SetDataBarsArea(const Value: TDataBarsAreaEh);
    procedure SetDayNameArea(const Value: TDayNameAreaEh);
    procedure SetFirstWeekDay(const Value: TFirstWeekDayEh);
    procedure SetGridIndex(const Value: Integer);
    procedure SetGridLineParams(const Value: TPlannerGridLineParamsEh);
    procedure SetHoursBarArea(const Value: THoursBarAreaEh);
    procedure SetOptions(const Value: TPlannerGridOprionsEh);
    procedure SetRangeMode(const Value: TPlannerDateRangeModeEh);
    procedure SetResourceCaptionArea(const Value: TResourceCaptionAreaEh);
    procedure SetSelectedSpanItem(const Value: TPlannerDataItemEh);
    procedure SetStartDate(const Value: TDateTime);
//    procedure SetPlannerDataSource(const Value: TPlannerDataSourceEh);
    procedure StartSpanMove(SpanItem: TTimeSpanDisplayItemEh; MousePos: TPoint);

    procedure CMFontChanged(var Message: TMessage); message CM_FONTCHANGED;
    procedure CMHintShow(var Message: TCMHintShow); message CM_HINTSHOW;

    procedure WMSetCursor(var Msg: TWMSetCursor); message WM_SETCURSOR;

  protected
    FBarsPerRes: Integer;
    FDataColsOffset: Integer;
    FDataRowsOffset: Integer;
    FDayNameFormat: TDayNameFormatEh;
    FDefaultTimeSpanBoxHeight: Integer;
    FDummyPlanItem: TPlannerDataItemEh;
    FDummyCheckPlanItem: TPlannerDataItemEh;
    FDummyPlanItemFor: TPlannerDataItemEh;
    FGridControls: TObjectList;
    FInitedInAllDaysArea: Boolean;
    FMouseDownPos: TPoint;
    FMoveHintWindow: THintWindow;
    FPlannerState: TPlannerStateEh;
    FResourceAxisPos: Integer;
    FResourcesView: array of TPlannerResourceViewEh;
    FShowUnassignedResource: Boolean;
    FSlideDirection: TGridScrollDirection;
    FSpanFrameColor: TColor;
    FSpanMoveSlidingSpeed: Integer;
    FSpanMoveSlidingTimer: TTimer;
    FStartDate: TDateTime;
    FDayNameBarPos: Integer;
    FTopLeftSpanShift: TSize;
    FHoursBarIndex: Integer;
    FDrawSpanItemArgs: TDrawSpanItemArgsEh;
    FDrawCellArgs: TPlannerViewCellDrawArgsEh;
    FTopGridLineIndex: Integer;
    FIgnorePlannerDataSourceChanged: Boolean;
    FSelectedFocusedCellColor: TColor;
    FSelectedUnfocusedCellColor: TColor;

    function AdjustDate(const Value: TDateTime): TDateTime; virtual;
    function IsWorkingDay(const Value: TDateTime): Boolean; virtual;
    function IsWorkingTime(const Value: TDateTime): Boolean; virtual;
    function DrawLongDayNames: Boolean; virtual;
    function DrawMonthDayWithWeekDayName: Boolean; virtual;
    function CellToDateTime(ACol, ARow: Integer): TDateTime; virtual;
    function NewItemParams(var StartTime, EndTime: TDateTime; var Resource: TPlannerResourceEh): Boolean; virtual;
    function GetResourceAtCell(ACol, ARow: Integer): TPlannerResourceEh; virtual;
    function GetResourceViewAtCell(ACol, ARow: Integer): Integer; virtual;
    function GetCoveragePeriodType: TPlannerCoveragePeriodTypeEh; virtual;
    function SpanItemByPlanItem(APlanItem: TPlannerDataItemEh): TTimeSpanDisplayItemEh; virtual;
    function EventNavBoxBorderColor: TColor; virtual;
    function EventNavBoxColor: TColor; virtual;
    function EventNavBoxFont: TFont; virtual;

    function CheckHitSpanItem(Button: TMouseButton; Shift: TShiftState; X, Y: Integer): TTimeSpanDisplayItemEh; virtual;
    function CheckSpanItemSizing(MousePos: TPoint; out SpanItem: TTimeSpanDisplayItemEh; var Side: TRectangleSideEh): Boolean; virtual;
    function ClientToGridRolPos(Pos: TPoint): TPoint; virtual;
    function CreateGridLineColors: TGridLineColorsEh; override;
    function GetPlannerControl: TPlannerControlEh;
    function InteractiveChangeAllowed: Boolean; virtual;
    function CheckPlanItemForRead(APlanItem: TPlannerDataItemEh): Boolean; virtual;
    function AddSpanItem(APlanItem: TPlannerDataItemEh): TTimeSpanDisplayItemEh; virtual;
    function ShowTopGridLine: Boolean; virtual;
    function TopGridLineCount: Integer; virtual;

    function GetNextEventAfter(ADateTime: TDateTime): Variant; virtual;
    function GetNextEventBefore(ADateTime: TDateTime): Variant; virtual;

    procedure BuildGridData; virtual;
    procedure CheckSetDummyPlanItem(Item, NewItem: TPlannerDataItemEh); virtual;
    procedure CurrentCellMoved(OldCurrent: TGridCoord); override;
    procedure EnsureDataForPeriod(AStartDate, AEndDate: TDateTime); virtual;
    procedure GridLayoutChanged; virtual;
    procedure GroupSpanItems; virtual;
    procedure MakeSpanItems; virtual;
    procedure NotifyPlanItemChanged(Item, OldItem: TPlannerDataItemEh); virtual;
    procedure ProcessedSpanItems;
    procedure ReadPlanItem(APlanItem: TPlannerDataItemEh); virtual;
    procedure ResetAllData;
    procedure ResetLoadedTimeRange;
    procedure ResetResviewArray; virtual;
    procedure SelectionChanged; reintroduce; virtual;

    procedure CalcRectForInCellCols(SpanItem: TTimeSpanDisplayItemEh; var DrawRect: TRect); virtual;
    procedure CalcRectForInCellRows(SpanItem: TTimeSpanDisplayItemEh; var DrawRect: TRect); virtual;
    procedure GetViewPeriod(var AStartDate, AEndDate: TDateTime); virtual;
    procedure InitSpanItem(ASpanItem: TTimeSpanDisplayItemEh); virtual;
    procedure SetDisplayPosesSpanItems; virtual;
    procedure SetGroupPosesSpanItems(Resource: TPlannerResourceEh); virtual;
    procedure UpdateDefaultTimeSpanBoxHeight; virtual;
    procedure StartDateChanged; virtual;
    procedure UpdateSelectedPlanItem; virtual;

    procedure RangeModeChanged; virtual;
    procedure SortPlanItems; virtual;
    procedure GotoNextEventInTheFuture; virtual;
    procedure GotoPriorEventInThePast; virtual;
    procedure RealignGridControl(AGridControl: TPlannerGridControlEh); virtual;

    procedure ResetDayNameFormat(LongDayFacor, ShortDayFacor: Double); virtual;

    procedure PlannerDataSourceChanged;
    procedure CancelMode; override;

    procedure Resize; override;
    procedure Paint; override;
    procedure PaintSpanItems; virtual;
    procedure SetPaintColors; override;

    procedure DrawCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); override;
    procedure GetCellType(ACol, ARow: Integer; var CellType: TPlannerViewCellTypeEh; var ALocalCol, ALocalRow: Integer); virtual;

    procedure ClearGridCells;

    procedure SetCellCanvasParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); override;
    procedure InitSpanItemMoving(SpanItem: TTimeSpanDisplayItemEh; MousePos: TPoint); virtual;

    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
    procedure KeyDown(var Key: Word; Shift: TShiftState); override;
    procedure CellMouseClick(const Cell: TGridCoord; Button: TMouseButton; Shift: TShiftState; const ACellRect: TRect; const GridMousePos, CellMousePos: TPoint); override;

    procedure UpdateDummySpanItemSize(MousePos: TPoint); virtual;
    procedure StopPlannerState(Accept: Boolean; X, Y: Integer);
    procedure CreateGridControls; virtual;
    procedure DeleteGridControls; virtual;
    procedure RealignGridControls; virtual;
    procedure ResetGridControlsState; virtual;
    procedure SetPlannerState(ANewState: TPlannerStateEh); virtual;
    procedure PlannerStateChanged(AOldState: TPlannerStateEh); virtual;
    procedure GetWeekDayNamesParams(ACol, ARow, ALocalCol, ALocalRow: Integer; var WeekDayNum: Integer; var WeekDayName: String); virtual;

    function CanSpanMoveSliding(MousePos: TPoint; var ASpeed: Integer; var ASlideDirection: TGridScrollDirection): Boolean;
    procedure ShowMoveHintWindow(APlanItem: TPlannerDataItemEh; MousePos: TPoint); virtual;
    procedure HideMoveHintWindow; virtual;
    procedure SlidingTimerEvent(Sender: TObject); virtual;
    procedure StopSpanMoveSliding; virtual;
    procedure StartSpanMoveSliding(ASpeed: Integer; ASlideDirection: TGridScrollDirection); virtual;
    procedure ReadState(Reader: TReader); override;
    procedure SetParent(AParent: TWinControl); override;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure CreateWnd; override;

    function IsResourceCaptionNeedVisible: Boolean; virtual;
    function IsDayNameAreaNeedVisible: Boolean; virtual;
    function ResourcesCount: Integer; virtual;
    function DefaultHoursBarSize: Integer; virtual;
    function GetDayNameAreaDefaultSize: Integer; virtual;
    function GetDayNameAreaDefaultColor: TColor; virtual;
    function GetResourceCaptionAreaDefaultSize: Integer; virtual;
    function CreateHoursBarArea: THoursBarAreaEh; virtual;
    function CreateDayNameArea: TDayNameAreaEh; virtual;
    function CreateDataBarsArea: TDataBarsAreaEh; virtual;
    function CreateResourceCaptionArea: TResourceCaptionAreaEh; virtual;
    function GetDataBarsAreaDefaultBarSize: Integer; virtual;

    property HoursBarArea: THoursBarAreaEh read FHoursBarArea write SetHoursBarArea;
    property ResourceCaptionArea: TResourceCaptionAreaEh read FResourceCaptionArea write SetResourceCaptionArea;
    property DayNameArea: TDayNameAreaEh read FDayNameArea write SetDayNameArea;
    property DataBarsArea: TDataBarsAreaEh read FDataBarsArea write SetDataBarsArea;
    property RangeMode: TPlannerDateRangeModeEh read FRangeMode write SetRangeMode default pdrmDayEh;
    property ResourceCellFillColor: TColor read FResourceCellFillColor;

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    function DeletePrompt: Boolean;
    function NextDate: TDateTime; virtual;
    function PriorDate: TDateTime; virtual;
    function AppendPeriod(ATime: TDateTime; Periods: Integer): TDateTime; virtual;
    function GetPeriodCaption: String; virtual;

    procedure NextPeriod; virtual;
    procedure PriorPeriod; virtual;
    procedure CoveragePeriod(var AFromTime, AToTime: TDateTime); virtual;

    procedure DefaultDrawSpanItem(SpanItem: TTimeSpanDisplayItemEh; const ARect: TRect; DrawArgs: TDrawSpanItemArgsEh); virtual;
    procedure DrawSpanItemBackgroud(SpanItem: TTimeSpanDisplayItemEh; const ARect: TRect; DrawArgs: TDrawSpanItemArgsEh); virtual;
    procedure DrawSpanItemContent(SpanItem: TTimeSpanDisplayItemEh; const ARect: TRect; DrawArgs: TDrawSpanItemArgsEh); virtual;
    procedure DrawSpanItemSurround(SpanItem: TTimeSpanDisplayItemEh; var ARect: TRect; DrawArgs: TDrawSpanItemArgsEh); virtual;

    procedure DefaultDrawPlannerViewCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; CellType: TPlannerViewCellTypeEh; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure DefaultFillSpanItemHintShowParams(CursorPos: TPoint; SpanRect: TRect; InSpanCursorPos: TPoint; SpanItem: TTimeSpanDisplayItemEh; Params: TPlannerViewSpanHintParamsEh); virtual;

    procedure GetDrawCellParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; CellType: TPlannerViewCellTypeEh; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure GetAlldayDataCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure GetDataCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure GetDateBarDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure GetDateCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure GetDayNamesCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure GetInterResourceCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure GetResourceCaptionCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure GetTimeCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure GetTopLeftCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure GetWeekNoCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;

    procedure DrawAlldayDataCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure DrawDataCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure DrawDateBar(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure DrawDateCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure DrawDayNamesCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure DrawInterResourceCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure DrawResourceCaptionCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure DrawTimeCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure DrawTopLeftCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure DrawWeekNoCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;

    property Col;
    property Row;
    property Canvas;

    property FirstWeekDay: TFirstWeekDayEh read FFirstWeekDay write SetFirstWeekDay default fwdMondayEh;
    property SpanItems[Index: Integer]: TTimeSpanDisplayItemEh read GetSpanItems;
    property SpanItemsCount: Integer read GetSpanItemsCount;
    property PlannerDataSource: TPlannerDataSourceEh read GetPlannerDataSource;// write SetPlannerDataSource;
    property ActiveMode: Boolean read FActiveMode write SetActiveMode;

    property StartDate: TDateTime read FStartDate write SetStartDate;
    property CurrentTime: TDateTime read FCurrentTime write SetCurrentTime;
    property SelectedPlanItem: TPlannerDataItemEh read FSelectedPlanItem write SetSelectedSpanItem;
    property PlannerControl: TPlannerControlEh read GetPlannerControl;
    property CoveragePeriodType: TPlannerCoveragePeriodTypeEh read GetCoveragePeriodType;
    property Options: TPlannerGridOprionsEh read FOptions write SetOptions default [pgoAddEventOnDoubleClickEh];

    property ViewIndex: Integer read GetGridIndex write SetGridIndex stored False;
    property GridLineParams: TPlannerGridLineParamsEh read GetGridLineParams write SetGridLineParams;

    property OnDrawCell: TPlannerViewDrawCellEventEh read FOnDrawCell write FOnDrawCell;
    property OnSelectionChanged: TNotifyEvent read FOnSelectionChanged write FOnSelectionChanged;
    property OnSpanItemHintShow: TPlannerViewSpanItemHintShowEventEh read FOnSpanItemHintShow write FOnSpanItemHintShow;
  end;

  TCustomPlannerGridClassEh = class of TCustomPlannerViewEh;

{ TPlannerWeekViewEh }

  TPlannerWeekViewEh = class(TCustomPlannerViewEh)
  private
    FMinDayColWidth: Integer;
    FWorkingTimeEnd: TTime;
    FWorkingTimeStart: TTime;
    FShowWorkingTimeOnly: Boolean;
    FGridStartWorkingTime: TTime;
    FGridWorkingTimeLength: TDateTime;

    function GetAllDayListCount: Integer;
    function GetAllDayListItem(Index: Integer): TTimeSpanDisplayItemEh;
    function GetColsStartTime(ADayCol: Integer): TDateTime;
    function GetDataBarsArea: TDataBarsVertAreaEh;
    function GetDayNameArea: TDayNameVertAreaEh;
    function GetHoursColArea: THoursVertBarAreaEh;
    function GetInDayListCount: Integer;
    function GetInDayListItem(Index: Integer): TTimeSpanDisplayItemEh;
    function GetResourceCaptionArea: TResourceVertCaptionAreaEh;

    procedure SetDataBarsArea(const Value: TDataBarsVertAreaEh);
    procedure SetDayNameArea(const Value: TDayNameVertAreaEh);
    procedure SetGridShowHours;
    procedure SetHoursColArea(const Value: THoursVertBarAreaEh);
    procedure SetMinDayColWidth(const Value: Integer);
    procedure SetResourceCaptionArea(const Value: TResourceVertCaptionAreaEh);
    procedure SetShowWorkingTimeOnly(const Value: Boolean);
    procedure SetWorkingTimeEnd(const Value: TTime);
    procedure SetWorkingTimeStart(const Value: TTime);

  protected
    FAllDayRowIndex: Integer;
    FAllDayLinesCount: Integer;
    FDayCols: Integer;
    FRowMinutesLength: Integer;
    FColDaysLength: Integer;
    FInDayList: TList;
    FAllDayList: TList;
    FInterResourceCols: Integer;
    FRolRowCount: Integer;

    function AdjustDate(const Value: TDateTime): TDateTime; override;
    function CellToDateTime(ACol, ARow: Integer): TDateTime; override;
    function CheckHitSpanItem(Button: TMouseButton; Shift: TShiftState; X, Y: Integer): TTimeSpanDisplayItemEh; override;
    function CreateDayNameArea: TDayNameAreaEh; override;
    function CreateHoursBarArea: THoursBarAreaEh; override;
    function CreateResourceCaptionArea: TResourceCaptionAreaEh; override;
    function DefaultHoursBarSize: Integer; override;
    function DrawLongDayNames: Boolean; override;
    function DrawMonthDayWithWeekDayName: Boolean; override;
    function GetCoveragePeriodType: TPlannerCoveragePeriodTypeEh; override;
    function GetGridOffsetForResource(Resource: TPlannerResourceEh): Integer; virtual;
    function GetResourceAtCell(ACol, ARow: Integer): TPlannerResourceEh; override;
    function GetResourceViewAtCell(ACol, ARow: Integer): Integer; override;
    function IsDayNameAreaNeedVisible: Boolean; override;
    function IsInterResourceCell(ACol, ARow: Integer): Boolean; virtual;
    function IsPlanItemHitAllGridArea(APlanItem: TPlannerDataItemEh): Boolean; virtual;
    function NewItemParams(var StartTime, EndTime: TDateTime; var Resource: TPlannerResourceEh): Boolean; override;
    function SelectCell(ACol, ARow: Longint): Boolean; override;
    function TimeToGridRolPos(AColStartTime, ATime: TDateTime): Integer;

    procedure InitSpanItem(ASpanItem: TTimeSpanDisplayItemEh); override;
    procedure CalcPosByPeriod(AColStartTime, AStartTime, AEndTime: TDateTime; var AStartGridPos, AStopGridPos: Integer); virtual;
    procedure GetViewPeriod(var AStartDate, AEndDate: TDateTime); override;

    procedure SetDisplayPosesSpanItems; override;
    procedure SetGroupPosesSpanItems(Resource: TPlannerResourceEh); override;

    procedure BuildGridData; override;
    procedure BuildDaysGridMode; virtual;

    procedure CalcRolRows; virtual;
    procedure StartDateChanged; override;

    procedure CheckDrawCellBorder(ACol, ARow: Integer; BorderType: TGridCellBorderTypeEh; var IsDraw: Boolean; var BorderColor: TColor; var IsExtent: Boolean); override;
//    procedure DrawCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); override;
    procedure GetCellType(ACol, ARow: Integer; var CellType: TPlannerViewCellTypeEh; var ALocalCol, ALocalRow: Integer); override;
    procedure Resize; override;
    procedure UpdateDummySpanItemSize(MousePos: TPoint); override;
    procedure CellMouseClick(const Cell: TGridCoord; Button: TMouseButton; Shift: TShiftState; const ACellRect: TRect; const GridMousePos, CellMousePos: TPoint); override;
    procedure DayNamesCellClick(const Cell: TGridCoord; Button: TMouseButton; Shift: TShiftState; const ACellRect: TRect; const GridMousePos, CellMousePos: TPoint); virtual;
    procedure PaintSpanItems; override;
    procedure DrawInDaySpanItems; virtual;
    procedure DrawAllDaySpanItems; virtual;
    procedure SortPlanItems; override;
    procedure FillSpecDaysList; virtual;
    procedure SetResOffsets; virtual;

    property ColsStartTime[ADayCol: Integer]: TDateTime read GetColsStartTime;

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    function AppendPeriod(ATime: TDateTime; Periods: Integer): TDateTime; override;
    function GetPeriodCaption: String; override;
    function NextDate: TDateTime; override;
    function PriorDate: TDateTime; override;

    property InDayList[Index: Integer]: TTimeSpanDisplayItemEh read GetInDayListItem;
    property InDayListCount: Integer read GetInDayListCount;
    property AllDayList[Index: Integer]: TTimeSpanDisplayItemEh read GetAllDayListItem;
    property AllDayListCount: Integer read GetAllDayListCount;

    property MinDayColWidth: Integer read FMinDayColWidth write SetMinDayColWidth;

    property WorkingTimeStart: TTime read FWorkingTimeStart write SetWorkingTimeStart;
    property WorkingTimeEnd: TTime read FWorkingTimeEnd write SetWorkingTimeEnd;
    property ShowWorkingTimeOnly: Boolean read FShowWorkingTimeOnly write SetShowWorkingTimeOnly default False;

  published
    property HoursColArea: THoursVertBarAreaEh read GetHoursColArea write SetHoursColArea;
    property ResourceCaptionArea: TResourceVertCaptionAreaEh read GetResourceCaptionArea write SetResourceCaptionArea;
    property DayNameArea: TDayNameVertAreaEh read GetDayNameArea write SetDayNameArea;
    property DataBarsArea: TDataBarsVertAreaEh read GetDataBarsArea write SetDataBarsArea;
    property GridLineParams;

    property OnDrawCell;
    property OnSelectionChanged;
    property OnSpanItemHintShow;
  end;

{ TPlannerDayViewEh }

  TPlannerDayViewEh = class(TPlannerWeekViewEh)
  protected
    function GetCoveragePeriodType: TPlannerCoveragePeriodTypeEh; override;
    function IsDayNameAreaNeedVisible: Boolean; override;

    procedure StartDateChanged; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

{ TPlannerMonthViewEh }

  TPlannerMonthViewEh = class(TCustomPlannerViewEh)
  private
    FSortedSpans: TList;
    FWeekColArea: TWeekBarAreaEh;

    function GetDataBarsArea: TDataBarsVertAreaEh;
    function GetDayNameArea: TDayNameVertAreaEh;
    function GetSortedSpan(Index: Integer): TTimeSpanDisplayItemEh;

    procedure SetDataBarsArea(const Value: TDataBarsVertAreaEh);
    procedure SetDayNameArea(const Value: TDayNameVertAreaEh);
    procedure SetMinDayColWidth(const Value: Integer);
    procedure SetWeekColArea(const Value: TWeekBarAreaEh);

    procedure CMFontChanged(var Message: TMessage); message CM_FONTCHANGED;

  protected
    FDefaultLineHeight: Integer;
    FShowWeekNoCaption: Boolean;
    FMinDayColWidth: Integer;
    FDataColsFor1Res: Integer;
    FDataDayNumAreaHeight: Integer;
    FMovingDaysShift: Integer;
    FWeekColIndex: Integer;

    function AdjustDate(const Value: TDateTime): TDateTime; override;
    function CalcShowKeekNoCaption(RowHeight: Integer): Boolean; virtual;
    function CellToDateTime(ACol, ARow: Integer): TDateTime; override;
    function DefaultHoursBarSize: Integer; override;
    function DrawMonthDayWithWeekDayName: Boolean; override;
    function GetCoveragePeriodType: TPlannerCoveragePeriodTypeEh; override;
    function GetResourceAtCell(ACol, ARow: Integer): TPlannerResourceEh; override;
    function GetResourceViewAtCell(ACol, ARow: Integer): Integer; override;
    function IsDayNameAreaNeedVisible: Boolean; override;
    function IsInterResourceCell(ACol, ARow: Integer): Boolean; virtual;
    function NewItemParams(var StartTime, EndTime: TDateTime; var Resource: TPlannerResourceEh): Boolean; override;
    function TimeToGridLineRolPos(ADateTime: TDateTime): Integer; virtual;
    function WeekNoColWidth: Integer; virtual;
    function CreateDayNameArea: TDayNameAreaEh; override;
    function CreateResourceCaptionArea: TResourceCaptionAreaEh; override;
    function CreateHoursBarArea: THoursBarAreaEh; override;

    procedure DrawCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); override;
    procedure DrawMonthDayNum(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure GetCellType(ACol, ARow: Integer; var CellType: TPlannerViewCellTypeEh; var ALocalCol, ALocalRow: Integer); override;

    procedure GetViewPeriod(var AStartDate, AEndDate: TDateTime); override;
    procedure SetDisplayPosesSpanItems; override;
    procedure SetDisplayPosesSpanItemsForResource(AResource: TPlannerResourceEh; Index: Integer); virtual;
    procedure SetGroupPosesSpanItems(Resource: TPlannerResourceEh); override;
    procedure SetResOffsets; virtual;
    procedure CalcPosByPeriod(AStartTime, AEndTime: TDateTime; var AStartGridPos, AStopGridPos: Integer); virtual;
    procedure InitSpanItem(ASpanItem: TTimeSpanDisplayItemEh); override;
    procedure SortPlanItems; override;
    procedure ReadPlanItem(APlanItem: TPlannerDataItemEh); override;
    procedure ReadDivByWeekPlanItem(StartDate, BoundDate: TDateTime; APlanItem: TPlannerDataItemEh);

    procedure InitSpanItemMoving(SpanItem: TTimeSpanDisplayItemEh; MousePos: TPoint); override;
    procedure UpdateDummySpanItemSize(MousePos: TPoint); override;

    procedure CellMouseClick(const Cell: TGridCoord; Button: TMouseButton; Shift: TShiftState; const ACellRect: TRect; const GridMousePos, CellMousePos: TPoint); override;
    procedure WeekNoCellClick(const Cell: TGridCoord; Button: TMouseButton; Shift: TShiftState; const ACellRect: TRect; const GridMousePos, CellMousePos: TPoint); virtual;

    procedure Resize; override;
    procedure BuildGridData; override;
    procedure BuildMonthGridMode; virtual;
    procedure CheckDrawCellBorder(ACol, ARow: Integer; BorderType: TGridCellBorderTypeEh; var IsDraw: Boolean; var BorderColor: TColor; var IsExtent: Boolean); override;

    property SortedSpan[Index: Integer]: TTimeSpanDisplayItemEh read GetSortedSpan;

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    function NextDate: TDateTime; override;
    function PriorDate: TDateTime; override;
    function AppendPeriod(ATime: TDateTime; Periods: Integer): TDateTime; override;
    function GetPeriodCaption: String; override;

    procedure DrawWeekNoCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;
    procedure DrawDataCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;
    procedure GetWeekNoCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;
    procedure GetDataCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;

    property MinDayColWidth: Integer read FMinDayColWidth write SetMinDayColWidth;

  published
    property WeekColArea: TWeekBarAreaEh read FWeekColArea write SetWeekColArea;
    property ResourceCaptionArea;
    property DayNameArea: TDayNameVertAreaEh read GetDayNameArea write SetDayNameArea;
    property DataBarsArea: TDataBarsVertAreaEh read GetDataBarsArea write SetDataBarsArea;

    property OnDrawCell;
    property OnSelectionChanged;
    property OnSpanItemHintShow;
  end;

  TTimePlanRangeKind = (rkDayByHoursEh, rkWeekByHoursEh, rkWeekByDaysEh, rkMonthByDaysEh);
  TDayslineRangeEh = (dlrWeekEh, dlrMonthEh);
  THourslineRangeEh = (hlrDayEh, hlrWeekEh);

  TAxisTimeBandOreintationEh = (atboVerticalEh, atboHorizonalEh);

{ TAxisTimeBandPlannerViewEh }

  TPlannerAxisTimelineViewEh = class(TCustomPlannerViewEh)
  private
    FRandeKind: TTimePlanRangeKind;
    FBandOreintation: TAxisTimeBandOreintationEh;
    FDatesBarArea: TDatesBarAreaEh;

    function GetSortedSpan(Index: Integer): TTimeSpanDisplayItemEh;

    procedure SetRandeKind(const Value: TTimePlanRangeKind);
    procedure SetBandOreintation(const Value: TAxisTimeBandOreintationEh);
    procedure SetDatesBarArea(const Value: TDatesBarAreaEh);

  protected
    FCellsInBand: Integer;
    FBarsInBand: Integer;
    FRolLenInSecs: Integer;
    FSortedSpans: TList;
    FTimeAxis: TGridAxisDataEh;
    FResourceAxis: TGridAxisDataEh;
    FMasterGroupLineColor: TColor;
    FMovingDaysShift: Integer;
    FMovingTimeShift: TDateTime;

    function AdjustDate(const Value: TDateTime): TDateTime; override;
    function FullRedrawOnSroll: Boolean; override;
    function GetResourceAtCell(ACol, ARow: Integer): TPlannerResourceEh; override;
    function GetGridOffsetForResource(Resource: TPlannerResourceEh): Integer; virtual;
    function DateTimeToGridRolPos(ADateTime: TDateTime): Integer; virtual;
    function IsInterResourceCell(ACol, ARow: Integer): Boolean; virtual;
    function NewItemParams(var StartTime, EndTime: TDateTime; var Resource: TPlannerResourceEh): Boolean; override;
    function GetDefaultDatesBarSize: Integer; virtual;
    function GetDefaultDatesBarVisible: Boolean; virtual;
    function IsDayNameAreaNeedVisible: Boolean; override;
    function GetDataBarsAreaDefaultBarSize: Integer; override;
    function CreateDatesBarArea: TDatesBarAreaEh; virtual;
    function GetCoveragePeriodType: TPlannerCoveragePeriodTypeEh; override;

    procedure CalcLayouts; virtual;
    procedure Resize; override;

    procedure GetViewPeriod(var AStartDate, AEndDate: TDateTime); override;
    procedure InitSpanItem(ASpanItem: TTimeSpanDisplayItemEh); override;
    procedure BuildGridData; override;
    procedure BuildDaysGridData; virtual;
    procedure BuildHoursGridData; virtual;

    procedure CalcPosByPeriod(AStartTime, AEndTime: TDateTime; var AStartGridPos, AStopGridPos: Integer); virtual;
    procedure SetGroupPosesSpanItems(Resource: TPlannerResourceEh); override;
    procedure SortPlanItems; override;
    procedure SetResOffsets; virtual;
    procedure CheckDrawCellBorder(ACol, ARow: Integer; BorderType: TGridCellBorderTypeEh; var IsDraw: Boolean; var BorderColor: TColor; var IsExtent: Boolean); override;
    procedure RangeModeChanged; override;
    procedure StartDateChanged; override;
    procedure ReadPlanItem(APlanItem: TPlannerDataItemEh); override;
    procedure PlannerStateChanged(AOldState: TPlannerStateEh); override;

    property SortedSpan[Index: Integer]: TTimeSpanDisplayItemEh read GetSortedSpan;
    property RangeKind: TTimePlanRangeKind read FRandeKind write SetRandeKind;
    property BandOreintation: TAxisTimeBandOreintationEh read FBandOreintation write SetBandOreintation;
    property DatesBarArea: TDatesBarAreaEh read FDatesBarArea write SetDatesBarArea;

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    function NextDate: TDateTime; override;
    function PriorDate: TDateTime; override;
    function AppendPeriod(ATime: TDateTime; Periods: Integer): TDateTime; override;
    function GetPeriodCaption: String; override;

    procedure CoveragePeriod(var AFromTime, AToTime: TDateTime); override;
  end;


{ TPlannerVertTimelineViewEh }

  TPlannerVertTimelineViewEh = class(TPlannerAxisTimelineViewEh)
  private
    FMinDataColWidth: Integer;

    function GetDataBarsArea: TDataBarsVertAreaEh;
    function GetDatesColArea: TDatesColAreaEh;
    function GetDayNameArea: TDayNameVertAreaEh;
    function GetHoursColArea: THoursVertBarAreaEh;
    function GetResourceCaptionArea: TResourceVertCaptionAreaEh;
    function GetSortedSpan(Index: Integer): TTimeSpanDisplayItemEh;

    procedure SetDataBarsArea(const Value: TDataBarsVertAreaEh);
    procedure SetDatesColArea(const Value: TDatesColAreaEh);
    procedure SetDayNameArea(const Value: TDayNameVertAreaEh);
    procedure SetHoursColArea(const Value: THoursVertBarAreaEh);
    procedure SetMinDataColWidth(const Value: Integer);
    procedure SetResourceCaptionArea(const Value: TResourceVertCaptionAreaEh);

  protected
    FDayGroupRows: Integer;
    FDayGroupCol: Integer;
    FDaySplitModeCol: Integer;

    function CellToDateTime(ACol, ARow: Integer): TDateTime; override;
    function IsInterResourceCell(ACol, ARow: Integer): Boolean; override;
    function GetResourceViewAtCell(ACol, ARow: Integer): Integer; override;
    function DefaultHoursBarSize: Integer; override;
    function GetDayNameAreaDefaultSize: Integer; override;
    function GetResourceCaptionAreaDefaultSize: Integer; override;
    function CreateDayNameArea: TDayNameAreaEh; override;
    function CreateResourceCaptionArea: TResourceCaptionAreaEh; override;
    function CreateHoursBarArea: THoursBarAreaEh; override;
    function CreateDatesBarArea: TDatesBarAreaEh; override;

    procedure GetViewPeriod(var AStartDate, AEndDate: TDateTime); override;
    procedure InitSpanItem(ASpanItem: TTimeSpanDisplayItemEh); override;
    procedure SetGroupPosesSpanItems(Resource: TPlannerResourceEh); override;
    procedure SetGroupPosesSpanItemsForDayStep(Resource: TPlannerResourceEh); virtual;

    procedure BuildHoursGridData; override;
    procedure CalcLayouts; override;

    procedure DrawCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); override;
    procedure GetCellType(ACol, ARow: Integer; var CellType: TPlannerViewCellTypeEh; var ALocalCol, ALocalRow: Integer); override;
    procedure DrawTimeGroupCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure DrawDaySplitModeDateCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ADataRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;

    procedure SetDisplayPosesSpanItems; override;
    procedure CalcPosByPeriod(AStartTime, AEndTime: TDateTime; var AStartGridPos, AStopGridPos: Integer); override;
    procedure SetResOffsets; override;
    procedure CheckDrawCellBorder(ACol, ARow: Integer; BorderType: TGridCellBorderTypeEh; var IsDraw: Boolean; var BorderColor: TColor; var IsExtent: Boolean); override;

    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;

    procedure InitSpanItemMoving(SpanItem: TTimeSpanDisplayItemEh; MousePos: TPoint); override;
    procedure UpdateDummySpanItemSize(MousePos: TPoint); override;

    property SortedSpan[Index: Integer]: TTimeSpanDisplayItemEh read GetSortedSpan;
    property HoursColArea: THoursVertBarAreaEh read GetHoursColArea write SetHoursColArea;

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure DrawTimeCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;
    procedure DrawDataCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;
    procedure DrawDateBar(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;
    procedure DrawDateCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;
    procedure GetDateBarDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;

    property MinDataColWidth: Integer read FMinDataColWidth write SetMinDataColWidth default -1;
    property DayNameArea: TDayNameVertAreaEh read GetDayNameArea write SetDayNameArea;
    property DataBarsArea: TDataBarsVertAreaEh read GetDataBarsArea write SetDataBarsArea;
    property DatesColArea: TDatesColAreaEh read GetDatesColArea write SetDatesColArea;
    property ResourceCaptionArea: TResourceVertCaptionAreaEh read GetResourceCaptionArea write SetResourceCaptionArea;
  end;

{ TPlannerVertDayslineViewEh }

  TPlannerVertDayslineViewEh = class(TPlannerVertTimelineViewEh)
  private
    function GetTimeRange: TDayslineRangeEh;
    procedure SetTimeRange(const Value: TDayslineRangeEh);
  protected
    function GetDefaultDatesBarSize: Integer; override;
    function GetDefaultDatesBarVisible: Boolean; override;

    procedure BuildDaysGridData; override;
    procedure DrawDaySplitModeDateCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ADataRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure GetDateCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;

  published
    property ResourceCaptionArea;
//    property DayNameArea;
    property DatesColArea;
    property DataBarsArea;
    property TimeRange: TDayslineRangeEh read GetTimeRange write SetTimeRange default dlrWeekEh;

    property OnDrawCell;
    property OnSelectionChanged;
    property OnSpanItemHintShow;
  end;

{ TPlannerVertHourslineViewEh }

  TPlannerVertHourslineViewEh = class(TPlannerVertTimelineViewEh)
  private
    function GetTimeRange: THourslineRangeEh;
    procedure SetTimeRange(const Value: THourslineRangeEh);

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

  published
    property ResourceCaptionArea;
//    property DayNameArea;
    property HoursColArea;
    property DatesColArea;
    property DataBarsArea;
    property TimeRange: THourslineRangeEh read GetTimeRange write SetTimeRange default hlrDayEh;

    property OnDrawCell;
    property OnSelectionChanged;
    property OnSpanItemHintShow;
  end;

{ TPlannerHorzTimelineViewEh }

  TPlannerHorzTimelineViewEh = class(TPlannerAxisTimelineViewEh)
  private
    FResourceColWidth: Integer;
    FShowDateRow: Boolean;
    FMinDataRowHeight: Integer;

    function GetDataBarsArea: TDataBarsHorzAreaEh;
    function GetDayNameArea: TDayNameHorzAreaEh;
    function GetHoursColArea: THoursHorzBarAreaEh;
    function GetDatesRowArea: TDatesRowAreaEh;
    function GetResourceCaptionArea: TResourceHorzCaptionAreaEh;

    procedure SetDataBarsArea(const Value: TDataBarsHorzAreaEh);
    procedure SetDatesRowArea(const Value: TDatesRowAreaEh);
    procedure SetDayNameArea(const Value: TDayNameHorzAreaEh);
    procedure SetHoursColArea(const Value: THoursHorzBarAreaEh);
    procedure SetMinDataRowHeight(const Value: Integer);
    procedure SetResourceCaptionArea(const Value: TResourceHorzCaptionAreaEh);
    procedure SetResourceColWidth(const Value: Integer);
    procedure SetShowDateRow(const Value: Boolean);

  protected
    FDayGroupCols: Integer;
    FDayGroupRow: Integer;
    FDataStartCol: Integer;
    FDaySplitModeRow: Integer;

    function CalTimeRowHeight: Integer; virtual;
    function CellToDateTime(ACol, ARow: Integer): TDateTime; override;
    function IsInterResourceCell(ACol, ARow: Integer): Boolean; override;
    function GetResourceViewAtCell(ACol, ARow: Integer): Integer; override;
    function DefaultHoursBarSize: Integer; override;
    function GetResourceCaptionAreaDefaultSize: Integer; override;
    function CreateDayNameArea: TDayNameAreaEh; override;
    function CreateDataBarsArea: TDataBarsAreaEh; override;
    function CreateResourceCaptionArea: TResourceCaptionAreaEh; override;
    function CreateHoursBarArea: THoursBarAreaEh; override;
    function CreateDatesBarArea: TDatesBarAreaEh; override;

    procedure InitSpanItem(ASpanItem: TTimeSpanDisplayItemEh); override;

//    procedure DrawCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); override;
    procedure GetCellType(ACol, ARow: Integer; var CellType: TPlannerViewCellTypeEh; var ALocalCol, ALocalRow: Integer); override;
    procedure DrawTimeGroupCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState); virtual;
    procedure DrawDaySplitModeDateCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ADataCol: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); virtual;
    procedure CheckDrawCellBorder(ACol, ARow: Integer; BorderType: TGridCellBorderTypeEh; var IsDraw: Boolean; var BorderColor: TColor; var IsExtent: Boolean); override;

    procedure CalcRectForInCellRows(SpanItem: TTimeSpanDisplayItemEh; var DrawRect: TRect); override;
    procedure CalcLayouts; override;
    procedure BuildHoursGridData; override;
    procedure SetDisplayPosesSpanItems; override;
    procedure SetResOffsets; override;

    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
    procedure InitSpanItemMoving(SpanItem: TTimeSpanDisplayItemEh; MousePos: TPoint); override;
    procedure UpdateDummySpanItemSize(MousePos: TPoint); override;

    property HoursRowArea: THoursHorzBarAreaEh read GetHoursColArea write SetHoursColArea;

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure DrawDataCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;
    procedure DrawTimeCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;
    procedure DrawDateBar(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;
    procedure DrawDateCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;
    procedure GetDateBarDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;

    property ResourceColWidth: Integer read FResourceColWidth write SetResourceColWidth;
    property ShowDateRow: Boolean read FShowDateRow write SetShowDateRow default True;

    property MinDataRowHeight: Integer read FMinDataRowHeight write SetMinDataRowHeight default -1;
    property DayNameArea: TDayNameHorzAreaEh read GetDayNameArea write SetDayNameArea ;

    property DatesRowArea: TDatesRowAreaEh read GetDatesRowArea write SetDatesRowArea;
    property DataBarsArea: TDataBarsHorzAreaEh read GetDataBarsArea write SetDataBarsArea;
    property ResourceCaptionArea: TResourceHorzCaptionAreaEh read GetResourceCaptionArea write SetResourceCaptionArea;

  end;

{ TPlannerHorzDayslineViewEh }

  TPlannerHorzDayslineViewEh = class(TPlannerHorzTimelineViewEh)
  private
    function GetTimeRange: TDayslineRangeEh;
    procedure SetTimeRange(const Value: TDayslineRangeEh);

  protected
    procedure BuildDaysGridData; override;
    procedure DrawDaySplitModeDateCell(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ADataRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure GetDateCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh); override;

  published
    property ResourceCaptionArea;
//    property DayNameArea;
    property DatesRowArea;
    property DataBarsArea;
    property TimeRange: TDayslineRangeEh read GetTimeRange write SetTimeRange default dlrWeekEh;

    property OnDrawCell;
    property OnSelectionChanged;
    property OnSpanItemHintShow;
  end;

{ TPlannerHorzHourslineViewEh }

  TPlannerHorzHourslineViewEh = class(TPlannerHorzTimelineViewEh)
  private
    function GetTimeRange: THourslineRangeEh;
    procedure SetTimeRange(const Value: THourslineRangeEh);

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

  published
    property ResourceCaptionArea;
//    property DayNameArea;
    property HoursRowArea;
    property DatesRowArea;
    property DataBarsArea;
    property TimeRange: THourslineRangeEh read GetTimeRange write SetTimeRange default hlrDayEh;

    property OnDrawCell;
    property OnSelectionChanged;
    property OnSpanItemHintShow;
  end;

{ TPlannerControlTimeSpanParamsEh }

  TPlannerControlTimeSpanParamsEh = class(TDrawElementParamsEh)
  private
    FPlanner: TPlannerControlEh;
    FDefaultColor: TColor;
    FDefaultAltColor: TColor;
    FDefaultBorderColor: TColor;

  protected
    function DefaultFont: TFont; override;
    procedure NotifyChanges; override;
    procedure ResetDefaultProps;

    function GetDefaultColor: TColor; override;
    function GetDefaultAltColor: TColor; override;
    function GetDefaultFillStyle: TPropFillStyleEh; override;
    function GetDefaultBorderColor: TColor; override;
    function GetDefaultHue: TColor; override;

  public
    constructor Create(APlanner: TPlannerControlEh);
    destructor Destroy; override;
    property Planner: TPlannerControlEh read FPlanner;

  published
    property Color;
    property Font;
    property FontStored;
    property AltColor;
    property FillStyle;
    property BorderColor;
    property Hue;
  end;

{ TEventNavBoxParamsEh }

  TEventNavBoxParamsEh = class(TDrawElementParamsEh)
  private
    FPlanner: TPlannerControlEh;
    FVisible: Boolean;
    procedure SetVisible(const Value: Boolean);

  protected
    function DefaultFont: TFont; override;
    function GetDefaultColor: TColor; override;
    function GetDefaultBorderColor: TColor; override;

    procedure NotifyChanges; override;
  public
    constructor Create(APlanner: TPlannerControlEh);
    destructor Destroy; override;
    property Planner: TPlannerControlEh read FPlanner;

  published
    property Color;
    property Font;
    property FontStored;
    property BorderColor;
    property Visible: Boolean read FVisible write SetVisible default True;
  end;

{ TPlannerToolBoxEh }

  TPlannerToolBoxEh = class(TCustomPanel)
  private
    FNextPeriodButton: TSpeedButtonEh;
    FPriorPeriodButton: TSpeedButtonEh;
    FPeriodInfo: TLabel;

  protected
    procedure PriorPeriodClick(Sender: TObject);
    procedure NextPeriodClick(Sender: TObject);
    procedure ButtonPaint(Sender: TObject);

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure UpdatePeriodInfo;

    property NextPriodButton: TSpeedButtonEh read FNextPeriodButton;
    property PriorPriodButton: TSpeedButtonEh read FPriorPeriodButton;
    property PeriodInfo: TLabel read FPeriodInfo;
  end;

{$IFDEF FPC}
{$ELSE}
{ TCustomPlannerControlPrintServiceEh }

  TCustomPlannerControlPrintServiceEh = class(TBaseGridPrintServiceEh)
  private
    FPlanner: TPlannerControlEh;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    property Planner: TPlannerControlEh read FPlanner write FPlanner;

    property ColorSchema;
    property FitToPagesTall;
    property FitToPagesWide;
    property Orientation;
    property PageFooter;
    property PageHeader;
    property PageMargins;
    property Scale;
    property ScalingMode;
    property TextAfterContent;
    property TextBeforeContent;

    property OnBeforePrint;
    property OnBeforePrintPage;
    property OnBeforePrintPageContent;
    property OnPrintDataBeforeGrid;
    property OnCalcLayoutDataBeforeGrid;

    property OnAfterPrint;
    property OnAfterPrintPage;
    property OnAfterPrintPageContent;
    property OnPrintDataAfterGrid;
    property OnCalcLayoutDataAfterGrid;

    property OnPrinterSetupDialog;
  end;
{$ENDIF}

{ TPlannerControlEh }

  TPlanItemChangeModeEh = (picmModifyEh, picmInsertEh);

  TPlannerViewOptionEh = (pvoUseGlobalWorkingTimeCalendarEh,
    pvoPlannerToolBoxEh, pvoAutoloadPlanItemsEh);
  TPlannerViewOptionsEh = set of TPlannerViewOptionEh;

  TActivePlannerViewChangedEventEh = procedure(PlannerControl: TPlannerControlEh;
    OldActivePlannerGrid: TCustomPlannerViewEh) of object;

  TShowPlanItemDialogEventEh = procedure(PlannerControl: TPlannerControlEh;
    PlannerView: TCustomPlannerViewEh; Item: TPlannerDataItemEh;
    ChangeMode: TPlanItemChangeModeEh) of object;

  TTimeItemChangingEventEh = procedure (PlannerControl: TPlannerControlEh;
    PlannerView: TCustomPlannerViewEh; Item: TPlannerDataItemEh;
    NewValuesItem: TPlannerDataItemEh; var ChangesAllowed: Boolean; var ErrorText: String) of object;

  TTimeItemChangedEventEh = procedure (PlannerControl: TPlannerControlEh;
    PlannerView: TCustomPlannerViewEh; Item: TPlannerDataItemEh;
    OldValuesItem: TPlannerDataItemEh) of object;

  TPlannerControlEh = class(TCustomPanel, ISimpleChangeNotificationEh)
  private
    FActivePlannerGrid: TCustomPlannerViewEh;
    FCurrentTime: TDateTime;
    FNotificationConsumers: TInterfaceList;
    FOptions: TPlannerViewOptionsEh;
    FPlannerGrids: TList;
    FPlannerDataSource: TPlannerDataSourceEh;
    FTopPanel: TPlannerToolBoxEh;
    FViewMode: TPlannerDateRangeModeEh;
    FTimeSpanParams: TPlannerControlTimeSpanParamsEh;
    FEventNavBoxParams: TEventNavBoxParamsEh;
    FOnDrawSpanItem: TDrawSpanItemEventEh;
    FOnActivePlannerViewChanged: TActivePlannerViewChangedEventEh;
    FOnShowPlanItemDialog: TShowPlanItemDialogEventEh;
    FOnCheckPlannerItemInteractiveChanging: TTimeItemChangingEventEh;
    FOnTimeItemInteractiveChanged: TTimeItemChangedEventEh;
    FOnSpanItemHintShow: TPlannerViewSpanItemHintShowEventEh;
    FIgnorePlannerDataSourceChanged: Boolean;
    {$IFDEF FPC}
    {$ELSE}
    FPrintService: TCustomPlannerControlPrintServiceEh;
    {$ENDIF}

    function GetActivePlannerGrid: TCustomPlannerViewEh;
    function GetActivePlannerGridIndex: Integer;
    function GetCoveragePeriodType: TPlannerCoveragePeriodTypeEh;
    function GetCurrentTime: TDateTime;
    function GetPlannerGrid(Index: Integer): TCustomPlannerViewEh;
    function GetPlannerGridCount: Integer;
    function GetStartDate: TDateTime;
    function GetPlannerDataSource: TPlannerDataSourceEh;

    procedure SetActivePlannerGrid(const Value: TCustomPlannerViewEh);
    procedure SetActivePlannerGridIndex(const Value: Integer);
    procedure SetCurrentTime(const Value: TDateTime);
    procedure SetEventNavBoxParams(const Value: TEventNavBoxParamsEh);
    procedure SetOptions(const Value: TPlannerViewOptionsEh);
    procedure SetStartDate(const Value: TDateTime);
    procedure SetPlannerDataSource(const Value: TPlannerDataSourceEh);
    procedure SetTimeSpanParams(const Value: TPlannerControlTimeSpanParamsEh);
    procedure SetViewMode(const Value: TPlannerDateRangeModeEh);
    procedure ViewModeChanged;

    procedure CMFontChanged(var Message: TMessage); message CM_FONTCHANGED;
    procedure SetOnDrawSpanItem(const Value: TDrawSpanItemEventEh);

  protected
    function GetDefaultEventNavBoxBorderColor: TColor; virtual;
    function GetDefaultEventNavBoxColor: TColor; virtual;
    function CreatePlannerItem: TPlannerDataItemEh; virtual;

    procedure ActivePlannerViewChanged(OldActivePlannerGrid: TCustomPlannerViewEh); virtual;
    procedure ChangeActivePlannerGrid(const APlannerGrid: TCustomPlannerViewEh); virtual;
    procedure CreateParams(var Params: TCreateParams); override;
    procedure CreateWnd; override;
    procedure GridCurrentTimeChanged(ANewCurrentTime: TDateTime); virtual;
    procedure LayoutChanged;
    procedure Loaded; override;
    procedure NavBoxParamsChanges; virtual;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure NotifyConsumersPlannerDataChanged; virtual;
    procedure PlannerDataSourceChange(Sender: TObject);
    procedure PlannerDataSourceChanged; virtual;
    procedure SetChildOrder(Child: TComponent; Order: Integer); override;
    procedure StartDateChanged; virtual;
    procedure GetViewPeriod(var AStartDate, AEndDate: TDateTime); virtual;
    procedure EnsureDataForPeriod(AStartDate, AEndDate: TDateTime); virtual;

  protected
    procedure ISimpleChangeNotificationEh.Change = PlannerDataSourceChange;

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    function NewItemParams(var StartTime, EndTime: TDateTime; var Resource: TPlannerResourceEh): Boolean;
    function NextDate: TDateTime;
    function PriorDate: TDateTime;
    function GetPeriodCaption: String;
    function CurWorkingTimeCalendar: TWorkingTimeCalendarEh; virtual;
//    function NewTimePlanEvent: TTimePlanItemEh;
    function CreatePlannerGrid(PlannerGridClass: TCustomPlannerGridClassEh; AOwner: TComponent): TCustomPlannerViewEh; virtual;

    procedure DefaultFillSpanItemHintShowParams(PlannerView: TCustomPlannerViewEh; CursorPos: TPoint; SpanRect: TRect; InSpanCursorPos: TPoint; SpanItem: TTimeSpanDisplayItemEh; Params: TPlannerViewSpanHintParamsEh); virtual;
    procedure ShowModifyPlanItemDialog(PlanItem: TPlannerDataItemEh); virtual;
    procedure ShowNewPlanItemDialog; virtual;
    procedure ShowDefaultPlanItemDialog(PlannerView: TCustomPlannerViewEh; Item: TPlannerDataItemEh; ChangeMode: TPlanItemChangeModeEh); virtual;

    function CheckPlannerItemInteractiveChanging(PlannerView: TCustomPlannerViewEh; Item: TPlannerDataItemEh; NewValuesItem: TPlannerDataItemEh; var ErrorText: String): Boolean; virtual;
    procedure PlannerItemInteractiveChanged(PlannerView: TCustomPlannerViewEh; Item: TPlannerDataItemEh; OldValuesItem: TPlannerDataItemEh); virtual;

    procedure CoveragePeriod(var AFromTime, AToTime: TDateTime); virtual;
//    procedure EditDialogTimePlanItem(PlanItem: TTimePlanItemEh);
    procedure RemovePlannerGrid(APlannerGrid: TCustomPlannerViewEh);

    procedure StopAutoLoad;
    procedure ResetAutoLoadProcess;

    procedure DefaultDrawSpanItem(PlannerView: TCustomPlannerViewEh; SpanItem: TTimeSpanDisplayItemEh; Rect: TRect; State: TDrawSpanItemDrawStateEh); virtual;

    procedure RegisterChanges(Value: IPlannerControlChangeReceiverEh);
    procedure UnRegisterChanges(Value: IPlannerControlChangeReceiverEh);
    procedure NextPeriod;
    procedure PriorPeriod;
    procedure GetChildren(Proc: TGetChildProc; Root: TComponent); override;
    procedure EnsureDataForViewPeriod; virtual;

    property StartDate: TDateTime read GetStartDate write SetStartDate;
    property CurrentTime: TDateTime read GetCurrentTime write SetCurrentTime;
    property ViewMode: TPlannerDateRangeModeEh read FViewMode write SetViewMode default pdrmDayEh;
    property CoveragePeriodType: TPlannerCoveragePeriodTypeEh read GetCoveragePeriodType;

    property PlannerViewCount: Integer read GetPlannerGridCount;
    property PlannerView[Index: Integer]: TCustomPlannerViewEh read GetPlannerGrid;
    property ActivePlannerGridIndex: Integer read GetActivePlannerGridIndex write SetActivePlannerGridIndex;

  published
    property ActivePlannerView: TCustomPlannerViewEh read GetActivePlannerGrid write SetActivePlannerGrid;
    property PlannerDataSource: TPlannerDataSourceEh read GetPlannerDataSource write SetPlannerDataSource;
    property Options: TPlannerViewOptionsEh read FOptions write SetOptions default [pvoUseGlobalWorkingTimeCalendarEh, pvoPlannerToolBoxEh];
    property TimeSpanParams: TPlannerControlTimeSpanParamsEh read FTimeSpanParams write SetTimeSpanParams;
    property EventNavBoxParams: TEventNavBoxParamsEh read FEventNavBoxParams write SetEventNavBoxParams;

    property Align;
    property Anchors;
    property BiDiMode;
    property BorderStyle;
    property Color;
    property Constraints;
    {$IFDEF FPC}
    {$ELSE}
    property Ctl3D;
    {$ENDIF}
    property DragCursor;
    property DragKind;
    property DragMode;
    property Enabled;
    property Font;
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
    property PrintService: TCustomPlannerControlPrintServiceEh read FPrintService;
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

    property OnDrawSpanItem: TDrawSpanItemEventEh read FOnDrawSpanItem write SetOnDrawSpanItem;
    property OnActivePlannerViewChanged: TActivePlannerViewChangedEventEh read FOnActivePlannerViewChanged write FOnActivePlannerViewChanged;
    property OnShowPlanItemDialog: TShowPlanItemDialogEventEh read FOnShowPlanItemDialog write FOnShowPlanItemDialog;
    property OnCheckPlannerItemInteractiveChanging: TTimeItemChangingEventEh read FOnCheckPlannerItemInteractiveChanging write FOnCheckPlannerItemInteractiveChanging;
    property OnPlannerItemInteractiveChanged: TTimeItemChangedEventEh read FOnTimeItemInteractiveChanged write FOnTimeItemInteractiveChanged;
    property OnSpanItemHintShow: TPlannerViewSpanItemHintShowEventEh read FOnSpanItemHintShow write FOnSpanItemHintShow;

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

function ChangeColorLuminance(AColor: TColor; ALuminance: Integer): TColor;
function ChangeColorSaturation(AColor: TColor; ASaturation: Integer): TColor;

function NormalizeDateTime(ADateTime: TDateTime): TDateTime;

procedure FillRectStyle(Canvas: TCanvas; ARect: TRect; AColor, AltColor: TColor; Style: TPropFillStyleEh);
function ChangeRelativeColorLuminance(AColor: TColor; Percent: Integer): TColor;
function RectsIntersected(const Rect1, Rect2: TRect): Boolean;

implementation

uses
{$IFDEF EH_LIB_17}
  UIConsts,
{$ENDIF}
{$IFDEF FPC}
{$ELSE}
  PrintPlannersEh,
{$ENDIF}
  EhLibConsts,
  PlannerItemDialog,
  PlannerToolCtrlsEh;

function RectsIntersected(const Rect1, Rect2: TRect): Boolean;
begin
  Result := (Rect1.Left <= Rect2.Right) and
            (Rect1.Right >= Rect2.Left) and
            (Rect1.Top <= Rect2.Bottom) and
            (Rect1.Bottom >= Rect2.Top);
end;

function SysFormatDateEh(const ASysFormat: String; ADate: TDateTime): String;
var
  SysTime: SystemTime;
  Buffer: array[Byte] of Char;
begin
  DateTimeToSystemTime(ADate, SysTime);
  if  GetDateFormat(LOCALE_USER_DEFAULT, DATE_USE_ALT_CALENDAR,
    @SysTime, PChar(ASysFormat), Buffer, SizeOf(Buffer)) <> 0
  then
    Result := Buffer
  else
    Result := '';
end;

function DateRangeToStr(AStartDate, AEndDate: TDateTime): String;
var
  y1, m1, d1: Word;
  y2, m2, d2: Word;
  s1, s2: String;
  ALastDate: TDateTime;
begin
  DecodeDate(AStartDate, y1, m1, d1);
  DecodeDate(AEndDate, y2, m2, d2);
  if DateOf(AEndDate) - DateOf(AStartDate) <= 1 then
  begin
    Result := SysFormatDateEh('d MMMM yyyy', AStartDate);
    if Result = '' then
      Result := FormatDateTime('D MMMM YYYY', AStartDate)
  end else if StartOfTheMonth(AStartDate) = StartOfTheMonth(AEndDate) then
  begin
    if DateOf(AEndDate) = AEndDate
      then ALastDate := AEndDate - 1
      else ALastDate := AEndDate;

    s1 := SysFormatDateEh('d', AStartDate) + ' - ';
    s2 := SysFormatDateEh('d MMMM yyyy', ALastDate);
    if s2 = '' then
    begin
      s1 := FormatDateTime('D', AStartDate) + ' - ';
      s2 := FormatDateTime('D MMMM YYYY', ALastDate);
    end;
    Result := s1 + s2;
  end else if (StartOfTheMonth(AStartDate) = AStartDate) and
              (IncMonth(AStartDate) = AEndDate)
  then
  begin
    Result := FormatDateTime('mmmm', AStartDate + 7) + ' ' +
              FormatDateTime('yyyy', AStartDate + 7);
  end else if StartOfTheYear(AStartDate) = StartOfTheYear(AEndDate) then
  begin
    if DateOf(AEndDate) = AEndDate
      then ALastDate := AEndDate - 1
      else ALastDate := AEndDate;
    s1 := SysFormatDateEh('d MMMM', AStartDate) + ' - ';
    s2 := SysFormatDateEh('d MMMM yyyy', ALastDate);
    if s2 = '' then
    begin
      s1 := FormatDateTime('D MMMM', AStartDate) + ' - ';
      s2 := FormatDateTime('D MMMM YYYY', ALastDate);
    end;
    Result := s1 + s2;
  end else
  begin
    if DateOf(AEndDate) = AEndDate
      then ALastDate := AEndDate - 1
      else ALastDate := AEndDate;

    s1 := SysFormatDateEh('d MMMM yyyy', AStartDate) + ' - ';
    s2 := SysFormatDateEh('d MMMM yyyy', ALastDate);
    if s2 = '' then
    begin
      s1 := FormatDateTime('D MMMM YYYY', AStartDate) + ' - ';
      s2 := FormatDateTime('D MMMM YYYY', ALastDate);
    end;
    Result := s1 + s2;
  end;
end;

procedure FillRectStyle(Canvas: TCanvas; ARect: TRect; AColor, AltColor: TColor;
  Style: TPropFillStyleEh);
begin
  if Style = fsSolidEh then
  begin
    Canvas.Brush.Color := AColor;
    Canvas.FillRect(ARect);
  end else if Style = fsVerticalGradientEh then
  begin
    FillGradientEh(Canvas, ARect, AColor, AltColor);
  end;
end;

function NormalizeDateTime(ADateTime: TDateTime): TDateTime;
var
  y, m, d, ho, mi, sec, msec: Word;
begin
  DecodeDateTime(ADateTime, y, m, d, ho, mi, sec, msec);
  Result := EncodeDateTime(y, m, d, ho, mi, sec, msec);
end;

function GetBeginningOfWeekFromOS: Integer;
var
  FDOW: array[0..1] of Char;
begin
   Result := 0;
   if GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IFIRSTDAYOFWEEK, @FDOW, SizeOf(FDOW)) > 0 then
   begin
     Result := Ord(FDOW[0]) - Ord('0');
     Assert(Result in [0..6]);
   end;
end;

function ChangeColorSaturation(AColor: TColor; ASaturation: Integer): TColor;
{$IFDEF EH_LIB_17}
var
  H, S, L: Single;
begin
  RGBtoHSL(ColorToRGB(AColor), H, S, L);
  Result := MakeColor(HSLtoRGB(H, ASaturation / 256, L), 0);
end;
{$ELSE}
var
  H, S, L: Word;
begin
  ColorRGBToHLS(ColorToRGB(AColor), H, L, S);
  Result := ColorHLSToRGB(H, L, ASaturation);
end;
{$ENDIF}

function ChangeColorLuminance(AColor: TColor; ALuminance: Integer): TColor;
{$IFDEF EH_LIB_17}
var
  H, S, L: Single;
begin
  RGBtoHSL(ColorToRGB(AColor), H, S, L);
  Result := MakeColor(HSLtoRGB(H, S, ALuminance / 256), 0);
end;
{$ELSE}
var
  H, S, L: Word;
begin
  ColorRGBToHLS(ColorToRGB(AColor), H, L, S);
  Result := ColorHLSToRGB(H, ALuminance, S);
end;
{$ENDIF}

function ChangeRelativeColorLuminance(AColor: TColor; Percent: Integer): TColor;
{$IFDEF EH_LIB_17}
var
  H, S, L, NL: Single;
begin
// The implementation in the derived class
  RGBtoHSL(ColorToRGB(AColor), H, S, L);
  if Percent > 0
    then NL := L + (1-L) / 100 * Percent
    else NL := L + L / 100 * Percent;
  Result := MakeColor(HSLtoRGB(H, S, NL), 0);
end;
{$ELSE}
var
  H, S, L, NL: Word;
begin
  ColorRGBToHLS(ColorToRGB(AColor), H, L, S);
  if Percent < 0
    then NL := L - Trunc(Percent / 100 * L)
    else NL := L + Trunc(Percent / 100 * (240 - L));
  Result := ColorHLSToRGB(H, NL, S);
end;
{$ENDIF}

{ TCustomPlannerViewEh }

constructor TCustomPlannerViewEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  inherited Options := [
    goFixedVertLineEh, goFixedHorzLineEh, goVertLineEh, goHorzLineEh,
    goDrawFocusSelectedEh];
  FRangeMode := pdrmDayEh;
  FFirstWeekDay := fwdMondayEh;
  FFirstWeekDayNum := 2;
  FBarsPerRes := 1;
  FSpanItems := TObjectList.Create(False);
  FHoursBarArea := CreateHoursBarArea;
  FResourceCaptionArea := CreateResourceCaptionArea;
  FDayNameArea :=  CreateDayNameArea;
  FDataBarsArea := CreateDataBarsArea;
  BorderStyle := bsNone;
  FDataRowsOffset := 0;
  FDataColsOffset := 0;
  FGridControls := TObjectList.Create(False);
  CreateGridControls;
  FDummyPlanItem := TPlannerDataItemEh.Create(nil);
  FDummyCheckPlanItem := TPlannerDataItemEh.Create(nil);
  HorzScrollBar.SmoothStep := True;
  ActiveMode := False;
  ShowHint := True;
  FMoveHintWindow := TMoveHintWindow.Create(Self);
  FOptions := [pgoAddEventOnDoubleClickEh];
  FSpanMoveSlidingTimer := TTimer.Create(Self);
  FSpanMoveSlidingTimer.Enabled := False;
  FSpanMoveSlidingTimer.OnTimer := SlidingTimerEvent;
  FSpanMoveSlidingTimer.Interval := 1;
  Align := alClient;
  Visible := False;
  FDrawSpanItemArgs := TDrawSpanItemArgsEh.Create;
  FDrawCellArgs := TPlannerViewCellDrawArgsEh.Create;
end;

destructor TCustomPlannerViewEh.Destroy;
begin
  Destroying;
//  PlannerDataSource := nil;
  if PlannerControl <> nil then
    PlannerControl.RemovePlannerGrid(Self);
  ClearSpanItems;
  ClearGridCells;
  FreeAndNil(FSpanItems);
  FreeAndNil(FDummyPlanItem);
  FreeAndNil(FDummyCheckPlanItem);
  DeleteGridControls;
  FreeAndNil(FGridControls);
  FreeAndNil(FMoveHintWindow);
  FreeAndNil(FSpanMoveSlidingTimer);
  FreeAndNil(FHoursBarArea);
  FreeAndNil(FResourceCaptionArea);
  FreeAndNil(FDayNameArea);
  FreeAndNil(FDataBarsArea);
  FreeAndNil(FDrawSpanItemArgs);
  FreeAndNil(FDrawCellArgs);
  FreeAndNil(FHintFont);
  inherited Destroy;
end;

procedure TCustomPlannerViewEh.CreateGridControls;
var
  GridControl: TPlannerGridControlEh;
begin
  GridControl := TPlannerGridControlEh.Create(Self);
  GridControl.FControlType := pgctNextEventEh;
  GridControl.Parent := Self;
  FGridControls.Add(GridControl);

  GridControl := TPlannerGridControlEh.Create(Self);
  GridControl.FControlType := pgctPriorEventEh;
  GridControl.Parent := Self;
  FGridControls.Add(GridControl);
end;

function TCustomPlannerViewEh.CreateGridLineColors: TGridLineColorsEh;
begin
  Result := TPlannerGridLineParamsEh.Create(Self);
end;

function TCustomPlannerViewEh.GetGridLineParams: TPlannerGridLineParamsEh;
begin
  Result := TPlannerGridLineParamsEh(GridLineColors);
end;

procedure TCustomPlannerViewEh.SetGridLineParams(
  const Value: TPlannerGridLineParamsEh);
begin
  GridLineColors := Value;
end;

function TCustomPlannerViewEh.DefaultHoursBarSize: Integer;
begin
  Result := -1;
end;

function TCustomPlannerViewEh.GetDataBarsAreaDefaultBarSize: Integer;
begin
  if HandleAllocated then
  begin
    Canvas.Font := Font;
    Result := Trunc(Canvas.TextHeight('Wg') * 1.5);
  end else
    Result := 16;
end;

function TCustomPlannerViewEh.GetDayNameAreaDefaultColor: TColor;
begin
//  Result := clMoneyGreen;
  Result := FResourceCellFillColor;
end;

function TCustomPlannerViewEh.GetDayNameAreaDefaultSize: Integer;
begin
  if HandleAllocated then
  begin
    Canvas.Font := DayNameArea.Font;
    Result := Trunc(Canvas.TextHeight('Wg') * 1.5);
  end else
    Result := 16;
end;

procedure TCustomPlannerViewEh.DeleteGridControls;
var
  i: Integer;
begin
  for i := 0 to FGridControls.Count-1 do
  begin
    FGridControls[i].Free;
    FGridControls[i] := nil;
  end;
end;

procedure TCustomPlannerViewEh.GridLayoutChanged;
begin
  if not (csLoading in ComponentState) and ActiveMode then
  begin
    HoursBarArea.RefreshDefaultFont;
    ResourceCaptionArea.RefreshDefaultFont;
    DayNameArea.RefreshDefaultFont;
    BuildGridData;
  end;
end;

procedure TCustomPlannerViewEh.BuildGridData;
//var
//  H, S, L: Single;
begin
// The implementation in the derived class
  Invalidate;
  {$IFDEF FPC}
//  NotifyControls(CM_INVALIDATE);
  {$ELSE}
  NotifyControls(CM_INVALIDATE);
  {$ENDIF}
  if ShowTopGridLine
    then FTopGridLineIndex := 0
    else FTopGridLineIndex := -1;
end;

procedure TCustomPlannerViewEh.ClearGridCells;
var
  i, j: Integer;
begin
  for i := 0 to ColCount-1 do
    for j := 0 to RowCount-1 do
      SpreadArray[i,j].Clear;
end;


procedure TCustomPlannerViewEh.SetCellCanvasParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState);
begin
  if (SelectedPlanItem <> nil) and
     (ACol >= FixedColCount-FrozenColCount) and
     (ARow >= FixedRowCount-FrozenRowCount) then
  begin
    Canvas.Font.Color := StyleServices.GetSystemColor(Font.Color);
    Canvas.Brush.Color := StyleServices.GetSystemColor(Color);
  end else
  begin
    inherited SetCellCanvasParams(ACol, ARow, ARect, State);
    if gdCurrent in State then
    begin
      Canvas.Font.Color := StyleServices.GetSystemColor(clWindowText);
      if gdFocused in State then
        Canvas.Brush.Color := FSelectedFocusedCellColor
      else
        Canvas.Brush.Color := FSelectedUnfocusedCellColor;
    end;
  end;
end;

procedure TCustomPlannerViewEh.DrawCell(ACol, ARow: Integer; ARect: TRect;
  State: TGridDrawState);
var
  CellType: TPlannerViewCellTypeEh;
  ALocalCol, ALocalRow: Integer;
  Processed: Boolean;
begin
  GetCellType(ACol, ARow, CellType, ALocalCol, ALocalRow);

//  TPlannerViewDrawCellParamsEh
  GetDrawCellParams(ACol, ARow, ARect, State, CellType, ALocalCol, ALocalRow, FDrawCellArgs);

  Processed := False;
  if Assigned(OnDrawCell) then
    OnDrawCell(Self, ACol, ARow, ARect, State, CellType, ALocalCol, ALocalRow, Processed);
  if not Processed then
    DefaultDrawPlannerViewCell(ACol, ARow, ARect, State, CellType, ALocalCol, ALocalRow, FDrawCellArgs);
end;

procedure TCustomPlannerViewEh.GetCellType(ACol, ARow: Integer;
  var CellType: TPlannerViewCellTypeEh; var ALocalCol, ALocalRow: Integer);
begin
end;

procedure TCustomPlannerViewEh.DefaultDrawPlannerViewCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; CellType: TPlannerViewCellTypeEh;
  ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh);
begin
  case CellType of
    pctTopLeftCellEh:
      DrawTopLeftCell(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctDataCellEh:
      DrawDataCell(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctAlldayDataCellEh:
      DrawAlldayDataCell(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctResourceCaptionCellEh:
      DrawResourceCaptionCell(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctInterResourceSpaceEh:
      DrawInterResourceCell(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctDayNameCellEh:
      DrawDayNamesCell(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctDateBarEh:
      DrawDateBar(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctDateCellEh:
      DrawDateCell(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctTimeCellEh:
      DrawTimeCell(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctWeekNoCell:
      DrawWeekNoCell(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
  end;
end;

procedure TCustomPlannerViewEh.DrawResourceCaptionCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
var
  s: String;
  Resource: TPlannerResourceEh;
begin
  Resource := GetResourceAtCell(ACol, ARow);
  Canvas.Font := ResourceCaptionArea.Font;
  if Resource = nil then
  begin
    Canvas.Brush.Color := StyleServices.GetSystemColor(ResourceCaptionArea.GetActualColor);
    s := SPlannerResourceUnassignedEh; //'Unassigned'
  end else
  begin
    if Resource.Color = clDefault
      then Canvas.Brush.Color := FResourceCellFillColor
      else Canvas.Brush.Color := Resource.FaceColor;
    s := Resource.Name;
  end;
  WriteTextEh(Canvas, ARect, True, 0, 0, s, taCenter, tlCenter, False, False, 0, 0, False, False);
end;

procedure TCustomPlannerViewEh.GetWeekDayNamesParams(ACol, ARow: Integer;
  ALocalCol, ALocalRow: Integer; var WeekDayNum: Integer; var WeekDayName: String);
begin
  if RangeMode = pdrmMonthEh then
  begin
    WeekDayNum := (ACol-FDataColsOffset) mod FBarsPerRes + 1;
  end else
  begin
    WeekDayNum := (ACol-1) mod FBarsPerRes + 1;
  end;

  WeekDayNum := WeekDayNum + FFirstWeekDayNum - 1;
  if WeekDayNum > 7 then WeekDayNum := WeekDayNum - 7;

  if FDayNameFormat = dnfLongFormatEh then
    WeekDayName := FormatSettings.LongDayNames[WeekDayNum]
  else if FDayNameFormat = dnfShortFormatEh then
    WeekDayName := FormatSettings.ShortDayNames[WeekDayNum]
  else
    WeekDayName := '';
end;

procedure TCustomPlannerViewEh.DrawDayNamesCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
var
  CellDate: TDateTime;
  Resource: TPlannerResourceEh;
  DName: String;
  DNum: Integer;
begin
  Resource := GetResourceAtCell(ACol, ARow);
  GetWeekDayNamesParams(ACol, ARow, ALocalCol, ALocalRow, DNum, DName);

  Canvas.Font := DayNameArea.Font;
  Canvas.Font.Color := StyleServices.GetSystemColor(DayNameArea.Font.Color);
  if (Resource <> nil) and (Resource.Color <> clDefault)
    then Canvas.Brush.Color := StyleServices.GetSystemColor(Resource.FaceColor)
    else Canvas.Brush.Color := StyleServices.GetSystemColor(DayNameArea.GetActualColor);

  WriteTextEh(Canvas, ARect, True, 0, 0, DName, taCenter, tlCenter, False, False, 0, 0, False, False);

  if DrawMonthDayWithWeekDayName then
  begin
    CellDate := CellToDateTime(ACol, ARow);
    DName := FormatDateTime('DD', CellDate);
    Canvas.Font.Style := [fsBold];
    WriteTextEh(Canvas, ARect, False, 2, 2, DName, taLeftJustify, tlTop, False, False, 0, 0, False, True);
  end;
end;

procedure TCustomPlannerViewEh.DrawInterResourceCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
begin
  Canvas.Brush.Color := StyleServices.GetSystemColor(Color);
  Canvas.FillRect(ARect);
end;

function TCustomPlannerViewEh.DrawLongDayNames: Boolean;
begin
  Result := True;
end;

function TCustomPlannerViewEh.DrawMonthDayWithWeekDayName: Boolean;
begin
  Result := True;
end;

procedure TCustomPlannerViewEh.DrawDataCell(ACol, ARow: Integer; ARect: TRect;
  State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
var
  AFillRect: TRect;
  StartDateCell: TDateTime;
  DataCell: TGridCoord;
begin
  DataCell := GridCoordToDataCoord(GridCoord(ACol, ARow));

  if (DataCell.X < 0) or (DataCell.Y < 0)  then
  begin
    AFillRect := ARect;
    Canvas.Brush.Color := StyleServices.GetSystemColor(Color);
    Canvas.FillRect(AFillRect);
  end else
  begin
    SetCellCanvasParams(ACol, ARow, ARect, State);

    if not (gdCurrent in State) then
    begin
      StartDateCell := CellToDateTime(ACol, ARow);
      if IsWorkingTime(StartDateCell) and IsWorkingDay(StartDateCell)
        then Canvas.Brush.Color := StyleServices.GetSystemColor(Color)
        else Canvas.Brush.Color := ApproximateColor(GridLineColors.GetBrightColor,
               FInternalColor, 255 div 10 * 9);
    end;

    AFillRect := ARect;
    Canvas.FillRect(AFillRect);
    Canvas.TextRect(AFillRect, AFillRect.Left+2, AFillRect.Top+2, GetCellDisplayText(ACol, ARow));

  end;
end;

procedure TCustomPlannerViewEh.DrawDateBar(ACol, ARow: Integer; ARect: TRect;
  State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
begin
end;

procedure TCustomPlannerViewEh.DrawDateCell(ACol, ARow: Integer; ARect: TRect;
  State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
begin
end;

procedure TCustomPlannerViewEh.DrawAlldayDataCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
begin
  Canvas.Brush.Color := StyleServices.GetSystemColor(Color);
  Canvas.FillRect(ARect);
end;

procedure TCustomPlannerViewEh.DrawTopLeftCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
begin
  Canvas.Brush.Color := StyleServices.GetSystemColor(clBtnFace);
  Canvas.FillRect(ARect);
end;

procedure TCustomPlannerViewEh.DrawWeekNoCell(ACol, ARow: Integer; ARect: TRect;
  State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
begin
end;

procedure TCustomPlannerViewEh.GetDrawCellParams(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; CellType: TPlannerViewCellTypeEh;
  ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh);
begin
  DrawArgs.Value := Cell[ACol, ARow].Value;
  DrawArgs.Text := '';
  DrawArgs.Alignment := taLeftJustify;
  DrawArgs.Layout := tlTop;
  DrawArgs.WordWrap := False;
  DrawArgs.Orientation := tohHorizontal;
  DrawArgs.HorzMargin := 0;
  DrawArgs.VertMargin := 0;
  DrawArgs.FontColor := Font.Color;
  DrawArgs.FontName := Font.Name;
  DrawArgs.FontSize := Font.Size;
  DrawArgs.FontStyle := Font.Style;
  DrawArgs.BackColor := Color;

  case CellType of
    pctTopLeftCellEh:
      GetTopLeftCellDrawParams(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctDataCellEh:
      GetDataCellDrawParams(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctAlldayDataCellEh:
      GetAlldayDataCellDrawParams(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctResourceCaptionCellEh:
      GetResourceCaptionCellDrawParams(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctInterResourceSpaceEh:
      GetInterResourceCellDrawParams(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctDayNameCellEh:
      GetDayNamesCellDrawParams(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctDateBarEh:
      GetDateBarDrawParams(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctDateCellEh:
      GetDateCellDrawParams(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctTimeCellEh:
      GetTimeCellDrawParams(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
    pctWeekNoCell:
      GetWeekNoCellDrawParams(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
  end;
end;

procedure TCustomPlannerViewEh.GetAlldayDataCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh);
begin
end;

procedure TCustomPlannerViewEh.GetDataCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh);
begin
end;

procedure TCustomPlannerViewEh.GetDateBarDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh);
begin
end;

procedure TCustomPlannerViewEh.GetDateCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh);
begin
end;

procedure TCustomPlannerViewEh.GetDayNamesCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh);
begin
end;

procedure TCustomPlannerViewEh.GetInterResourceCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh);
begin
end;

procedure TCustomPlannerViewEh.GetResourceCaptionCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh);
begin
end;

procedure TCustomPlannerViewEh.GetTimeCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh);
begin
end;

procedure TCustomPlannerViewEh.GetTopLeftCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh);
begin
end;

procedure TCustomPlannerViewEh.GetWeekNoCellDrawParams(ACol, ARow: Integer; ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh);
begin
end;

procedure TCustomPlannerViewEh.EnsureDataForPeriod(AStartDate,
  AEndDate: TDateTime);
begin
  if PlannerDataSource <> nil then
  begin
    FIgnorePlannerDataSourceChanged := True;
    try
      PlannerDataSource.EnsureDataForPeriod(AStartDate, AEndDate);
    finally
      FIgnorePlannerDataSourceChanged := False;
    end;
  end;
end;

function TCustomPlannerViewEh.EventNavBoxBorderColor: TColor;
begin
  if PlannerControl <> nil
    then Result := PlannerControl.EventNavBoxParams.GetActualBorderColor
    else Result := ChangeColorLuminance(FResourceCellFillColor, 150);
end;

function TCustomPlannerViewEh.EventNavBoxColor: TColor;
begin
  if PlannerControl <> nil
    then Result := PlannerControl.EventNavBoxParams.GetActualColor
    else Result := FResourceCellFillColor;
end;

function TCustomPlannerViewEh.EventNavBoxFont: TFont;
begin
  if PlannerControl <> nil
    then Result := PlannerControl.EventNavBoxParams.Font
    else Result := Font;
end;

procedure TCustomPlannerViewEh.DrawTimeCell(ACol, ARow: Integer; ARect: TRect;
  State: TGridDrawState; ALocalCol, ALocalRow: Integer; DrawArgs: TPlannerViewCellDrawArgsEh);
var
  SprCell: TSpreadGridCellEh;
  SH, SM: String;
  mSize: TSize;
  HourRect, MinRect: TRect;
begin

  SprCell := SpreadArray[ACol, ARow];

  HourRect := ARect;
  HourRect.Right := (HourRect.Right + HourRect.Left) div 2 + 5;
  MinRect := ARect;
  MinRect.Left := HourRect.Right;

  SH := Copy(VarToStr(SprCell.Value), 1, 2);
  if (SH <> '') and (SH[1] = '0') then
    SH := Copy(VarToStr(SprCell.Value), 2, 1);

  SM := ' ' + Copy(VarToStr(SprCell.Value), 1, 2);

  Canvas.Brush.Color := StyleServices.GetSystemColor(HoursBarArea.GetActualColor);
  Canvas.Font := Font;

  mSize := Canvas.TextExtent(SM);

  Canvas.Font := HoursBarArea.Font;
  Canvas.Font.Color := StyleServices.GetSystemColor(HoursBarArea.Font.Color);
  WriteText(Canvas, HourRect, -4, mSize.cy div 2, SH, taRightJustify);

  SM := '00';

  Canvas.Font.Size := Canvas.Font.Size * 2 div 3;
  WriteTextEh(Canvas, MinRect, True, 0, 2, SM,
    taLeftJustify, tlTop, False, False, 0, 0, False, True);
end;


procedure TCustomPlannerViewEh.RangeModeChanged;
begin
  GridLayoutChanged;
end;

procedure TCustomPlannerViewEh.Resize;
begin
  inherited Resize;
  if HandleAllocated and not (csLoading in ComponentState) then
  begin
    RangeModeChanged;
    RealignGridControls;
  end;
end;

function TCustomPlannerViewEh.ResourcesCount: Integer;
begin
  if PlannerDataSource <> nil
    then Result := PlannerDataSource.Resources.Count
    else Result := 0;
end;

procedure TCustomPlannerViewEh.SetFirstWeekDay(const Value: TFirstWeekDayEh);
begin
  FFirstWeekDay := Value;
  if FFirstWeekDay = fwdSundayEh
    then FFirstWeekDayNum := 1
    else FFirstWeekDayNum := 2;
end;

procedure TCustomPlannerViewEh.SetRangeMode(
  const Value: TPlannerDateRangeModeEh);
begin
  if FRangeMode <> Value then
  begin
    FRangeMode := Value;
    RangeModeChanged;
  end;
end;

procedure TCustomPlannerViewEh.SetResourceCaptionArea(
  const Value: TResourceCaptionAreaEh);
begin
  FResourceCaptionArea.Assign(Value);
end;

function TCustomPlannerViewEh.AdjustDate(const Value: TDateTime): TDateTime;
begin
  Result := DateOf(Value);
end;

procedure TCustomPlannerViewEh.SetSelectedSpanItem(
  const Value: TPlannerDataItemEh);
begin
  if FSelectedPlanItem <> Value then
  begin
    FSelectedPlanItem := Value;
    InvalidateGrid;
    SelectionChanged;
  end;
end;

procedure TCustomPlannerViewEh.SetStartDate(const Value: TDateTime);
var
  AdjustedDate: TDateTime;
begin
  AdjustedDate := AdjustDate(Value);
  if FStartDate <> AdjustedDate then
  begin
//    ReadjustDates
    FStartDate := AdjustedDate;
    ResetLoadedTimeRange;
    StartDateChanged;
  end;
end;

procedure TCustomPlannerViewEh.SetCurrentTime(const Value: TDateTime);
begin
  if FCurrentTime <> Value then
  begin
    FCurrentTime := Value;
    PlannerControl.GridCurrentTimeChanged(FCurrentTime);
    StartDate := Value;
  end;
end;

//procedure TCustomPlannerViewEh.SetPlannerDataSource(const Value: TPlannerDataSourceEh);
//begin
//  if FPlannerDataSource <> Value then
//  begin
//    if Assigned(FPlannerDataSource) then
//      FPlannerDataSource.UnRegisterChanges(Self);
//    FPlannerDataSource := Value;
//    if Assigned(FPlannerDataSource) then
//    begin
//      FPlannerDataSource.RegisterChanges(Self);
//      FPlannerDataSource.FreeNotification(Self);
//    end;
//    PlannerDataSourceChanged;
//  end;
//end;

procedure TCustomPlannerViewEh.SelectionChanged;
begin
  if Assigned(OnSelectionChanged) then
    OnSelectionChanged(Self);
end;

procedure TCustomPlannerViewEh.SetActiveMode(const Value: Boolean);
begin
  if FActiveMode <> Value then
  begin
    FActiveMode := Value;
    if Value then
      PlannerDataSourceChanged;
  end;
end;

procedure TCustomPlannerViewEh.GetViewPeriod(var AStartDate, AEndDate: TDateTime);
begin
  AStartDate := FStartDate;
  AEndDate := FStartDate + 1;
end;

procedure TCustomPlannerViewEh.GotoNextEventInTheFuture;
var
  ANextTime: Variant;
  AStartDate, AEndDate: TDateTime;
begin
  GetViewPeriod(AStartDate, AEndDate);
  ANextTime := GetNextEventAfter(AEndDate);
  if not VarIsNull(ANextTime) then
    CurrentTime := ANextTime;
end;

function TCustomPlannerViewEh.GetCoveragePeriodType: TPlannerCoveragePeriodTypeEh;
begin
  Result := pcpDayEh;
end;

function TCustomPlannerViewEh.GetNextEventAfter(ADateTime: TDateTime): Variant;
var
  i: Integer;
  PlanItem: TPlannerDataItemEh;
begin
  Result := Null;

  if Assigned(PlannerDataSource) then
  begin
//    PlannerDataSource.EnsurePeriod(AStartDate, AEndDate); // Must be in the PlannerDataSource
    for i := 0 to PlannerDataSource.Count - 1 do
    begin
      PlanItem := PlannerDataSource[i];
      if (PlanItem.StartTime > ADateTime) then // Let it be sorted
      begin
        Result := PlanItem.StartTime;
        Break;
      end;
    end;
  end;
end;

procedure TCustomPlannerViewEh.GotoPriorEventInThePast;
var
  ANextTime: Variant;
  AStartDate, AEndDate: TDateTime;
begin
  GetViewPeriod(AStartDate, AEndDate);
  ANextTime := GetNextEventBefore(AStartDate);
  if not VarIsNull(ANextTime) then
    CurrentTime := ANextTime;
end;

function TCustomPlannerViewEh.GetNextEventBefore(ADateTime: TDateTime): Variant;
var
  i: Integer;
  PlanItem: TPlannerDataItemEh;
begin
  Result := Null;

  if Assigned(PlannerDataSource) then
  begin
//    PlannerDataSource.EnsurePeriod(AStartDate, AEndDate); // Must be in the PlannerDataSource
    for i := PlannerDataSource.Count - 1 downto 0 do
    begin
      PlanItem := PlannerDataSource[i];
      if (PlanItem.StartTime < ADateTime) then // Let it be sorted
      begin
        Result := PlanItem.StartTime;
        Break;
      end;
    end;
  end;
end;

procedure TCustomPlannerViewEh.RealignGridControl(AGridControl: TPlannerGridControlEh);
var
  w1, w2, h: Integer;
  ARect: TRect;
  LeftBound: Integer;
begin
  if not HandleAllocated {or not AGridControl.HandleAllocated} then Exit;
  Canvas.Font := EventNavBoxFont;
  h := Canvas.TextHeight(SPlannerNextEventEh);
  w1 := Canvas.TextWidth(SPlannerNextEventEh);
  w2 := Canvas.TextWidth(SPlannerPriorEventEh);
  if w2 > w1 then w1 := w2;
  h := h * 2;
  w1 := Round(w1 * 1.5);
  ARect := Rect(0, 0, h, w1);
  ARect := CenteredRect(Rect(0, VertAxis.FixedBoundary, ARect.Right - ARect.Left, VertAxis.ContraStart), ARect);
  if HorzAxis.FixedBoundary > 0
    then LeftBound := HorzAxis.FixedBoundary-1
    else LeftBound := 0;
  if AGridControl.ControlType = pgctNextEventEh then
    OffsetRect(ARect, HorzAxis.ContraStart - (ARect.Right - ARect.Left), 0)
  else
    OffsetRect(ARect, LeftBound, 0);
  AGridControl.BoundsRect := ARect;
end;

procedure TCustomPlannerViewEh.RealignGridControls;
var
  i: Integer;
begin
  for i := 0 to FGridControls.Count-1 do
    TPlannerGridControlEh(FGridControls[i]).Realign;
end;

procedure TCustomPlannerViewEh.ResetGridControlsState;
var
  i: Integer;
  GridCtrl: TPlannerGridControlEh;
  AStartDate, AEndDate: TDateTime;
begin
  GetViewPeriod(AStartDate, AEndDate);
  for i := 0 to FGridControls.Count-1 do
  begin
    GridCtrl := TPlannerGridControlEh(FGridControls[i]);
    GridCtrl.Visible := (FSpanItems.Count = 0);
    if (PlannerControl <> nil) and not PlannerControl.EventNavBoxParams.Visible then
      GridCtrl.Visible := False;
    if GridCtrl.ControlType = pgctNextEventEh then
      GridCtrl.ClickEnabled := not VarIsNull(GetNextEventAfter(AEndDate))
    else
      GridCtrl.ClickEnabled := not VarIsNull(GetNextEventBefore(AStartDate));
  end;
end;

function TCustomPlannerViewEh.IsResourceCaptionNeedVisible: Boolean;
begin
  Result := (PlannerDataSource <> nil) and (PlannerDataSource.Resources.Count > 0);
end;

function TCustomPlannerViewEh.IsDayNameAreaNeedVisible: Boolean;
begin
  Result := False;
end;

function TCustomPlannerViewEh.IsWorkingDay(const Value: TDateTime): Boolean;
begin
  Result := DayOfTheWeek(Value) <= 5;
end;

function TCustomPlannerViewEh.IsWorkingTime(const Value: TDateTime): Boolean;
var
  y, m, d, ho, mi, sec, msec: Word;
  StartWorkTime, StopWorkTime: TDateTime;
begin
  DecodeDateTime(Value, y, m, d, ho, mi, sec, msec);
  StartWorkTime := EncodeDateTime(y, m, d, 8, 00, 00, 00);
  StopWorkTime := EncodeDateTime(y, m, d, 17, 00, 00, 00);
  Result := (Value >= StartWorkTime) and (Value < StopWorkTime);
end;

function TCustomPlannerViewEh.DeletePrompt: Boolean;
var
  Msg: string;
  Handle: THandle;
begin
  Handle := GetFocus;
  Msg := SPlannerDeletePlanItemEh; //'Delete Event?';
  Result := {not (dgConfirmDelete in Options) or}
    (MessageDlg(Msg, mtConfirmation, mbOKCancel, 0) <> idCancel);
  if GetFocus <> Handle then
    Windows.SetFocus(Handle);
end;

procedure TCustomPlannerViewEh.KeyDown(var Key: Word; Shift: TShiftState);
begin
  inherited KeyDown(Key, Shift);
  if (Key = VK_ESCAPE) and (FPlannerState <> psNormalEh) then
  begin
    CancelMode;
    StopPlannerState(False, -1, -1);
  end;

  if (Key = VK_DELETE) and (FSelectedPlanItem <> nil) then
    if DeletePrompt then
      PlannerControl.PlannerDataSource.DeleteItem(FSelectedPlanItem);

  if (Key = VK_F2) and (FPlannerState = psNormalEh) and (FSelectedPlanItem <> nil) then
  begin
    StartSpanMove(SpanItemByPlanItem(FSelectedPlanItem), ScreenToClient(Mouse.CursorPos));
    SetPlannerState(psSpanTestMovingEh);
  end;
end;

procedure TCustomPlannerViewEh.MouseDown(Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
const
  SideStates: array[TRectangleSideEh] of TPlannerStateEh =
    (psSpanLeftSizingEh, psSpanRightSizingEh, psSpanTopSizingEh, psSpanButtomSizingEh);
var
  HitSpanItem: TTimeSpanDisplayItemEh;
  SpanItem: TTimeSpanDisplayItemEh;
  Side: TRectangleSideEh;
begin
  FMouseDownPos := Point(X, Y);
  HitSpanItem := CheckHitSpanItem(Button, Shift, X, Y);
  if HitSpanItem <> nil then
  begin
    if not CanFocus then Exit;
    SetFocus;
    SelectedPlanItem := HitSpanItem.PlanItem;
    if (Button = mbLeft) and (ssDouble in Shift) then
      HitSpanItem.DblClick;
  end else
  begin
    SelectedPlanItem := nil;
    inherited MouseDown(Button, Shift, X, Y);
  end;
  if (Button = mbLeft) and not (ssDouble in Shift) and CheckSpanItemSizing(Point(X, Y), SpanItem, Side) then
  begin
    SetPlannerState(SideStates[Side]);
//    FDummySpanItem.Assign(SpanItem);
    FDummyPlanItem.Assign(SpanItem.PlanItem);
    FDummyPlanItemFor := SpanItem.PlanItem;
    MouseCapture := True;
  end;
end;

procedure TCustomPlannerViewEh.MouseMove(Shift: TShiftState; X, Y: Integer);
var
  ASpeed: Integer;
  ASlideDirection: TGridScrollDirection;
begin
  inherited MouseMove(Shift, X, Y);
  if (FPlannerState = psNormalEh) and CheckStartSpanMove(Point(X, Y)) then
    SetPlannerState(psSpanMovingEh);
  if FPlannerState in
      [psSpanLeftSizingEh, psSpanRightSizingEh, psSpanButtomSizingEh,
       psSpanTopSizingEh, psSpanMovingEh, psSpanTestMovingEh]
  then
    UpdateDummySpanItemSize(Point(X, Y));


  if FPlannerState in [psSpanMovingEh, psSpanTestMovingEh] then
  begin
    if CanSpanMoveSliding(Point(X, Y), ASpeed, ASlideDirection) then
      StartSpanMoveSliding(ASpeed, ASlideDirection)
    else
      StopSpanMoveSliding;
    UpdateDummySpanItemSize(Point(X, Y));
  end;
end;

procedure TCustomPlannerViewEh.MouseUp(Button: TMouseButton; Shift: TShiftState;
  X, Y: Integer);
begin
  inherited MouseUp(Button, Shift, X, Y);
  if FPlannerState in [psSpanLeftSizingEh, psSpanRightSizingEh,
      psSpanTopSizingEh, psSpanButtomSizingEh, psSpanMovingEh]
  then
    StopPlannerState(True, X, Y);
end;

procedure TCustomPlannerViewEh.CellMouseClick(const Cell: TGridCoord;
  Button: TMouseButton; Shift: TShiftState; const ACellRect: TRect;
  const GridMousePos, CellMousePos: TPoint);
begin
  inherited CellMouseClick(Cell, Button, Shift, ACellRect, GridMousePos, CellMousePos);
  if (Button = mbLeft) and
     (ssDouble in Shift) and
     (pgoAddEventOnDoubleClickEh in Options) and
     (Cell.X >= FDataColsOffset) and
     (Cell.Y >= FDataRowsOffset) and
     (SelectedPlanItem = nil)
  then
    PlannerControl.ShowNewPlanItemDialog;
end;

function TCustomPlannerViewEh.CheckStartSpanMove(MousePos: TPoint): Boolean;
var
  HitSpanItem: TTimeSpanDisplayItemEh;
begin
  Result := False;
  if MouseCapture and
   ( (Abs(FMouseDownPos.X - MousePos.X) > 5) or (Abs(FMouseDownPos.Y - MousePos.Y) > 5)) then
  begin
    HitSpanItem := CheckHitSpanItem(mbLeft, [], FMouseDownPos.X, FMouseDownPos.Y);
    if (HitSpanItem <> nil) and
       (sichSpanMovingEh in HitSpanItem.AllowedInteractiveChanges) then
    begin
      StartSpanMove(HitSpanItem, FMouseDownPos);
      Result := True;
    end;
  end;
end;

procedure TCustomPlannerViewEh.InitSpanItemMoving(SpanItem: TTimeSpanDisplayItemEh;
  MousePos: TPoint);
begin

end;

procedure TCustomPlannerViewEh.StartSpanMove(SpanItem: TTimeSpanDisplayItemEh;
  MousePos: TPoint);
begin
  if (SpanItem <> nil) and
     (sichSpanMovingEh in SpanItem.AllowedInteractiveChanges) then
  begin
    FTopLeftSpanShift.cx := -5;
    FTopLeftSpanShift.cy := -5;
    InitSpanItemMoving(SpanItem, MousePos);
    FDummyPlanItem.Assign(SpanItem.PlanItem);
    FDummyPlanItemFor := SpanItem.PlanItem;
  end;
end;

procedure TCustomPlannerViewEh.StopPlannerState(Accept: Boolean; X, Y: Integer);
var
  DataChanged: Boolean;
begin
  if FPlannerState <> psNormalEh then
  begin
    DataChanged :=
      (FDummyPlanItemFor.StartTime <> FDummyPlanItem.StartTime) or
      (FDummyPlanItemFor.EndTime <> FDummyPlanItem.EndTime) or
      (FDummyPlanItemFor.Resource <> FDummyPlanItem.Resource) or
      (FDummyPlanItemFor.AllDay <> FDummyPlanItem.AllDay);
    if Accept and DataChanged then
    begin
      FDummyCheckPlanItem.Assign(FDummyPlanItemFor);
      FDummyPlanItemFor.BeginEdit;
      FDummyPlanItemFor.StartTime := FDummyPlanItem.StartTime;
      FDummyPlanItemFor.EndTime := FDummyPlanItem.EndTime;
      FDummyPlanItemFor.Resource := FDummyPlanItem.Resource;
      FDummyPlanItemFor.AllDay := FDummyPlanItem.AllDay;
      FDummyPlanItemFor.EndEdit(True);
      NotifyPlanItemChanged(FDummyPlanItemFor, FDummyCheckPlanItem);
    end;
    SetPlannerState(psNormalEh);
    FDummyPlanItemFor := nil;
    MouseCapture := False;
    PlannerDataSourceChanged;
  end;
end;

procedure TCustomPlannerViewEh.CancelMode;
begin
  inherited CancelMode;
  if FPlannerState in [psSpanLeftSizingEh, psSpanRightSizingEh,
      psSpanTopSizingEh, psSpanButtomSizingEh, psSpanMovingEh]
  then
    StopPlannerState(False, -1, -1);
end;

procedure TCustomPlannerViewEh.StartDateChanged;
begin
  if PlannerControl <> nil then
    PlannerControl.StartDateChanged;
end;

function TCustomPlannerViewEh.CheckSpanItemSizing(MousePos: TPoint;
  out SpanItem: TTimeSpanDisplayItemEh; var Side: TRectangleSideEh): Boolean;
var
//  InRolPos: TPoint;
  Pos: TPoint;
  i: Integer;
  CurSpanItem: TTimeSpanDisplayItemEh;
  SpanDrawRect: TRect;
begin
  Result := False;
{  if (MousePos.X >= HorzAxis.FixedBoundary) and (MousePos.X < HorzAxis.ContraStart) and
     (MousePos.Y >= VertAxis.FixedBoundary) and (MousePos.Y < VertAxis.ContraStart) then
  begin
    InRolPos := ClientToGridRolPos(Point(MousePos.X, MousePos.Y));
  end else
    Exit; }

  Pos := Point(MousePos.X, MousePos.Y);

  for i := 0 to FSpanItems.Count-1 do
  begin
    CurSpanItem := SpanItems[i];
    CurSpanItem.GetInGridDrawRect(SpanDrawRect);
//    SpanDrawRect := CurSpanItem.GridBoundRect;
    if PtInRect(SpanDrawRect, Pos) then
    begin
      if (SpanDrawRect.Top + 5 >= Pos.Y) and
         (sichSpanTopSizingEh in CurSpanItem.AllowedInteractiveChanges) then
      begin
        Side := rsTopEh;
        SpanItem := CurSpanItem;
        Result := True;
      end else if (SpanDrawRect.Bottom - 5 <= Pos.Y) and
                  (sichSpanButtomSizingEh in CurSpanItem.AllowedInteractiveChanges) then
      begin
        Side := rsBottomEh;
        SpanItem := CurSpanItem;
        Result := True;
      end else if (SpanDrawRect.Left + 8 >= Pos.X) and
                  (sichSpanLeftSizingEh in CurSpanItem.AllowedInteractiveChanges) then
      begin
        Side := rsLeftEh;
        SpanItem := CurSpanItem;
        Result := True;
      end else if (SpanDrawRect.Right - 8 <= Pos.X) and
                  (sichSpanRightSizingEh in CurSpanItem.AllowedInteractiveChanges) then
      begin
        Side := rsRightEh;
        SpanItem := CurSpanItem;
        Result := True;
      end;
      Break;
    end;
  end;
end;

function TCustomPlannerViewEh.CheckHitSpanItem(Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer): TTimeSpanDisplayItemEh;
var
  Pos: TPoint;
  i: Integer;
  SpanItem: TTimeSpanDisplayItemEh;
  SpanDrawRect: TRect;
begin
  Result := nil;
  if (X >= HorzAxis.FixedBoundary) and (X < HorzAxis.ContraStart) and
     (Y >= VertAxis.FixedBoundary) and (Y < VertAxis.ContraStart) then
  begin
    Pos := Point(X, Y);
    for i := 0 to FSpanItems.Count-1 do
    begin
      SpanItem := SpanItems[i];
      SpanItem.GetInGridDrawRect(SpanDrawRect);
//      SpanDrawRect := SpanItem.GridBoundRect;
      if PtInRect(SpanDrawRect, Pos) then
      begin
        Result := SpanItem;
        Break;
      end;
    end;
  end;
end;

function TCustomPlannerViewEh.ClientToGridRolPos(Pos: TPoint): TPoint;
begin
  Result := Point(Pos.X + HorzAxis.RolStartVisPos, Pos.Y + VertAxis.RolStartVisPos);
end;

function TCustomPlannerViewEh.NewItemParams(var StartTime,
  EndTime: TDateTime; var Resource: TPlannerResourceEh): Boolean;
begin
  StartTime := 0;
  EndTime := 0;
  Resource := nil;
  Result := True;
end;

function TCustomPlannerViewEh.GridCoordToDataCoord(const AGridCoord: TGridCoord): TGridCoord;
begin
  Result.X := AGridCoord.X - FDataColsOffset;
  Result.Y := AGridCoord.Y - FDataRowsOffset;
end;

function TCustomPlannerViewEh.CellToDateTime(ACol, ARow: Integer): TDateTime;
begin
  Result := 0;
end;

procedure TCustomPlannerViewEh.DrawSpanItem(SpanItem: TTimeSpanDisplayItemEh;
  DrawRect: TRect);
var
  DrawState: TDrawSpanItemDrawStateEh;
  DrawProcessed: Boolean;
begin

  DrawState := [];
  if (SpanItem.PlanItem = SelectedPlanItem) and not (FPlannerState in [psSpanMovingEh, psSpanTestMovingEh]) then
    DrawState := DrawState + [sidsSelectedEh]
  else if (FPlannerState in [psSpanMovingEh, psSpanTestMovingEh]) and (SpanItem.PlanItem = FDummyPlanItem) then
    DrawState := DrawState + [sidsSelectedEh];

  if Focused then
    DrawState := DrawState + [sidsFocusedEh];


  DrawProcessed := False;
  FDrawSpanItemArgs.DrawState := DrawState;
  FDrawSpanItemArgs.Text := SpanItem.PlanItem.Title;
  FDrawSpanItemArgs.Alignment := SpanItem.FAlignment;

  if (PlannerControl <> nil) and Assigned(PlannerControl.OnDrawSpanItem) then
    PlannerControl.OnDrawSpanItem(PlannerControl, Self, SpanItem, DrawRect, FDrawSpanItemArgs, DrawProcessed);

  if not DrawProcessed then
    DefaultDrawSpanItem(SpanItem, DrawRect, FDrawSpanItemArgs);
end;

procedure TCustomPlannerViewEh.DefaultDrawSpanItem(
  SpanItem: TTimeSpanDisplayItemEh; const ARect: TRect;
  DrawArgs: TDrawSpanItemArgsEh);
var
  AContentRect: TRect;
begin
  DrawSpanItemBackgroud(SpanItem, ARect, DrawArgs);
  AContentRect := ARect;
  DrawSpanItemSurround(SpanItem, AContentRect, DrawArgs);
  DrawSpanItemContent(SpanItem, AContentRect, DrawArgs);
end;

procedure TCustomPlannerViewEh.DrawSpanItemBackgroud(
  SpanItem: TTimeSpanDisplayItemEh; const ARect: TRect;
  DrawArgs: TDrawSpanItemArgsEh);
begin
  Canvas.Font := GetPlannerControl.TimeSpanParams.Font;
  if SpanItem.PlanItem.FillColor <> clDefault then
  begin
    Canvas.Brush.Color := StyleServices.GetSystemColor(SpanItem.PlanItem.FillColor);
    FillGradientEh(Canvas, ARect, ChangeRelativeColorLuminance(Canvas.Brush.Color, 20), Canvas.Brush.Color);
  end else
  begin
    FillGradientEh(Canvas, ARect,
      GetPlannerControl.TimeSpanParams.GetActualColor,
      GetPlannerControl.TimeSpanParams.GetActualAltColor);
  end;

end;

procedure TCustomPlannerViewEh.DrawSpanItemSurround(
  SpanItem: TTimeSpanDisplayItemEh; var ARect: TRect;
  DrawArgs: TDrawSpanItemArgsEh);
var
  FrameRect: TRect;
//  TextRect: TRect;
  AFrameColor: TColor;
  CheckLenS: String;
  CheckLen: Integer;
  DrawFullOutInfo: Boolean;
  ImageRect: TRect;
  s: String;
begin

  if SpanItem.PlanItem.FillColor <> clDefault
    then AFrameColor := SpanItem.PlanItem.GetFrameColor
    else AFrameColor := GetPlannerControl.TimeSpanParams.GetActualBorderColor;

  if sidsSelectedEh in DrawArgs.DrawState then
  begin
    FrameRect := ARect;
    if Focused then
      Canvas.Brush.Color := ApproximateColor(
        StyleServices.GetSystemColor(clHighlight),
        StyleServices.GetSystemColor(Font.Color), 255 div 2)
    else
    begin
      if SpanItem.PlanItem.FillColor <> clDefault
        then Canvas.Brush.Color := SpanItem.PlanItem.GetFrameColor
        else Canvas.Brush.Color := AFrameColor;
    end;
    Canvas.FrameRect(FrameRect);
    InflateRect(FrameRect, -1, -1);
    Canvas.FrameRect(FrameRect);
  end else
  begin
    Canvas.Brush.Color := AFrameColor;
    Canvas.FrameRect(ARect);
  end;
  Canvas.Brush.Style := bsClear;

//  TextRect := ARect;
  CheckLenS := SPlannerPeriodFromEh+': DD MMM ' + SPlannerPeriodToEh+': DD MMM';
  CheckLen := Round(Canvas.TextWidth(CheckLenS) * 1.2);
  if (ARect.Right - ARect.Left) > CheckLen
    then DrawFullOutInfo := True
    else DrawFullOutInfo := False;

  if SpanItem.FDrawBackOutInfo then
  begin
    if SpanItem.TimeOrientation = toHorizontalEh then
    begin
      ImageRect := Rect(ARect.Left, ARect.Top, ARect.Left+16, ARect.Bottom);
      DrawClipped(PlannerDataMod.PlannerImList, nil, Canvas, ImageRect, 0, 0, taCenter, ImageRect);
      if DrawFullOutInfo
        then s := SPlannerPeriodFromEh+': ' + FormatDateTime('DD MMM', SpanItem.PlanItem.StartTime)
        else s := '';
      ARect.Left := ARect.Left + 12;
      WriteTextEh(Canvas, ARect, False, 4, 3, s,
        taLeftJustify, tlTop, True, True, 0, 0, False, True);
      ARect.Left := ARect.Left + Canvas.TextWidth(s) + 4;
      ARect.Left := ARect.Left + 12;
    end else
    begin
      ImageRect := Rect(ARect.Left, ARect.Top, ARect.Left+16, ARect.Top+16);
      DrawClipped(PlannerDataMod.PlannerImList, nil, Canvas, ImageRect, 2, 0, taCenter, ImageRect);
      ARect.Top := ARect.Top + 12;
    end;
  end;

  if SpanItem.FDrawForwardOutInfo then
  begin
    if SpanItem.TimeOrientation = toHorizontalEh then
    begin
      ImageRect := Rect(ARect.Right-16, ARect.Top, ARect.Right, ARect.Bottom);
      DrawClipped(PlannerDataMod.PlannerImList, nil, Canvas, ImageRect, 1, 0, taCenter, ImageRect);
      if DrawFullOutInfo
        then s := SPlannerPeriodToEh+': ' + FormatDateTime('DD MMM', SpanItem.PlanItem.EndTime)
        else s := '';
      ARect.Right := ARect.Right - 12;
      WriteTextEh(Canvas, ARect, False, 4, 3, s,
        taRightJustify, tlTop, True, True, 0, 0, False, True);
      ARect.Right := ARect.Right - Canvas.TextWidth(s) - 4;
      ARect.Right := ARect.Right - 12;
    end else
    begin
      ImageRect := Rect(ARect.Left, ARect.Bottom-16, ARect.Left+16, ARect.Bottom);
      DrawClipped(PlannerDataMod.PlannerImList, nil, Canvas, ImageRect, 3, 0, taCenter, ImageRect);
      ARect.Bottom := ARect.Bottom + 12;
    end;
  end;
end;

procedure TCustomPlannerViewEh.DrawSpanItemContent(
  SpanItem: TTimeSpanDisplayItemEh; const ARect: TRect;
  DrawArgs: TDrawSpanItemArgsEh);
begin
  WriteTextEh(Canvas, ARect, False, 4, 3, DrawArgs.Text,
    DrawArgs.Alignment, tlTop, True, True, 0, 0, False, True);
end;

procedure TCustomPlannerViewEh.SetPaintColors;
begin
  inherited SetPaintColors;
  FResourceCellFillColor := ApproximateColor(
    StyleServices.GetSystemColor(clBtnFace), StyleServices.GetSystemColor(clWindow), 128);
  FSelectedFocusedCellColor :=  ApproximateColor(
    StyleServices.GetSystemColor(clHighlight),
    StyleServices.GetSystemColor(clBtnFace), 200);
  FSelectedUnfocusedCellColor :=  ApproximateColor(
    StyleServices.GetSystemColor(clBtnShadow),
    StyleServices.GetSystemColor(clBtnFace), 200);
  FSpanFrameColor := ChangeColorLuminance(StyleServices.GetSystemColor(clSkyBlue), 85);
end;

procedure TCustomPlannerViewEh.Paint;
begin
  inherited Paint;
  PaintSpanItems;
end;

procedure TCustomPlannerViewEh.PaintSpanItems;
var
  SpanItem: TTimeSpanDisplayItemEh;
  i: Integer;
  SpanDrawRect: TRect;
  RestRgn: HRgn;
  RestRect: TRect;
  GridClientRect: TRect;
begin
  GridClientRect := ClientRect;
  if SpanItemsCount > 0 then
  begin
    RestRect := Rect(HorzAxis.FixedBoundary, VertAxis.FixedBoundary, HorzAxis.ContraStart, VertAxis.ContraStart);
    RestRgn := SelectClipRectangleEh(Canvas, RestRect);
    try
    for i := 0 to SpanItemsCount-1 do
    begin
      SpanItem := SpanItems[i];
      SpanItem.GetInGridDrawRect(SpanDrawRect);
      if RectsIntersected(GridClientRect, SpanDrawRect) then
        DrawSpanItem(SpanItem, SpanDrawRect);
    end;
    finally
      RestoreClipRectangleEh(Canvas, RestRgn);
    end;
  end;
(*
  if FDummySpanItemFor <> nil then
  begin
//    SpanDrawRect := FDummySpanItem.GridBoundRect;
//    OffsetRect(SpanDrawRect, -HorzAxis.RolStartVisPos, -VertAxis.RolStartVisPos);
    FDummySpanItem.GetSpanDrawRect(SpanDrawRect, Self);
    DrawSpanItem(FDummySpanItem, SpanDrawRect);
  end;
*)
end;

function TCustomPlannerViewEh.NextDate: TDateTime;
begin
  Result := 0;
end;

procedure TCustomPlannerViewEh.NextPeriod;
begin
  CurrentTime := AppendPeriod(CurrentTime, 1);
end;

procedure TCustomPlannerViewEh.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
//  if Operation = opRemove then
//  begin
//    if AComponent = PlannerDataSource then
//      PlannerDataSource := nil;
//  end;
end;

procedure TCustomPlannerViewEh.PriorPeriod;
begin
  CurrentTime := AppendPeriod(CurrentTime, -1);
end;

function TCustomPlannerViewEh.PriorDate: TDateTime;
begin
  Result := 0;
end;

function TCustomPlannerViewEh.AppendPeriod(ATime: TDateTime;
  Periods: Integer): TDateTime;
begin
  {$IFDEF FPC}
  Result := 0;
  {$ELSE}
  {$ENDIF}
  raise Exception.Create('AppendPeriod is not implemented');
end;

procedure TCustomPlannerViewEh.CalcRectForInCellCols(SpanItem: TTimeSpanDisplayItemEh;
  var DrawRect: TRect);
var
  RectWidth: Integer;
  SecWidth: Integer;
  OldDrawRect: TRect;
begin
  RectWidth := DrawRect.Right - DrawRect.Left;

  if SpanItem.InCellCols > 0 then
  begin
    SecWidth := RectWidth div SpanItem.InCellCols;
    OldDrawRect := DrawRect;
    DrawRect.Left := OldDrawRect.Left + SecWidth * SpanItem.InCellFromCol;
    DrawRect.Right := OldDrawRect.Left + SecWidth * (SpanItem.InCellToCol + 1) + 1;
  end;
end;

procedure TCustomPlannerViewEh.CalcRectForInCellRows(SpanItem: TTimeSpanDisplayItemEh;
  var DrawRect: TRect);
var
  RectHeight: Integer;
  SecHeight: Integer;
  OldDrawRect: TRect;
begin
  RectHeight := DrawRect.Bottom - DrawRect.Top;

  if SpanItem.InCellRows > 0 then
  begin
    SecHeight := RectHeight div SpanItem.InCellRows;
    OldDrawRect := DrawRect;
    DrawRect.Top := OldDrawRect.Top + SecHeight * SpanItem.InCellFromRow;
    DrawRect.Bottom := OldDrawRect.Top + SecHeight * (SpanItem.InCellToRow + 1) + 1;
  end;
end;

procedure TCustomPlannerViewEh.InitSpanItem(ASpanItem: TTimeSpanDisplayItemEh);
begin

end;

procedure TCustomPlannerViewEh.ReadPlanItem(APlanItem: TPlannerDataItemEh);
var
  SpanItem: TTimeSpanDisplayItemEh;
begin
  SpanItem := AddSpanItem(APlanItem);
//  SpanItem.FPlanItem := APlanItem;
  SpanItem.FHorzLocating := brrlGridRolAreaEh;
  SpanItem.FVertLocating := brrlGridRolAreaEh;
  SpanItem.FAllowedInteractiveChanges := [];
  InitSpanItem(SpanItem);
end;

procedure TCustomPlannerViewEh.ReadState(Reader: TReader);
begin
  inherited ReadState(Reader);
end;

procedure TCustomPlannerViewEh.ClearSpanItems;
var
  i: Integer;
begin
  for i := 0 to FSpanItems.Count-1 do
    FSpanItems[i].Free;
  FSpanItems.Clear;
end;

procedure TCustomPlannerViewEh.SetGroupPosesSpanItems(Resource: TPlannerResourceEh);
begin

end;

procedure TCustomPlannerViewEh.SetOptions(const Value: TPlannerGridOprionsEh);
begin
  if FOptions <> Value then
  begin
    FOptions := Value;
  end;
end;

function TCustomPlannerViewEh.CreateDataBarsArea: TDataBarsAreaEh;
begin
  Result := TDataBarsVertAreaEh.Create(Self);
end;

function TCustomPlannerViewEh.CreateDayNameArea: TDayNameAreaEh;
begin
  {$IFDEF FPC}
  Result := nil;
  {$ELSE}
  {$ENDIF}
  raise Exception.Create('TCustomPlannerViewEh.CreateDayNameArea must be defined');
end;

procedure TCustomPlannerViewEh.SetDayNameArea(const Value: TDayNameAreaEh);
begin
  FDayNameArea.Assign(Value);
end;

procedure TCustomPlannerViewEh.SetDisplayPosesSpanItems;
begin

end;

procedure TCustomPlannerViewEh.ResetDayNameFormat(LongDayFacor, ShortDayFacor: Double);
var
  ColWidth: Integer;
  i: Integer;
  Dname: String;
  AWidth, AMaxWidth: Integer;
begin
  if not HandleAllocated then Exit;
  ColWidth := ColWidths[FDataColsOffset];
  FDayNameFormat := dnfNonEh;
//  Canvas.Font := DayNameFont;
  Canvas.Font := Font;
  AMaxWidth := 0;
  for i := 1 to 7 do
  begin
    DName := FormatSettings.LongDayNames[i];
    AWidth := Canvas.TextWidth(DName);
    if AWidth > AMaxWidth then
      AMaxWidth := AWidth;
  end;
  if AMaxWidth * LongDayFacor < ColWidth then
    FDayNameFormat := dnfLongFormatEh;

  if FDayNameFormat <> dnfLongFormatEh then
  begin
    AMaxWidth := 0;
    for i := 1 to 7 do
    begin
      DName := FormatSettings.ShortDayNames[i];
      AWidth := Canvas.TextWidth(DName);
      if AWidth > AMaxWidth then
        AMaxWidth := AWidth;
    end;
    if AMaxWidth * ShortDayFacor < ColWidth then
      FDayNameFormat := dnfShortFormatEh;
  end;
end;

function TCustomPlannerViewEh.CheckPlanItemForRead(APlanItem: TPlannerDataItemEh): Boolean;
var
  AStartDate, AEndDate: TDateTime;
begin
  GetViewPeriod(AStartDate, AEndDate);

  if ((APlanItem.EndTime > AStartDate) and (APlanItem.StartTime < AEndDate)) or
     ((APlanItem.StartTime > AStartDate) and (APlanItem.StartTime < AEndDate)) or
     ((APlanItem.EndTime > AStartDate) and (APlanItem.EndTime < AEndDate))
  then
    Result := True
  else
    Result := False;
end;

procedure TCustomPlannerViewEh.ResetLoadedTimeRange;
begin
  if (csLoading in ComponentState) or not ActiveMode then Exit;
  MakeSpanItems;
  ProcessedSpanItems;
end;

procedure TCustomPlannerViewEh.ResetResviewArray;
var
  i: Integer;
begin
  if (PlannerDataSource <> nil) then
  begin
    if FShowUnassignedResource
      then SetLength(FResourcesView, PlannerDataSource.Resources.Count+1)
      else SetLength(FResourcesView, PlannerDataSource.Resources.Count);
  end else
  begin
    if FShowUnassignedResource
      then SetLength(FResourcesView, 1)
      else SetLength(FResourcesView, 0);
  end;

  for i := 0 to Length(FResourcesView)-1 do
  begin
    if (PlannerDataSource <> nil) and (i < PlannerDataSource.Resources.Count)
      then FResourcesView[i].Resource := PlannerDataSource.Resources[i]
      else FResourcesView[i].Resource := nil;
  end;
end;

procedure TCustomPlannerViewEh.MakeSpanItems;
var
  AStartDate, AEndDate: TDateTime;
  i: Integer;
  PlanItem: TPlannerDataItemEh;
begin
  GetViewPeriod(AStartDate, AEndDate);

//  RebuildSpanItems;
  ClearSpanItems;
  FShowUnassignedResource := False;

  if Assigned(PlannerDataSource) then
  begin
    EnsureDataForPeriod(AStartDate, AEndDate);
    for i := 0 to PlannerDataSource.Count - 1 do
    begin
      PlanItem := PlannerDataSource[i];

      if CheckPlanItemForRead(PlanItem) then
      begin
        if PlanItem <> FDummyPlanItemFor then
          ReadPlanItem(PlanItem);
      end;

      if PlanItem.Resource = nil then
        FShowUnassignedResource := True;
    end;
    if FDummyPlanItemFor <> nil then
      ReadPlanItem(FDummyPlanItem);
  end;

  if (ResourceCaptionArea.Visible = True) and (IsResourceCaptionNeedVisible = False) then
    FShowUnassignedResource := True;
end;

procedure TCustomPlannerViewEh.ProcessedSpanItems;
var
  i: Integer;
begin
  SortPlanItems;

  if (PlannerDataSource <> nil) and (PlannerDataSource.Resources.Count > 0) then
    for i := 0 to Length(FResourcesView)-1 do
      SetGroupPosesSpanItems(FResourcesView[i].Resource)
  else
    SetGroupPosesSpanItems(nil);

  SetDisplayPosesSpanItems;
  ResetGridControlsState;
  Invalidate;
end;

function CompareSpanItemFuncByPlan(Item1, Item2: Pointer): Integer;
var
  SpanItem1, SpanItem2: TTimeSpanDisplayItemEh;
begin
  SpanItem1 := TTimeSpanDisplayItemEh(Item1);
  SpanItem2 := TTimeSpanDisplayItemEh(Item2);
  if SpanItem1.PlanItem.StartTime < SpanItem2.PlanItem.StartTime then
    Result := -1
  else if SpanItem1.PlanItem.StartTime > SpanItem2.PlanItem.StartTime then
    Result := 1
  else if SpanItem1.PlanItem.EndTime < SpanItem2.PlanItem.EndTime then
    Result := -1
  else if SpanItem1.PlanItem.EndTime > SpanItem2.PlanItem.EndTime then
    Result := 1
  else
    Result := 0;
end;

{function CompareDayCellSpanItemFunc(Item1, Item2: Pointer): Integer;
var
  SpanItem1, SpanItem2: TTimeSpanDisplayItemEh;
//  DateItem1, DateItem2: TDateTime;
begin
  SpanItem1 := TTimeSpanDisplayItemEh(Item1);
  SpanItem2 := TTimeSpanDisplayItemEh(Item2);
  if DateOf(SpanItem1.PlanItem.StartTime) < DateOf(SpanItem2.PlanItem.StartTime) then
    Result := -1
  else if DateOf(SpanItem1.PlanItem.StartTime) > DateOf(SpanItem2.PlanItem.StartTime) then
    Result := 1
  else if DateOf(SpanItem1.PlanItem.EndTime) > DateOf(SpanItem2.PlanItem.EndTime) then
    Result := -1
  else if DateOf(SpanItem1.PlanItem.EndTime) < DateOf(SpanItem2.PlanItem.EndTime) then
    Result := 1
  else if SpanItem1.PlanItem.StartTime < SpanItem2.PlanItem.StartTime then
    Result := -1
  else if SpanItem1.PlanItem.StartTime > SpanItem2.PlanItem.StartTime then
    Result := 1
  else if SpanItem1.PlanItem.EndTime < SpanItem2.PlanItem.EndTime then
    Result := -1
  else if SpanItem1.PlanItem.EndTime > SpanItem2.PlanItem.EndTime then
    Result := 1
  else if SpanItem1.PlanItem = SpanItem2.PlanItem then
  begin
    if SpanItem1.StartGridRollPos < SpanItem2.StartGridRollPos then
      Result := -1
    else if SpanItem1.StartGridRollPos > SpanItem2.StartGridRollPos then
      Result := 1
    else
      Result := 0;
  end else
    Result := 0;
end;
}

function CompareSpanItemFuncBySpan(Item1, Item2: Pointer): Integer;
var
  SpanItem1, SpanItem2: TTimeSpanDisplayItemEh;
//  DateItem1, DateItem2: TDateTime;
begin
  SpanItem1 := TTimeSpanDisplayItemEh(Item1);
  SpanItem2 := TTimeSpanDisplayItemEh(Item2);
  if DateOf(SpanItem1.StartTime) < DateOf(SpanItem2.StartTime) then
    Result := -1
  else if DateOf(SpanItem1.StartTime) > DateOf(SpanItem2.StartTime) then
    Result := 1
  else if DateOf(SpanItem1.EndTime) > DateOf(SpanItem2.EndTime) then
    Result := -1
  else if DateOf(SpanItem1.EndTime) < DateOf(SpanItem2.EndTime) then
    Result := 1
  else if SpanItem1.PlanItem.StartTime < SpanItem2.PlanItem.StartTime then
    Result := -1
  else if SpanItem1.PlanItem.StartTime > SpanItem2.PlanItem.StartTime then
    Result := 1
  else if SpanItem1.PlanItem.EndTime < SpanItem2.PlanItem.EndTime then
    Result := -1
  else if SpanItem1.PlanItem.EndTime > SpanItem2.PlanItem.EndTime then
    Result := 1
  else
    Result := 0;
end;

procedure TCustomPlannerViewEh.SortPlanItems;
begin
end;

function TCustomPlannerViewEh.SpanItemByPlanItem(
  APlanItem: TPlannerDataItemEh): TTimeSpanDisplayItemEh;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to SpanItemsCount-1 do
  begin
    if SpanItems[i].PlanItem = APlanItem then
    begin
      Result := SpanItems[i];
      Exit;
    end;
  end;
end;

procedure TCustomPlannerViewEh.GroupSpanItems;
var
  i: Integer;
begin
  SortPlanItems;

  if ResourcesCount > 0 then
    for i := 0 to Length(FResourcesView)-1 do
      SetGroupPosesSpanItems(FResourcesView[i].Resource)
  else
    SetGroupPosesSpanItems(nil);
end;

procedure TCustomPlannerViewEh.ResetAllData;
begin
  MakeSpanItems;
  ResetResviewArray;
  GroupSpanItems;

  GridLayoutChanged;

  SetDisplayPosesSpanItems;
  UpdateSelectedPlanItem;
  ResetGridControlsState;
  Invalidate;
end;

procedure TCustomPlannerViewEh.PlannerDataSourceChanged;
begin
  if csDestroying in ComponentState then Exit;
  if csLoading in ComponentState then Exit;
//  PlannerControl.PlannerDataSourceChanged;
  if  FIgnorePlannerDataSourceChanged then Exit;
  ResetAllData;
end;

function TCustomPlannerViewEh.TopGridLineCount: Integer;
begin
  Result := FTopGridLineIndex+1;
end;

procedure TCustomPlannerViewEh.UpdateDefaultTimeSpanBoxHeight;
begin
  FDefaultTimeSpanBoxHeight := 0;
  if not HandleAllocated then Exit;
  Canvas.Font := Font;
  FDefaultTimeSpanBoxHeight := Canvas.TextHeight('Wg') + 3 + 3;
end;

procedure TCustomPlannerViewEh.DefaultFillSpanItemHintShowParams(CursorPos: TPoint;
  SpanRect: TRect; InSpanCursorPos: TPoint; SpanItem: TTimeSpanDisplayItemEh;
  Params: TPlannerViewSpanHintParamsEh);
begin
  Params.HintStr :=
    FormatDateTime('H:MM', SpanItem.PlanItem.StartTime) + ' - ' +
    FormatDateTime('H:MM', SpanItem.PlanItem.EndTime) + ' ' +
    SpanItem.PlanItem.Title;

  if SpanItem.PlanItem.Body <> '' then
    Params.HintStr := Params.HintStr + sLineBreak + SpanItem.PlanItem.Body;
end;


procedure TCustomPlannerViewEh.CMFontChanged(var Message: TMessage);
begin
  inherited;
  GridLayoutChanged;
end;

procedure TCustomPlannerViewEh.CMHintShow(var Message: TCMHintShow);
var
  HitSpanItem: TTimeSpanDisplayItemEh;
//  HintInfo: PHintInfo;
  ARect: TRect;
  Processed: Boolean;
  Params: TPlannerViewSpanHintParamsEh;

  procedure AssignPlannerViewHintParams(var hp:  THintInfo;
    pvhp: TPlannerViewSpanHintParamsEh);
  begin
    pvhp.HintPos := ScreenToClient(hp.HintPos);
    pvhp.HintMaxWidth := hp.HintMaxWidth;
    pvhp.HintColor := hp.HintColor;
    pvhp.ReshowTimeout := hp.ReshowTimeout;
    pvhp.HideTimeout := hp.HideTimeout;
    pvhp.HintStr := '';
    pvhp.CursorRect := ARect;
    if FHintFont = nil then
      FHintFont := TFont.Create;
    pvhp.HintFont := FHintFont;
    pvhp.HintFont.Assign(Screen.HintFont);
  end;

  procedure AssignHintParams(pvhp: TPlannerViewSpanHintParamsEh;
    var hp:  THintInfo);
  begin
    hp.CursorRect := pvhp.CursorRect;
    hp.HintPos := ClientToScreen(pvhp.HintPos);
    hp.HintMaxWidth := pvhp.HintMaxWidth;
    hp.HintColor := pvhp.HintColor;
    hp.ReshowTimeout := pvhp.ReshowTimeout;
    hp.HideTimeout := pvhp.HideTimeout;
    hp.HintStr := pvhp.HintStr;
    hp.HintWindowClass := THintWindowEh;
    hp.HintData := pvhp.HintFont;
  end;

begin
  inherited;
  ARect := Rect(0,0,0,0);
  if (FPlannerState <> psNormalEh) then Exit;

  HitSpanItem := CheckHitSpanItem(mbLeft, [], HitTest.X, HitTest.Y);
  if HitSpanItem <> nil then
  begin
    Params := TPlannerViewSpanHintParamsEh.Create;
    try
      AssignPlannerViewHintParams(PHintInfo(Message.HintInfo)^, Params);

      if (PlannerControl <> nil) and Assigned(PlannerControl.OnSpanItemHintShow) then
        PlannerControl.OnSpanItemHintShow(PlannerControl, Self, HitTest, ARect, HitTest, HitSpanItem, Params, Processed);

      if Assigned(OnSpanItemHintShow) then
        OnSpanItemHintShow(PlannerControl, Self, HitTest, ARect, HitTest, HitSpanItem, Params, Processed);

      if not Processed then
        DefaultFillSpanItemHintShowParams(HitTest, ARect, HitTest, HitSpanItem, Params);

      AssignHintParams(Params, PHintInfo(Message.HintInfo)^);
    finally
      FreeAndNil(Params);
    end;
  end;
end;

procedure TCustomPlannerViewEh.CoveragePeriod(var AFromTime,
  AToTime: TDateTime);
begin
  AFromTime := StartDate;
  case CoveragePeriodType of
    pcpDayEh: AToTime := StartDate + 1;
    pcpWeekEh: AToTime := IncWeek(StartDate);
    pcpMonthEh: AToTime := IncWeek(StartDate, 6);
  end;
end;


procedure TCustomPlannerViewEh.WMSetCursor(var Msg: TWMSetCursor);
var
  SpanItem: TTimeSpanDisplayItemEh;
  Side: TRectangleSideEh;
  Cur: HCURSOR;
begin
  Cur := 0;
  if CheckSpanItemSizing(HitTest, SpanItem, Side) then
  begin
    if Side in [rsLeftEh, rsRightEh] then
      Cur := Screen.Cursors[crSizeWE]
    else
      Cur := Screen.Cursors[crSizeNS];
  end;
  if Cur <> 0 then
  begin
     Windows.SetCursor(Cur);
{$IFDEF FPC}
     Msg.Result := 1;
{$ENDIF}
  end else
    inherited;
end;

function TCustomPlannerViewEh.GetPlannerControl: TPlannerControlEh;
begin
  Result := TPlannerControlEh(Parent);
end;

function TCustomPlannerViewEh.GetResourceAtCell(ACol,
  ARow: Integer): TPlannerResourceEh;
begin
  Result := nil;
end;

function TCustomPlannerViewEh.GetResourceCaptionAreaDefaultSize: Integer;
var
  h: Integer;
begin
  if HandleAllocated then
  begin
    Canvas.Font := ResourceCaptionArea.Font;
    h := Canvas.TextHeight('Wg') + 4;
    if h > DefaultRowHeight
      then Result := h
      else  Result := DefaultRowHeight;
  end else
    Result := DefaultRowHeight;
end;

function TCustomPlannerViewEh.GetResourceViewAtCell(ACol, ARow: Integer): Integer;
begin
  Result := -1;
end;

procedure TCustomPlannerViewEh.UpdateDummySpanItemSize(MousePos: TPoint);
begin
end;

procedure TCustomPlannerViewEh.ShowMoveHintWindow(APlanItem: TPlannerDataItemEh; MousePos: TPoint);
var
//  HintInfo: PHintInfo;
  ARect: TRect;
  APos: TPoint;
  S: String;
begin
//  HintInfo := Message.HintInfo;
  if APlanItem.InsideDayRange then
    S := FormatDateTime('H:NN', APlanItem.StartTime) + ' - ' +
         FormatDateTime('H:NN', APlanItem.EndTime)
  else
    S := FormatDateTime('DD.MMM H:NN', APlanItem.StartTime) + ' - ' +
         FormatDateTime('DD.MMM H:NN', APlanItem.EndTime);

  APos := ClientToScreen(MousePos);
  APos.X := APos.X + GetSystemMetrics(SM_CXCURSOR);
  APos.Y := APos.Y + GetSystemMetrics(SM_CYCURSOR);

  ARect := FMoveHintWindow.CalcHintRect(Screen.Width, S, nil);
  OffsetRect(ARect, APos.X, APos.Y);

//  if not FMoveHintWindow.HandleAllocated or not IsWindowVisible(FMoveHintWindow.Handle) then
    FMoveHintWindow.ActivateHint(ARect, S);
//  else
//    FMoveHintWindow.SetBounds(ARect.Left, ARect.Top, ARect.Width, ARect.Height);
end;

function TCustomPlannerViewEh.ShowTopGridLine: Boolean;
begin
  Result := True;
end;

procedure TCustomPlannerViewEh.HideMoveHintWindow;
begin
  if FMoveHintWindow.HandleAllocated and IsWindowVisible(FMoveHintWindow.Handle) then
  begin
    ShowWindow(FMoveHintWindow.Handle, SW_HIDE);
    FMoveHintWindow.Visible := False;
  end;
end;

function TCustomPlannerViewEh.GetPeriodCaption: String;
begin
  Result := '';
end;

function TCustomPlannerViewEh.InteractiveChangeAllowed: Boolean;
begin
  Result := True;
end;

procedure TCustomPlannerViewEh.SetParent(AParent: TWinControl);
begin
  if (AParent <> nil) and not (AParent is TPlannerControlEh) then
    raise Exception.Create('TCustomPlannerGridEh.SetParent: Parent must be of TPlannerControlEh type  ');

  if (Parent <> nil) and not (csDestroying in Parent.ComponentState) then
// (TPlannerControlEh(Parent).FPlannerGrids <> nil) then
  begin
    TPlannerControlEh(Parent).FPlannerGrids.Remove(Self);
//    ShowMessage('TPlannerControlEh(Parent).FPlannerGrids.Remove(Self);');
  end;
  inherited SetParent(AParent);
  if AParent <> nil then
    TPlannerControlEh(AParent).FPlannerGrids.Add(Self);
end;

procedure TCustomPlannerViewEh.SetPlannerState(ANewState: TPlannerStateEh);
var
  OldPlannerState: TPlannerStateEh;
begin
  if FPlannerState <> ANewState then
  begin
    OldPlannerState := FPlannerState;
    FPlannerState := ANewState;
    PlannerStateChanged(OldPlannerState);
  end;
end;

procedure TCustomPlannerViewEh.PlannerStateChanged(AOldState: TPlannerStateEh);
begin
  HideMoveHintWindow;
end;

function TCustomPlannerViewEh.GetSpanItems(
  Index: Integer): TTimeSpanDisplayItemEh;
begin
  Result := TTimeSpanDisplayItemEh(FSpanItems[Index]);
end;

function TCustomPlannerViewEh.GetSpanItemsCount: Integer;
begin
  Result := FSpanItems.Count;
end;

function TCustomPlannerViewEh.AddSpanItem(APlanItem: TPlannerDataItemEh): TTimeSpanDisplayItemEh;
begin
  Result := TTimeSpanDisplayItemEh.Create(Self, APlanItem);
  FSpanItems.Add(Result);
end;

procedure TCustomPlannerViewEh.StartSpanMoveSliding(ASpeed: Integer;
  ASlideDirection: TGridScrollDirection);
begin
  FSpanMoveSlidingSpeed := ASpeed;
  FSlideDirection := ASlideDirection;
  if not FSpanMoveSlidingTimer.Enabled then
  begin
    FSpanMoveSlidingTimer.Enabled := True;
    SlidingTimerEvent(nil);
  end;
end;

procedure TCustomPlannerViewEh.SlidingTimerEvent(Sender: TObject);
begin
  if FSlideDirection = sdUp then
    SafeScrollData(0, -FSpanMoveSlidingSpeed)
  else if FSlideDirection = sdDown then
    SafeScrollData(0, FSpanMoveSlidingSpeed)
  else if FSlideDirection = sdLeft then
    SafeScrollData(-FSpanMoveSlidingSpeed, 0)
  else if FSlideDirection = sdRight then
    SafeScrollData(FSpanMoveSlidingSpeed, 0);
  UpdateDummySpanItemSize(ScreenToClient(Mouse.CursorPos));
end;

procedure TCustomPlannerViewEh.StopSpanMoveSliding;
begin
  FSpanMoveSlidingTimer.Enabled := False;
end;

function TCustomPlannerViewEh.CanSpanMoveSliding(MousePos: TPoint;
  var ASpeed: Integer; var ASlideDirection: TGridScrollDirection): Boolean;
var
  SlidingExtent: Integer;
  VertRolExtend, HorzRolExtend: Integer;
begin
  Result := False;
  if FPlannerState in [psSpanMovingEh, psSpanTestMovingEh] then
  begin
    SlidingExtent := DefaultRowHeight * 2;
    VertRolExtend := VertAxis.FixedBoundary + VertAxis.RolLen - VertAxis.RolStartVisPos;
    HorzRolExtend := HorzAxis.FixedBoundary + HorzAxis.RolLen - HorzAxis.RolStartVisPos;
    if (MousePos.Y > VertAxis.FixedBoundary) and
       (MousePos.Y < VertAxis.FixedBoundary + SlidingExtent) and
       (VertAxis.RolStartVisPos > 0) then
    begin
      ASpeed := ((VertAxis.FixedBoundary + SlidingExtent) - MousePos.Y) div 2;
      ASlideDirection := sdUp;
      Result := True;
    end
    else if (MousePos.Y > VertAxis.ContraStart - SlidingExtent) and
            (MousePos.Y < VertAxis.ContraStart) and
            (VertRolExtend > VertAxis.ContraStart) then
    begin
      ASpeed := (MousePos.Y - (VertAxis.ContraStart - SlidingExtent)) div 2;
      ASlideDirection := sdDown;
      Result := True;
    end else if (MousePos.X > HorzAxis.FixedBoundary) and
                (MousePos.X < HorzAxis.FixedBoundary + SlidingExtent) and
                (HorzAxis.RolStartVisPos > 0) then
    begin
      ASpeed := ((HorzAxis.FixedBoundary + SlidingExtent) - MousePos.X) div 2;
      ASlideDirection := sdLeft;
      Result := True;
    end
    else if (MousePos.X > HorzAxis.ContraStart - SlidingExtent) and
            (MousePos.X < HorzAxis.ContraStart) and
            (HorzRolExtend > HorzAxis.ContraStart) then
    begin
      ASpeed := (MousePos.X - (HorzAxis.ContraStart - SlidingExtent)) div 2;
      ASlideDirection := sdRight;
      Result := True;
    end;
  end;
end;

function TCustomPlannerViewEh.GetGridIndex: Integer;
begin
  if PlannerControl <> nil
    then Result := PlannerControl.FPlannerGrids.IndexOf(Self)
    else Result := -1;
end;

procedure TCustomPlannerViewEh.SetGridIndex(const Value: Integer);
var
  I, MaxGridIndex: Integer;
begin
  if PlannerControl <> nil then
  begin
    MaxGridIndex := PlannerControl.PlannerViewCount - 1;
    if Value > MaxGridIndex then
      raise EListError.Create('SGridIndexError');
    I := PlannerControl.FPlannerGrids.IndexOf(Self);
    PlannerControl.FPlannerGrids.Move(I, Value);
  end;
end;

function TCustomPlannerViewEh.CreateHoursBarArea: THoursBarAreaEh;
begin
  {$IFDEF FPC}
  Result := nil;
  {$ELSE}
  {$ENDIF}
  raise Exception.Create('TCustomPlannerViewEh.CreateHoursBarArea must be defined');
end;

function TCustomPlannerViewEh.CreateResourceCaptionArea: TResourceCaptionAreaEh;
begin
  Result := nil;
end;

procedure TCustomPlannerViewEh.CreateWnd;
begin
  inherited CreateWnd;
  ResetLoadedTimeRange;
  GridLayoutChanged;
end;

procedure TCustomPlannerViewEh.CurrentCellMoved(OldCurrent: TGridCoord);
begin
  inherited CurrentCellMoved(OldCurrent);
  SelectionChanged;
end;

procedure TCustomPlannerViewEh.SetHoursBarArea(const Value: THoursBarAreaEh);
begin
  FHoursBarArea.Assign(Value);
end;

procedure TCustomPlannerViewEh.SetDataBarsArea(const Value: TDataBarsAreaEh);
begin
  FDataBarsArea.Assign(Value);
end;

procedure TCustomPlannerViewEh.CheckSetDummyPlanItem(Item,
  NewItem: TPlannerDataItemEh);
var
  ErrorText: String;
begin
  ErrorText := '';
  if PlannerControl.CheckPlannerItemInteractiveChanging(Self, Item, NewItem, ErrorText) then
    FDummyPlanItem.Assign(NewItem);
end;

procedure TCustomPlannerViewEh.NotifyPlanItemChanged(Item,
  OldItem: TPlannerDataItemEh);
begin
  PlannerControl.PlannerItemInteractiveChanged(Self, Item, OldItem);
end;

procedure TCustomPlannerViewEh.UpdateSelectedPlanItem;
var
  i: Integer;
  SelFound: Boolean;
begin
  SelFound := False;
  for i := 0 to SpanItemsCount-1 do
  begin
    if SpanItems[i].PlanItem = SelectedPlanItem then
    begin
      SelFound := True;
      Break;
    end;
  end;
  if SelFound = False then
    SelectedPlanItem := nil;
end;

function TCustomPlannerViewEh.GetPlannerDataSource: TPlannerDataSourceEh;
begin
  if PlannerControl <> nil
    then Result := PlannerControl.PlannerDataSource
    else Result := nil;
end;

{ TTimeSpanDisplayItemEh }

procedure TTimeSpanDisplayItemEh.Assign(Source: TPersistent);
begin
  inherited Assign(Source);
  if Source is TTimeSpanDisplayItemEh then
  begin
    FPlanItem := TTimeSpanDisplayItemEh(Source).PlanItem;
    FStartGridRollPos := TTimeSpanDisplayItemEh(Source).StartGridRollPos;
    FStopGridRolPos := TTimeSpanDisplayItemEh(Source).StopGridRolPosl;

    FInCellCols := TTimeSpanDisplayItemEh(Source).InCellCols;
    FInCellFromCol := TTimeSpanDisplayItemEh(Source).InCellFromCol;
    FInCellToCol := TTimeSpanDisplayItemEh(Source).InCellToCol;
    FGridColNum := TTimeSpanDisplayItemEh(Source).GridColNum;

  end;
end;

constructor TTimeSpanDisplayItemEh.Create(AGrid: TCustomPlannerViewEh;
  APlanItem: TPlannerDataItemEh);
begin
  inherited Create(AGrid);
  FPlanItem := APlanItem;
  FAlignment := taLeftJustify;
end;

procedure TTimeSpanDisplayItemEh.DblClick;
begin
  PlannerView.PlannerControl.ShowModifyPlanItemDialog(PlanItem);
end;

destructor TTimeSpanDisplayItemEh.Destroy;
begin
  inherited Destroy;
end;

{ TDummyTimeSpanDisplayItemEh }

procedure TDummyTimeSpanDisplayItemEh.Assign(Source: TPersistent);
begin
  if Source is TTimeSpanDisplayItemEh then
  begin
    FPlanItem := TTimeSpanDisplayItemEh(Source).PlanItem;
    FStartTime := TTimeSpanDisplayItemEh(Source).PlanItem.StartTime;
    FEndTime := TTimeSpanDisplayItemEh(Source).PlanItem.EndTime;

    FStartGridRollPos := TTimeSpanDisplayItemEh(Source).StartGridRollPos;
    FStopGridRolPos := TTimeSpanDisplayItemEh(Source).StopGridRolPosl;

    FInCellCols := TTimeSpanDisplayItemEh(Source).InCellCols;
    FInCellFromCol := TTimeSpanDisplayItemEh(Source).InCellFromCol;
    FInCellToCol := TTimeSpanDisplayItemEh(Source).InCellToCol;
    FGridColNum := TTimeSpanDisplayItemEh(Source).GridColNum;

    FBoundRect := TTimeSpanDisplayItemEh(Source).BoundRect;
    FHorzLocating := TTimeSpanDisplayItemEh(Source).HorzLocating;
    FVertLocating := TTimeSpanDisplayItemEh(Source).VertLocating;
  end;
end;

procedure TDummyTimeSpanDisplayItemEh.SetEndTime(const Value: TDateTime);
begin
  FEndTime := Value;
end;

procedure TDummyTimeSpanDisplayItemEh.SetStartTime(const Value: TDateTime);
begin
  FStartTime := Value;
end;

{ TPlannerControlEh }

constructor TPlannerControlEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  ControlStyle := [csCaptureMouse, csClickEvents, csSetCaption, csOpaque,
    csDoubleClicks, csReplicatable {$IFDEF EH_LIB_17}, csPannable, csGestures {$ENDIF}];

  {$IFDEF FPC}
  {$ELSE}
  BevelKind := bkFlat;
  BevelOuter := bvNone;
  {$ENDIF}
  FOptions := [pvoUseGlobalWorkingTimeCalendarEh, pvoPlannerToolBoxEh];

  FNotificationConsumers := TInterfaceList.Create;

  FTopPanel := TPlannerToolBoxEh.Create(Self);
  FTopPanel.Parent := Self;
  FTopPanel.Align := alTop;
  FTopPanel.BevelOuter := bvNone;
  {$IFDEF FPC}
  {$ELSE}
  FTopPanel.BevelKind := bkFlat;
  FTopPanel.BevelEdges := [];//[beBottom];
  {$ENDIF}

  FPlannerGrids := TList.Create;
  FTimeSpanParams := TPlannerControlTimeSpanParamsEh.Create(Self);
  FEventNavBoxParams := TEventNavBoxParamsEh.Create(Self);

  CurrentTime := Today;
  ViewModeChanged;

  {$IFDEF FPC}
  {$ELSE}
  FPrintService := TPlannerControlPrintServiceEh.Create(Self);
  FPrintService.Planner := Self;
  FPrintService.Name := 'PrintService';
  FPrintService.SetSubComponent(True);
  {$ENDIF}
end;

destructor TPlannerControlEh.Destroy;
begin
  Destroying;

  while FPlannerGrids.Count > 0 do
    TCustomPlannerViewEh(FPlannerGrids[FPlannerGrids.Count-1]).Free;
  FreeAndNil(FPlannerGrids);

  FreeAndNil(FNotificationConsumers);
  FreeAndNil(FTopPanel);
  FreeAndNil(FTimeSpanParams);
  FreeAndNil(FEventNavBoxParams);
  {$IFDEF FPC}
  {$ELSE}
  FreeAndNil(FPrintService);
  {$ENDIF}

  inherited Destroy;
end;

procedure TPlannerControlEh.DefaultDrawSpanItem(
  PlannerView: TCustomPlannerViewEh; SpanItem: TTimeSpanDisplayItemEh;
  Rect: TRect; State: TDrawSpanItemDrawStateEh);
begin

end;

function TPlannerControlEh.CurWorkingTimeCalendar: TWorkingTimeCalendarEh;
begin
  if pvoUseGlobalWorkingTimeCalendarEh in Options
    then Result := GlobalWorkingTimeCalendar
    else Result := nil;
end;

function TPlannerControlEh.GetStartDate: TDateTime;
begin
  Result := ActivePlannerView.StartDate;
end;

procedure TPlannerControlEh.SetStartDate(const Value: TDateTime);
begin
  if ActivePlannerView <> nil then
    ActivePlannerView.StartDate := Value;
  FTopPanel.UpdatePeriodInfo;
end;

function TPlannerControlEh.GetCoveragePeriodType: TPlannerCoveragePeriodTypeEh;
begin
  if ActivePlannerView <> nil
    then Result := ActivePlannerView.CoveragePeriodType
    else Result := pcpDayEh;
end;

function TPlannerControlEh.GetCurrentTime: TDateTime;
begin
  Result := FCurrentTime;
end;

function TPlannerControlEh.GetDefaultEventNavBoxBorderColor: TColor;
begin
  Result := ChangeColorLuminance(ActivePlannerView.FResourceCellFillColor, 150);
end;

function TPlannerControlEh.GetDefaultEventNavBoxColor: TColor;
begin
  Result := ActivePlannerView.FResourceCellFillColor;
end;

procedure TPlannerControlEh.SetCurrentTime(const Value: TDateTime);
begin
  if FCurrentTime <> Value then
  begin
    if ActivePlannerView <> nil
      then ActivePlannerView.CurrentTime := Value
      else GridCurrentTimeChanged(Value);
    EnsureDataForViewPeriod;
  end;
end;

procedure TPlannerControlEh.SetEventNavBoxParams(
  const Value: TEventNavBoxParamsEh);
begin
  FEventNavBoxParams.Assign(Value);
end;

procedure TPlannerControlEh.PlannerDataSourceChange(Sender: TObject);
begin
  PlannerDataSourceChanged;
end;

procedure TPlannerControlEh.PlannerDataSourceChanged;
begin
  if FIgnorePlannerDataSourceChanged then Exit;
  FIgnorePlannerDataSourceChanged := True;
  try
    EnsureDataForViewPeriod;
  finally
    FIgnorePlannerDataSourceChanged := False;
  end;
  if ActivePlannerView <> nil then
    ActivePlannerView.PlannerDataSourceChanged;
  NotifyConsumersPlannerDataChanged;
end;

procedure TPlannerControlEh.EnsureDataForViewPeriod;
var
  AStartDate, AEndDate: TDateTime;
begin
  GetViewPeriod(AStartDate, AEndDate);
  EnsureDataForPeriod(AStartDate, AEndDate);
end;

procedure TPlannerControlEh.GetViewPeriod(var AStartDate, AEndDate: TDateTime);
var
  NewStartDate, NewEndDate: TDateTime;
  i: Integer;
begin
  if ActivePlannerView <> nil then
    ActivePlannerView.GetViewPeriod(AStartDate, AEndDate);
  if FNotificationConsumers <> nil then
  begin
    for i := 0 to FNotificationConsumers.Count - 1 do
    begin
      (FNotificationConsumers[I] as IPlannerControlChangeReceiverEh).GetViewPeriod(NewStartDate, NewEndDate);
      if (NewStartDate <> 0) or (NewEndDate <> 0) then
      begin
        if NewStartDate < AStartDate then
          AStartDate := NewStartDate;
        if NewEndDate > AEndDate then
          AEndDate := NewEndDate;
      end;
    end;
  end;
end;

procedure TPlannerControlEh.EnsureDataForPeriod(AStartDate, AEndDate: TDateTime);
begin
  if PlannerDataSource <> nil then
    PlannerDataSource.EnsureDataForPeriod(AStartDate, AEndDate);
end;

procedure TPlannerControlEh.GridCurrentTimeChanged(ANewCurrentTime: TDateTime);
begin
  FCurrentTime := ANewCurrentTime;
  FTopPanel.UpdatePeriodInfo;
end;

procedure TPlannerControlEh.SetOnDrawSpanItem(const Value: TDrawSpanItemEventEh);
begin
  FOnDrawSpanItem := Value;
  if ActivePlannerView <> nil then
    ActivePlannerView.Invalidate;
end;

procedure TPlannerControlEh.SetOptions(const Value: TPlannerViewOptionsEh);
begin
  if FOptions <> Value then
  begin
    FOptions := Value;
    if pvoPlannerToolBoxEh in FOptions
      then FTopPanel.Visible := True
      else FTopPanel.Visible := False;
  end;
end;

function TPlannerControlEh.GetPlannerDataSource: TPlannerDataSourceEh;
begin
  Result := FPlannerDataSource;
end;

procedure TPlannerControlEh.SetPlannerDataSource(const Value: TPlannerDataSourceEh);
begin
  if FPlannerDataSource <> Value then
  begin
    if Assigned(FPlannerDataSource) then
      FPlannerDataSource.UnRegisterChanges(Self);
    FPlannerDataSource := Value;
//    if ActivePlannerView <> nil then
//      ActivePlannerView.PlannerDataSource := Value;
    if Assigned(FPlannerDataSource) then
    begin
      FPlannerDataSource.RegisterChanges(Self);
      FPlannerDataSource.FreeNotification(Self);
    end;
    if FPlannerDataSource <> nil then
      FPlannerDataSource.FreeNotification(Self);
    PlannerDataSourceChanged;
  end;
end;

procedure TPlannerControlEh.SetTimeSpanParams(
  const Value: TPlannerControlTimeSpanParamsEh);
begin
  FTimeSpanParams.Assign(Value);
end;

procedure TPlannerControlEh.NavBoxParamsChanges;
begin
  if ActivePlannerView <> nil then
  begin
    ActivePlannerView.GridLayoutChanged;
    ActivePlannerView.ResetGridControlsState;
  end;
end;

function TPlannerControlEh.NewItemParams(var StartTime,
  EndTime: TDateTime; var Resource: TPlannerResourceEh): Boolean;
begin
  Result := ActivePlannerView.NewItemParams(StartTime, EndTime, Resource);
end;

function TPlannerControlEh.NextDate: TDateTime;
begin
  Result := ActivePlannerView.NextDate;
end;

function TPlannerControlEh.PriorDate: TDateTime;
begin
  Result := ActivePlannerView.PriorDate;
end;

procedure TPlannerControlEh.NextPeriod;
begin
  if ActivePlannerView <> nil then
    ActivePlannerView.NextPeriod;
end;

procedure TPlannerControlEh.PriorPeriod;
begin
  if ActivePlannerView <> nil then
    ActivePlannerView.PriorPeriod;
end;

procedure TPlannerControlEh.SetViewMode(const Value: TPlannerDateRangeModeEh);
begin
  if FViewMode <> Value then
  begin
    FViewMode := Value;
    ViewModeChanged;
  end;
end;

procedure TPlannerControlEh.ShowDefaultPlanItemDialog(
  PlannerView: TCustomPlannerViewEh; Item: TPlannerDataItemEh;
  ChangeMode: TPlanItemChangeModeEh);
begin
  if ChangeMode = picmModifyEh then
    EditPlanItem(Self, Item)
  else
    EditNewItem(Self);
end;

procedure TPlannerControlEh.ShowModifyPlanItemDialog(PlanItem: TPlannerDataItemEh);
begin
  if Assigned(OnShowPlanItemDialog) then
    OnShowPlanItemDialog(Self, ActivePlannerView, PlanItem, picmModifyEh)
  else
    ShowDefaultPlanItemDialog(ActivePlannerView, PlanItem, picmModifyEh);
end;

procedure TPlannerControlEh.ShowNewPlanItemDialog;
begin
  if Assigned(OnShowPlanItemDialog) then
    OnShowPlanItemDialog(Self, ActivePlannerView, nil, picmInsertEh)
  else
    ShowDefaultPlanItemDialog(ActivePlannerView, nil, picmInsertEh);
end;

procedure TPlannerControlEh.ResetAutoLoadProcess;
begin

end;

procedure TPlannerControlEh.StopAutoLoad;
begin

end;

procedure TPlannerControlEh.StartDateChanged;
begin
  FTopPanel.UpdatePeriodInfo;
  NotifyConsumersPlannerDataChanged;
end;

procedure TPlannerControlEh.RegisterChanges(Value: IPlannerControlChangeReceiverEh);
begin
  if FNotificationConsumers.IndexOf(Value) < 0 then
    FNotificationConsumers.Add(Value);
end;

procedure TPlannerControlEh.UnRegisterChanges(Value: IPlannerControlChangeReceiverEh);
begin
  if FNotificationConsumers <> nil then
    FNotificationConsumers.Remove(Value);
end;

procedure TPlannerControlEh.Notification(AComponent: TComponent;
  Operation: TOperation);
//var
//  i: Integer;
begin
  inherited Notification(AComponent, Operation);
  if Operation = opRemove then
  begin
    if AComponent = PlannerDataSource then
    begin
      PlannerDataSource := nil;
//      for i := 0 to PlannerViewCount-1 do
//        PlannerView[i].PlannerDataSource := nil;
    end;
  end;
end;

procedure TPlannerControlEh.NotifyConsumersPlannerDataChanged;
var
  i: Integer;
begin
  if FNotificationConsumers <> nil then
    for I := 0 to FNotificationConsumers.Count - 1 do
      (FNotificationConsumers[I] as IPlannerControlChangeReceiverEh).Change(Self);
end;

procedure TPlannerControlEh.ViewModeChanged;
begin
end;

function TPlannerControlEh.GetPeriodCaption: String;
begin
  if FActivePlannerGrid <> nil
    then Result := FActivePlannerGrid.GetPeriodCaption
    else Result := '';
end;

function TPlannerControlEh.CreatePlannerGrid(
  PlannerGridClass: TCustomPlannerGridClassEh;
  AOwner: TComponent): TCustomPlannerViewEh;
begin
  Result := PlannerGridClass.Create(AOwner);
  Result.Parent := Self;
  Result.Align := alClient;
  Result.Visible := False;
  if ActivePlannerView = nil then
    ActivePlannerView := Result;
end;

function TPlannerControlEh.CreatePlannerItem: TPlannerDataItemEh;
begin
  Result := PlannerDataSource.CreatePlannerItem;
end;

procedure TPlannerControlEh.CreateWnd;
begin
  inherited CreateWnd;
  TimeSpanParams.ResetDefaultProps;
end;

procedure TPlannerControlEh.RemovePlannerGrid(APlannerGrid: TCustomPlannerViewEh);
var
  RemoveIndex: Integer;
begin
  RemoveIndex := FPlannerGrids.Remove(APlannerGrid);
  if APlannerGrid = ActivePlannerView then
  begin
    if (PlannerViewCount > 0) and not (csDestroying in ComponentState) then
      ActivePlannerView := PlannerView[0]
    else
      ActivePlannerView := nil;
  end;
  if RemoveIndex >= 0 then
  begin
//    APlannerGrid.PlannerDataSource := nil;
    APlannerGrid.Parent := nil;
  end;
end;

function TPlannerControlEh.GetPlannerGrid(Index: Integer): TCustomPlannerViewEh;
begin
  Result := TCustomPlannerViewEh(FPlannerGrids[Index]);
end;

function TPlannerControlEh.GetPlannerGridCount: Integer;
begin
  Result := FPlannerGrids.Count;
end;

function TPlannerControlEh.GetActivePlannerGrid: TCustomPlannerViewEh;
begin
  Result := FActivePlannerGrid;
end;

procedure TPlannerControlEh.SetActivePlannerGrid(const Value: TCustomPlannerViewEh);
begin
  if (Value <> nil) and (Value.PlannerControl <> Self) then
    raise Exception.Create('PlannerGrid "' + Value.Name + '" is not belong to PlannerView');
//  FInSetActivePage := True;
  try
    ChangeActivePlannerGrid(Value);
  finally
//    FInSetActivePage := False;
  end;
end;

procedure TPlannerControlEh.ActivePlannerViewChanged(
  OldActivePlannerGrid: TCustomPlannerViewEh);
begin
  if not (csDestroying in ComponentState) then
    if Assigned(OnActivePlannerViewChanged) then
      OnActivePlannerViewChanged(Self, OldActivePlannerGrid);
end;

procedure TPlannerControlEh.ChangeActivePlannerGrid(
  const APlannerGrid: TCustomPlannerViewEh);
var
  OldActivePlannerGrid: TCustomPlannerViewEh;
begin
  if FActivePlannerGrid <> APlannerGrid then
  begin
    OldActivePlannerGrid := FActivePlannerGrid;
    if FActivePlannerGrid <> nil then
    begin
      FActivePlannerGrid.Visible := False;
      FActivePlannerGrid.ActiveMode := False;
//      FActivePlannerGrid.PlannerDataSource := nil;
    end;

    FActivePlannerGrid := APlannerGrid;

    if FActivePlannerGrid <> nil then
    begin
      FActivePlannerGrid.CurrentTime := CurrentTime;
//      FActivePlannerGrid.PlannerDataSource := PlannerDataSource;
      FActivePlannerGrid.ActiveMode := True;
      FActivePlannerGrid.Visible := True;
      FActivePlannerGrid.BringToFront;
    end;

    StartDateChanged;
    ActivePlannerViewChanged(OldActivePlannerGrid);
  end;
end;

function TPlannerControlEh.GetActivePlannerGridIndex: Integer;
begin
  if ActivePlannerView <> nil
    then Result := ActivePlannerView.ViewIndex
    else Result := -1;
end;

procedure TPlannerControlEh.SetActivePlannerGridIndex(const Value: Integer);
begin
  ActivePlannerView := PlannerView[Value];
end;

procedure TPlannerControlEh.GetChildren(Proc: TGetChildProc; Root: TComponent);
var
  I: Integer;
begin
  for I := 0 to PlannerViewCount - 1 do
    Proc(PlannerView[i]);
end;

procedure TPlannerControlEh.SetChildOrder(Child: TComponent; Order: Integer);
begin
  TCustomPlannerViewEh(Child).ViewIndex := Order;
end;

procedure TPlannerControlEh.LayoutChanged;
begin
  if ActivePlannerView <> nil then
    ActivePlannerView.GridLayoutChanged;
end;

procedure TPlannerControlEh.Loaded;
begin
  inherited Loaded;
  if (ActivePlannerView = nil) and (PlannerViewCount > 0) then
    ActivePlannerView := PlannerView[0];
//  if (pvoAutoloadPlanItemsEh in Options) and (PlannerDataSource <> nil) then
end;

procedure TPlannerControlEh.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  Params.Style := Params.Style or WS_CLIPCHILDREN;
  Params.ExStyle := Params.ExStyle or WS_EX_CONTROLPARENT;
end;

procedure TPlannerControlEh.CMFontChanged(var Message: TMessage);
begin
  inherited;
  TimeSpanParams.RefreshDefaultFont;
end;

procedure TPlannerControlEh.CoveragePeriod(var AFromTime, AToTime: TDateTime);
begin
  if ActivePlannerView <> nil then
    ActivePlannerView.CoveragePeriod(AFromTime, AToTime)
  else
  begin
    AFromTime := 0;
    AToTime := 0;
  end;
end;

function TPlannerControlEh.CheckPlannerItemInteractiveChanging(
  PlannerView: TCustomPlannerViewEh; Item, NewValuesItem: TPlannerDataItemEh;
  var ErrorText: String): Boolean;
begin
  Result := True;
  if Assigned(OnCheckPlannerItemInteractiveChanging) then
    OnCheckPlannerItemInteractiveChanging(Self, PlannerView, Item, NewValuesItem,
      Result, ErrorText);
end;

procedure TPlannerControlEh.PlannerItemInteractiveChanged(
  PlannerView: TCustomPlannerViewEh; Item, OldValuesItem: TPlannerDataItemEh);
begin
  if Assigned(OnPlannerItemInteractiveChanged) then
    OnPlannerItemInteractiveChanged(Self, PlannerView, Item, OldValuesItem);
end;

procedure TPlannerControlEh.DefaultFillSpanItemHintShowParams(
  PlannerView: TCustomPlannerViewEh; CursorPos: TPoint;
  SpanRect: TRect; InSpanCursorPos: TPoint; SpanItem: TTimeSpanDisplayItemEh;
  Params: TPlannerViewSpanHintParamsEh);
begin
  PlannerView.DefaultFillSpanItemHintShowParams(CursorPos, SpanRect,
    InSpanCursorPos, SpanItem, Params);
end;

{ TPlannerToolBoxEh }

constructor TPlannerToolBoxEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  FPriorPeriodButton := TSpeedButtonEh.Create(Self);
  FPriorPeriodButton.Parent := Self;
  FPriorPeriodButton.Top := 10;
  FPriorPeriodButton.Left := 10;
//  FPriorPeriodButton.Caption := '<';
  FPriorPeriodButton.Flat := True;
  FPriorPeriodButton.Style := ebsGlyphEh;
//  FPriorPeriodButton.Enabled := False;
  FPriorPeriodButton.OnClick := PriorPeriodClick;
  FPriorPeriodButton.OnPaint := ButtonPaint;

  FNextPeriodButton := TSpeedButtonEh.Create(Self);
  FNextPeriodButton.Parent := Self;
  FNextPeriodButton.Top := 10;
  FNextPeriodButton.Left := 35;
//  FNextPeriodButton.Caption := '>';
  FNextPeriodButton.Flat := True;
  FNextPeriodButton.Style := ebsGlyphEh;
//  FNextPeriodButton.Enabled := False;
  FNextPeriodButton.OnClick := NextPeriodClick;
  FNextPeriodButton.OnPaint := ButtonPaint;

  FPeriodInfo := TLabel.Create(Self);
  FPeriodInfo.Parent := Self;
  FPeriodInfo.Top := 10;
  FPeriodInfo.Left := 70;
  FPeriodInfo.Font.Size := 12;
  FPeriodInfo.Caption := FormatDateTime(FormatSettings.LongDateFormat, Date);
//  FPeriodInfo.Caption := '<';
  DoubleBuffered := True;
end;

destructor TPlannerToolBoxEh.Destroy;
begin
  inherited Destroy;
end;

procedure TPlannerToolBoxEh.ButtonPaint(Sender: TObject);
var
  AButton: TSpeedButtonEh;
  ADrawImages: TCustomImageList;
  AImageIndex: Integer;
  ADrawRect: TRect;
begin
  AButton := TSpeedButtonEh(Sender);
//  AButton.DefaultPaint;

  ADrawImages := PlannerDataMod.PlannerImList;
  if AButton = FPriorPeriodButton
    then AImageIndex := 5
    else AImageIndex := 6;
  if (ADrawImages <> nil) and (AImageIndex >= 0) then
  begin
    ADrawRect := Rect(0,0,ADrawImages.Width,ADrawImages.Height);
    ADrawRect := CenteredRect(Rect(0,0,AButton.Width,AButton.Height), ADrawRect);
//    ADrawRect.Bottom := ADrawRect.Bottom;
    if AButton.GetState = ebstPressedEh then
      OffsetRect(ADrawRect,1,1);
    if AButton.GetState in [ebstPressedEh, ebstHotEh] then
    begin
      AButton.Canvas.Brush.Color := ApproximateColor(
        StyleServices.GetSystemColor(clHighlight),
        StyleServices.GetSystemColor(clBtnFace), 200);
      AButton.Canvas.FillRect(Rect(0,0,Width,Height));
    end;

    ADrawImages.Draw(AButton.Canvas, ADrawRect.Left, ADrawRect.Top, AImageIndex, Enabled);
  end;
end;

procedure TPlannerToolBoxEh.NextPeriodClick(Sender: TObject);
begin
//  TPlannerControlEh(Owner).StartDate := TPlannerControlEh(Owner).NextDate;
  TPlannerControlEh(Owner).NextPeriod;
end;

procedure TPlannerToolBoxEh.PriorPeriodClick(Sender: TObject);
begin
//  TPlannerControlEh(Owner).StartDate := TPlannerControlEh(Owner).PriorDate;
  TPlannerControlEh(Owner).PriorPeriod;
end;

procedure TPlannerToolBoxEh.UpdatePeriodInfo;
begin
  FPeriodInfo.Caption := TPlannerControlEh(Owner).GetPeriodCaption;
end;

{ TPlannerWeekViewEh }

constructor TPlannerWeekViewEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FRowMinutesLength := 30;
  FColDaysLength := 1;
  FDayCols := 7;
  HorzScrollBar.VisibleMode := sbNeverShowEh;
  FDataRowsOffset := 1;
  FDataColsOffset := 1;
  FAllDayRowIndex := -1;
  FInDayList :=  TList.Create;
  FAllDayList := TList.Create;
  FMinDayColWidth := 50;
  FWorkingTimeStart := TTime(EncodeTime(8, 0, 0, 0));
  FWorkingTimeEnd := TTime(EncodeTime(17, 0, 0, 0));
  SetGridShowHours;
end;

destructor TPlannerWeekViewEh.Destroy;
begin
  Destroying;
  FreeAndNil(FInDayList);
  FreeAndNil(FAllDayList);
  inherited Destroy;
end;

procedure TPlannerWeekViewEh.BuildGridData;
begin
  inherited BuildGridData;
  SetGridShowHours;
  if Length(FResourcesView) > 1
    then FInterResourceCols := 1
    else FInterResourceCols := 0;
  UpdateDefaultTimeSpanBoxHeight;
  BuildDaysGridMode;
  if HandleAllocated then
    ResetDayNameFormat(2, 3);
  RealignGridControls;
end;

procedure TPlannerWeekViewEh.BuildDaysGridMode;
var
  i: Integer;
  ColGroups: Integer;
  ARowCount: Integer;
  ir: Integer;
  AnAllDayLinesCount: Integer;
  AStartHour: Integer;
  AFixedRowCount: Integer;
  AHoursBarWidth: Integer;
//  h: Integer;
  ColWidth, FitGap: Integer;
  DataColsWidth: Integer;
  AHoursColCount: Integer;
//  TextHeight, TextExtLead: Integer;
begin

  if FAllDayList = nil then Exit;

  ClearGridCells;
  FAllDayLinesCount := FAllDayList.Count;

  CalcRolRows;

  AFixedRowCount := TopGridLineCount;

  if ResourceCaptionArea.Visible then
  begin
    ColGroups := ResourcesCount;
    if FShowUnassignedResource then
      Inc(ColGroups);

    FResourceAxisPos := AFixedRowCount;
    AFixedRowCount := AFixedRowCount + 2;
    ARowCount := FRolRowCount + AFixedRowCount;
    FAllDayRowIndex := AFixedRowCount-1;
  end else
  begin
    ColGroups := ResourcesCount;
    if FShowUnassignedResource then
      Inc(ColGroups);
    if ColGroups = 0 then
      ColGroups := 1;
    AFixedRowCount := AFixedRowCount + 1;
    ARowCount := FRolRowCount + AFixedRowCount;
    FResourceAxisPos := -1;
    FAllDayRowIndex := AFixedRowCount-1;
  end;


  if DayNameArea.Visible then
  begin
    FDayNameBarPos := AFixedRowCount-1;
    Inc(ARowCount);
    Inc(AFixedRowCount);
    Inc(FAllDayRowIndex);
  end else
    FDayNameBarPos := -1;


  if HoursColArea.Visible then
  begin
    FixedColCount := 1;
    FDataColsOffset := 1;
    FHoursBarIndex := 0;
    AHoursColCount := 1;
  end else
  begin
    FixedColCount := 0;
    FDataColsOffset := 0;
    FHoursBarIndex := -1;
    AHoursColCount := 0;
  end;
  FixedRowCount := AFixedRowCount;
  FDataRowsOffset := AFixedRowCount;

  ColCount := FDayCols * ColGroups + AHoursColCount + (ColGroups-1);
  RowCount := ARowCount;
  SetGridSize(FullColCount, FullRowCount);
  if ColGroups = 1
    then FBarsPerRes := FDayCols
    else FBarsPerRes := FDayCols + 1;

  if FHoursBarIndex >= 0 then
  begin
    ColWidths[FHoursBarIndex] := HoursColArea.GetActualSize;
    AHoursBarWidth := HoursColArea.GetActualSize;
  end else
    AHoursBarWidth := 0;


  DataColsWidth := (GridClientWidth - AHoursBarWidth - (ColGroups-1)*3);
  ColWidth := DataColsWidth div (ColGroups * FDayCols);
  if ColWidth < MinDayColWidth then
  begin
    ColWidth := MinDayColWidth;
    HorzScrollBar.VisibleMode := sbAutoShowEh;
    FitGap := 0;
  end else
  begin
    HorzScrollBar.VisibleMode := sbNeverShowEh;
    FitGap := DataColsWidth mod (ColGroups * FDayCols);
  end;

  for i := FixedColCount to ColCount-1 do
  begin
    if IsInterResourceCell(i, 0) then
    begin
      ColWidths[i] := 3;
    end else
    begin
      if FitGap > 0
        then ColWidths[i] := ColWidth + 1
        else ColWidths[i] := ColWidth;
      Dec(FitGap);
    end;
  end;


  if HandleAllocated then
  begin
    DefaultRowHeight := DataBarsArea.GetActualRowHeight;
    for i := 0 to RowCount-1 do
      if i <> FAllDayRowIndex then
        RowHeights[i] := DefaultRowHeight;

    if TopGridLineCount > 0 then
      RowHeights[FTopGridLineIndex] := 1;

    if FResourceAxisPos >= 0 then
      RowHeights[FResourceAxisPos] := ResourceCaptionArea.GetActualSize;

    if FDayNameBarPos >= 0 then
      RowHeights[FDayNameBarPos] := DayNameArea.GetActualSize;
  end;


  if HoursColArea.Visible then
  begin
    Cell[0,0].Value := Date;

    AStartHour := Round(FGridStartWorkingTime / EncodeTime(1,0,0,0));

    MergeCells(0,0, 0, FixedRowCount-1);
    for i := 0 to (FRolRowCount-1) div 2 do
    begin
      ir := FixedRowCount+i*2;
      MergeCells(0,ir, 0,1);
      Cell[0,ir].Value := FormatFloat('00', i+AStartHour) + ':00';
    end;
  end;

  if FAllDayLinesCount > 1
    then AnAllDayLinesCount := FAllDayLinesCount
    else AnAllDayLinesCount := 1;


  if (FResourceAxisPos >= 0) and (FBarsPerRes > 1) then
  begin
    for i := 0 to ColGroups-1 do
    begin
      ir := FixedColCount + FBarsPerRes * i;
      MergeCells(ir,FResourceAxisPos, FBarsPerRes-1-FInterResourceCols,0);
   end;
  end;


  RowHeights[FAllDayRowIndex] :=
    AnAllDayLinesCount * FDefaultTimeSpanBoxHeight + (FDefaultTimeSpanBoxHeight * 2 div 3);
end;

function TPlannerWeekViewEh.AdjustDate(const Value: TDateTime): TDateTime;
begin
  if FDayCols = 1
    then Result := DateOf(Value)
    else Result := StartOfTheWeek(Value);
end;

procedure TPlannerWeekViewEh.CheckDrawCellBorder(ACol, ARow: Integer;
  BorderType: TGridCellBorderTypeEh; var IsDraw: Boolean;
  var BorderColor: TColor; var IsExtent: Boolean);
var
  DataCell: TGridCoord;
  DataRow: Integer;
  Resource: TPlannerResourceEh;
  ResNo: Integer;
begin
  inherited CheckDrawCellBorder(ACol, ARow, BorderType, IsDraw, BorderColor, IsExtent);

  Resource := nil;
  DataCell := GridCoordToDataCoord(GridCoord(ACol, ARow));

  if ((ACol-FDataColsOffset >= 0) or ((ACol-FDataColsOffset = -1 ) and (BorderType = cbtRightEh))) and
     (PlannerDataSource <> nil) and
     (PlannerDataSource.Resources.Count > 0) then
  begin
    if IsInterResourceCell(ACol, ARow) then
    begin
      if BorderType in [cbtTopEh, cbtBottomEh]  then
        IsDraw := False;
      Resource := PlannerDataSource.Resources[(ACol-FDataColsOffset) div FBarsPerRes];
    end else
    begin
      ResNo := (ACol-FDataColsOffset) div FBarsPerRes;
      if (ResNo >= 0) and (ResNo < PlannerDataSource.Resources.Count) then
        Resource := PlannerDataSource.Resources[ResNo];
    end;
    if (Resource <> nil) and (Resource.DarkLineColor <> clDefault) then
        BorderColor := Resource.DarkLineColor;
  end;

  if IsDraw and (DataCell.X >= 0) and (DataCell.Y >= 0) then
  begin
    DataRow := DataCell.Y;// - FixedRowCount;
    if (BorderType in [cbtTopEh, cbtBottomEh]) and (DataRow mod 2 = 0) then
      if (Resource <> nil) and (Resource.BrightLineColor <> clDefault)
        then BorderColor := Resource.BrightLineColor
        else BorderColor := GridLineParams.GetPaleColor;
  end;
end;

function TPlannerWeekViewEh.CheckHitSpanItem(Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer): TTimeSpanDisplayItemEh;
var
  Pos: TPoint;
  i: Integer;
  SpanItem: TTimeSpanDisplayItemEh;
  SpanDrawRect: TRect;
  ACellRect: TRect;
begin
  Result := nil;
  if (X >= HorzAxis.FixedBoundary) and (X < HorzAxis.ContraStart) and
     (Y >= VertAxis.FixedBoundary) and (Y < VertAxis.ContraStart) then
  begin
    Pos := Point(X, Y);
    for i := 0 to InDayListCount-1 do
    begin
      SpanItem := InDayList[i];
      SpanItem.GetInGridDrawRect(SpanDrawRect);
      if PtInRect(SpanDrawRect, Pos) then
      begin
        Result := SpanItem;
        Break;
      end;
    end;
  end
  else if (X >= HorzAxis.FixedBoundary) and (X < HorzAxis.ContraStart) and
          (Y >= RowHeights[0]) and (Y < VertAxis.FixedBoundary) then
  begin
    ACellRect := CellRect(0, FAllDayRowIndex);
    Pos := Point(X, Y - ACellRect.Top);
    for i := 0 to AllDayListCount-1 do
    begin
      SpanItem := AllDayList[i];
      SpanItem.GetInGridDrawRect(SpanDrawRect);
      if PtInRect(SpanDrawRect, Pos) then
      begin
        Result := SpanItem;
        Break;
      end;
    end;
  end;
end;

procedure TPlannerWeekViewEh.GetCellType(ACol, ARow: Integer;
  var CellType: TPlannerViewCellTypeEh; var ALocalCol, ALocalRow: Integer);
begin
  if (ACol = FHoursBarIndex) and (ARow < FixedRowCount) then
  begin
    CellType := pctTopLeftCellEh;
    ALocalCol := ACol;
    ALocalRow := ARow;
  end else if ACol = FHoursBarIndex then
  begin
    CellType := pctTimeCellEh;
    ALocalCol := ACol;
    ALocalRow := ARow - FixedRowCount;
  end else if (ACol > FHoursBarIndex) and (ARow = FResourceAxisPos) then
  begin
    if IsInterResourceCell(ACol, ARow)
      then CellType := pctInterResourceSpaceEh
      else CellType := pctResourceCaptionCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end else if (ACol > FHoursBarIndex) and (ARow = FDayNameBarPos) then
  begin
    if IsInterResourceCell(ACol, ARow)
      then CellType := pctInterResourceSpaceEh
      else CellType := pctDayNameCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end else
  begin
    if ARow = FAllDayRowIndex then
    begin
      CellType := pctAlldayDataCellEh;
      ALocalRow := 0;
    end else
    begin
      CellType := pctDataCellEh;
      ALocalRow := ARow - FixedRowCount;
    end;
    ALocalCol := ACol - FixedColCount;
  end;
end;

function TPlannerWeekViewEh.DrawLongDayNames: Boolean;
begin
  if FDayCols = 1
    then Result := True
    else Result := False;
end;

function TPlannerWeekViewEh.DrawMonthDayWithWeekDayName: Boolean;
begin
  if FDayCols = 1
    then Result := False
    else Result := True;
end;

procedure TPlannerWeekViewEh.GetViewPeriod(var AStartDate, AEndDate: TDateTime);
begin
  AStartDate := StartDate;
  AEndDate := AStartDate + FDayCols;
end;

function TPlannerWeekViewEh.NewItemParams(var StartTime,
  EndTime: TDateTime; var Resource: TPlannerResourceEh): Boolean;
begin
  StartTime := CellToDateTime(Col, Row);
  EndTime := StartTime + EncodeTime(0, 30, 0, 0);
  Resource :=  GetResourceAtCell(Col, Row);
  Result := True;
end;

function TPlannerWeekViewEh.NextDate: TDateTime;
begin
  Result := StartDate + FDayCols;
end;

function TPlannerWeekViewEh.PriorDate: TDateTime;
begin
  Result := StartDate - FDayCols;
end;

function TPlannerWeekViewEh.AppendPeriod(ATime: TDateTime; Periods: Integer): TDateTime;
begin
  Result := ATime + FDayCols * Periods;
end;

procedure TPlannerWeekViewEh.Resize;
begin
  inherited Resize;
  if HandleAllocated and ActiveMode then
    SetDisplayPosesSpanItems;
end;

function TPlannerWeekViewEh.SelectCell(ACol, ARow: Integer): Boolean;
begin
  if IsInterResourceCell(ACol, ARow) then
    Result := False
  else
    Result := inherited SelectCell(ACol, ARow);
end;

procedure TPlannerWeekViewEh.SetDisplayPosesSpanItems;
var
  ASpanRect: TRect;
  SpanItem: TTimeSpanDisplayItemEh;
  i: Integer;
  ACol: Integer;
  StartX: Integer;
  ResourceOffset: Integer;
  ADate: TDateTime;
  AEndCol: Integer;
  d: Integer;
  ARightRect: TRect;
begin
  if not ActiveMode then Exit;

  SetResOffsets;

//  StartX := ColWidths[0];
  StartX := 0;
  for ACol := FDataColsOffset to ColCount-1 do
  begin

    for i := 0 to FInDayList.Count-1 do
    begin
      SpanItem := TTimeSpanDisplayItemEh(FInDayList[i]);
      if SpanItem.GridColNum = ACol - FDataColsOffset then
      begin
        ASpanRect.Left := StartX + 10;
        ASpanRect.Right := ASpanRect.Left + ColWidths[ACol] - 20;
        ASpanRect.Top := {VertAxis.FixedBoundary + }SpanItem.StartGridRollPos;
        if ASpanRect.Top < 0 then
          ASpanRect.Top := 0;
        ASpanRect.Bottom := {VertAxis.FixedBoundary + }SpanItem.StopGridRolPosl;
        if ASpanRect.Bottom > VertAxis.RolLen then
          ASpanRect.Bottom := VertAxis.RolLen;


        ResourceOffset := GetGridOffsetForResource(SpanItem.PlanItem.Resource);
        OffsetRect(ASpanRect, ResourceOffset, 0);

        CalcRectForInCellCols(SpanItem, ASpanRect);

        SpanItem.FBoundRect := ASpanRect;
      end;
    end;

    StartX := StartX + ColWidths[ACol];
  end;


  if Length(FResourcesView) > 1
    then AEndCol := FBarsPerRes-2
    else AEndCol := FBarsPerRes-1;

//  ADate := StartDate;
  for i := 0 to AllDayListCount-1 do
  begin
    SpanItem := TTimeSpanDisplayItemEh(AllDayList[i]);
    ResourceOffset := GetGridOffsetForResource(SpanItem.PlanItem.Resource);
    ADate := StartDate;
    StartX := HorzAxis.FixedBoundary;
    for d := 0 to AEndCol do
    begin
      if (d = 0) and (SpanItem.PlanItem.StartTime < ADate) then
      begin
        SpanItem.FBoundRect.Left := StartX;
        SpanItem.FDrawBackOutInfo := True;
      end else if (SpanItem.PlanItem.StartTime >= ADate) and (SpanItem.PlanItem.StartTime < ADate + 1) then
      begin
        SpanItem.FBoundRect.Left := StartX + 8;
      end;

      if (SpanItem.PlanItem.EndTime >= ADate) and (SpanItem.PlanItem.EndTime < ADate + 1) then
      begin
        SpanItem.FBoundRect.Right := StartX + ColWidths[d + FDataColsOffset] - 8;
//        SpanItem.FGridBoundRect.Top := FDefaultTimeSpanBoxHeight * i;
//        SpanItem.FGridBoundRect.Bottom := FDefaultTimeSpanBoxHeight * i + FDefaultTimeSpanBoxHeight;
        SpanItem.FBoundRect.Top := FDefaultTimeSpanBoxHeight * SpanItem.FInCellFromCol;
        SpanItem.FBoundRect.Bottom := SpanItem.FBoundRect.Top + FDefaultTimeSpanBoxHeight;
        OffsetRect(SpanItem.FBoundRect, ResourceOffset, 0);
        Break;
      end;

      ADate := ADate + 1;
      StartX := StartX + ColWidths[d + FDataColsOffset];
    end;
  end;

  ADate := StartDate + AEndCol + 1;
  ARightRect := CellRect(AEndCol + FDataColsOffset, 0);
  for i := 0 to AllDayListCount-1 do
  begin
    SpanItem := TTimeSpanDisplayItemEh(AllDayList[i]);
    ResourceOffset := GetGridOffsetForResource(SpanItem.PlanItem.Resource);
    if (SpanItem.PlanItem.EndTime >= ADate) then
    begin
      SpanItem.FBoundRect.Right := ARightRect.Right;
      SpanItem.FDrawForwardOutInfo := True;
//      SpanItem.FGridBoundRect.Top := FDefaultTimeSpanBoxHeight * i;
//      SpanItem.FGridBoundRect.Bottom := FDefaultTimeSpanBoxHeight * i + FDefaultTimeSpanBoxHeight;
      SpanItem.FBoundRect.Top := FDefaultTimeSpanBoxHeight * SpanItem.FInCellFromCol;
      SpanItem.FBoundRect.Bottom := SpanItem.FBoundRect.Top + FDefaultTimeSpanBoxHeight;
      OffsetRect(SpanItem.FBoundRect, ResourceOffset, 0);
    end;
  end;

end;

procedure TPlannerWeekViewEh.SetGroupPosesSpanItems(Resource: TPlannerResourceEh);
var
  i: Integer;
  SpanItem: TTimeSpanDisplayItemEh;
  CurStack: TList;
  CurList: TList;
//  CutStackSize: Integer;
  CurColbarCount: Integer;
  FullEmpty: Boolean;

type
  TGetEndTime = function (SpanItem: TTimeSpanDisplayItemEh): TDateTime;

  procedure CheckPushOutStack(ATime: TDateTime; AGetEndTimeProg: TGetEndTime);
  var
    i: Integer;
    StackSpanItem: TTimeSpanDisplayItemEh;
  begin
    for i := 0 to CurStack.Count-1 do
    begin
      StackSpanItem := TTimeSpanDisplayItemEh(CurStack[i]);
      if (StackSpanItem <> nil) and (AGetEndTimeProg(StackSpanItem) <= ATime) then
      begin
        CurList.Add(CurStack[i]);
        CurStack[i] := nil;
      end;
    end;

    FullEmpty := True;
    for i := 0 to CurStack.Count-1 do
      if CurStack[i] <> nil then
      begin
        FullEmpty := False;
        Break;
      end;

    if FullEmpty then
    begin
      CurColbarCount := 1;
      CurList.Clear;
      CurStack.Clear;
    end;
  end;

  procedure PushInStack(ASpanItem: TTimeSpanDisplayItemEh);
  var
    i: Integer;
    StackSpanItem: TTimeSpanDisplayItemEh;
    PlaceFound: Boolean;
  begin
    PlaceFound := False;
    if CurStack.Count > 0 then
    begin
      for i := 0 to CurStack.Count-1 do
      begin
        StackSpanItem := TTimeSpanDisplayItemEh(CurStack[i]);
        if StackSpanItem = nil then
        begin
          ASpanItem.FInCellCols := CurColbarCount;
          ASpanItem.FInCellFromCol := i;
          ASpanItem.FInCellToCol := i;
          CurStack[i] := ASpanItem;
          PlaceFound := True;
          Break;
        end;
      end;
    end;
    if not PlaceFound then
    begin
      if CurStack.Count > 0 then
      begin
        CurColbarCount := CurColbarCount + 1;
        for i := 0 to CurStack.Count-1 do
          TTimeSpanDisplayItemEh(CurStack[i]).FInCellCols := CurColbarCount;
        for i := 0 to CurList.Count-1 do
          TTimeSpanDisplayItemEh(CurList[i]).FInCellCols := CurColbarCount;
      end;
      ASpanItem.FInCellCols := CurColbarCount;
      ASpanItem.FInCellFromCol := CurColbarCount-1;
      ASpanItem.FInCellToCol := CurColbarCount-1;
      CurStack.Add(ASpanItem);
    end;
  end;

  function GetSpanEndTime(SpanItem: TTimeSpanDisplayItemEh): TDateTime;
  begin
    Result := SpanItem.EndTime;
  end;

  function GetPlanEndTime(SpanItem: TTimeSpanDisplayItemEh): TDateTime;
  begin
    Result := SpanItem.PlanItem.EndTime;
  end;

begin
  CurStack := TList.Create;
  CurList := TList.Create;
  CurColbarCount := 1;

  for i := 0 to FInDayList.Count-1 do
  begin
    SpanItem := InDayList[i];
    if SpanItem.PlanItem.Resource = Resource then
    begin
      CheckPushOutStack(SpanItem.PlanItem.StartTime, @GetPlanEndTime);
      PushInStack(SpanItem);
    end;
  end;


  CurStack.Clear;
  CurList.Clear;
  CurColbarCount := 1;

  for i := 0 to FAllDayList.Count-1 do
  begin
    SpanItem := AllDayList[i];
    if SpanItem.PlanItem.Resource = Resource then
    begin
      CheckPushOutStack(SpanItem.StartTime, @GetSpanEndTime);
      PushInStack(SpanItem);
    end;
  end;

  CurStack.Free;
  CurList.Free;
end;

procedure TPlannerWeekViewEh.SetMinDayColWidth(const Value: Integer);
begin
  if FMinDayColWidth <> Value then
  begin
    FMinDayColWidth := Value;
    GridLayoutChanged;
  end;
end;

procedure TPlannerWeekViewEh.SetResOffsets;
var
  i: Integer;
begin
  for i := 0 to Length(FResourcesView)-1 do
    if i*FBarsPerRes < HorzAxis.RolCelCount then
    begin
      if i < ResourcesCount
        then FResourcesView[i].Resource := PlannerDataSource.Resources[i]
        else FResourcesView[i].Resource := nil;
      FResourcesView[i].GridOffset := HorzAxis.RolLocCelPosArr[i*FBarsPerRes];
      FResourcesView[i].GridStartAxisBar := i*FBarsPerRes;
    end;
end;

function TPlannerWeekViewEh.GetGridOffsetForResource(Resource: TPlannerResourceEh): Integer;
var
  i: Integer;
begin
  Result := 0;
  if PlannerDataSource <> nil then
  begin
    for i := 0 to Length(FResourcesView)-1 do
      if    ((i < PlannerDataSource.Resources.Count) and (PlannerDataSource.Resources[i] = Resource))
         or ((i = Length(FResourcesView)-1) and (Resource = nil)) then
      begin
        Result := FResourcesView[i].GridOffset;
        Break;
      end;
  end;
end;

procedure TPlannerWeekViewEh.SortPlanItems;
begin
  FillSpecDaysList;

  FInDayList.Sort(CompareSpanItemFuncByPlan);
  FAllDayList.Sort(CompareSpanItemFuncByPlan);
end;

function TPlannerWeekViewEh.IsPlanItemHitAllGridArea(APlanItem: TPlannerDataItemEh): Boolean;
var
  CheckDay: TDateTime;
begin
  Result := not APlanItem.InsideDay;
  if not Result then
  begin
    CheckDay := DateOf(APlanItem.EndTime);
    if APlanItem.EndTime < CheckDay + FGridStartWorkingTime then
    begin
      Result := True;
      Exit;
    end;
    CheckDay := DateOf(APlanItem.StartTime);
    if APlanItem.StartTime >= CheckDay + FGridStartWorkingTime + FGridWorkingTimeLength then
      Result := True;
  end;
end;

procedure TPlannerWeekViewEh.FillSpecDaysList;
var
  i: Integer;
  SpanItem: TTimeSpanDisplayItemEh;
begin
  FInDayList.Clear;
  FAllDayList.Clear;

  for i := 0 to SpanItemsCount-1 do
  begin
    SpanItem := SpanItems[i];
    if not IsPlanItemHitAllGridArea(SpanItem.PlanItem) then
    begin
      FInDayList.Add(SpanItem);
      SpanItem.FAllowedInteractiveChanges :=
        [sichSpanTopSizingEh, sichSpanButtomSizingEh, sichSpanMovingEh];
      SpanItem.FTimeOrientation := toVerticalEh;
    end else
    begin
      FAllDayList.Add(SpanItem);
      SpanItem.FHorzLocating := brrlWindowClientEh;
      SpanItem.FVertLocating := brrlWindowClientEh;
//      SpanItem.FInteractiveChangeAllowed := False;
      SpanItem.FAllowedInteractiveChanges :=
        [sichSpanLeftSizingEh, sichSpanRightSizingEh, sichSpanMovingEh];
      SpanItem.FTimeOrientation := toHorizontalEh;
      SpanItem.FAlignment := taCenter;
      SpanItem.StartTime := DateOf(SpanItem.PlanItem.StartTime);
      SpanItem.EndTime := DateOf(SpanItem.PlanItem.EndTime);
      if SpanItem.EndTime <> SpanItem.PlanItem.EndTime then
        SpanItem.EndTime := IncDay(SpanItem.EndTime);
    end;
  end;
end;

function TPlannerWeekViewEh.GetAllDayListItem(Index: Integer): TTimeSpanDisplayItemEh;
begin
  Result := TTimeSpanDisplayItemEh(FAllDayList[Index]);
end;

function TPlannerWeekViewEh.GetInDayListItem(Index: Integer): TTimeSpanDisplayItemEh;
begin
  Result := TTimeSpanDisplayItemEh(FInDayList[Index]);
end;

function TPlannerWeekViewEh.GetAllDayListCount: Integer;
begin
  Result := FAllDayList.Count;
end;

function TPlannerWeekViewEh.GetInDayListCount: Integer;
begin
  Result := FInDayList.Count;
end;

procedure TPlannerWeekViewEh.StartDateChanged;
begin
  inherited StartDateChanged;
  GridLayoutChanged;
end;

procedure TPlannerWeekViewEh.CalcPosByPeriod(AColStartTime, AStartTime, AEndTime: TDateTime;
  var AStartGridPos, AStopGridPos: Integer);
begin
  AStartGridPos := TimeToGridRolPos(AColStartTime, AStartTime);
  AStopGridPos := TimeToGridRolPos(AColStartTime, AEndTime);
end;

procedure TPlannerWeekViewEh.InitSpanItem(
  ASpanItem: TTimeSpanDisplayItemEh);
var
  i: Integer;
  StartTime: TDateTime;
begin
  for i := 0 to FDayCols-1 do
  begin
    StartTime := ColsStartTime[i];
    if (ASpanItem.PlanItem.StartTime >= StartTime) and
      (ASpanItem.PlanItem.StartTime < StartTime + 1) then
    begin
      CalcPosByPeriod(StartTime, ASpanItem.PlanItem.StartTime, ASpanItem.PlanItem.EndTime,
        ASpanItem.FStartGridRollPos, ASpanItem.FStopGridRolPos);
      ASpanItem.FGridColNum := i;
    end;
  end;
end;

procedure TPlannerWeekViewEh.CellMouseClick(const Cell: TGridCoord;
  Button: TMouseButton; Shift: TShiftState; const ACellRect: TRect;
  const GridMousePos, CellMousePos: TPoint);
begin
  inherited CellMouseClick(Cell, Button, Shift, ACellRect, GridMousePos, CellMousePos);
  if (Cell.X >= FDataColsOffset) and (Cell.Y = FDayNameBarPos) then
    DayNamesCellClick(Cell, Button, Shift, ACellRect, GridMousePos, CellMousePos);
end;

procedure TPlannerWeekViewEh.PaintSpanItems;
begin
  DrawInDaySpanItems;
  DrawAllDaySpanItems;
end;

procedure TPlannerWeekViewEh.DrawInDaySpanItems;
var
  SpanItem: TTimeSpanDisplayItemEh;
  i: Integer;
  SpanDrawRect: TRect;
  RestRgn: HRgn;
  RestRect: TRect;
  GridClientRect: TRect;
begin
  GridClientRect := ClientRect;
  if InDayListCount > 0 then
  begin
    RestRect := Rect(HorzAxis.FixedBoundary, VertAxis.FixedBoundary, HorzAxis.ContraStart, VertAxis.ContraStart);
    RestRgn := SelectClipRectangleEh(Canvas, RestRect);
    try
      for i := 0 to InDayListCount-1 do
      begin
        SpanItem := InDayList[i];
        SpanDrawRect := SpanItem.BoundRect;
        OffsetRect(SpanDrawRect,
          HorzAxis.FixedBoundary-HorzAxis.RolStartVisPos,
          VertAxis.FixedBoundary-VertAxis.RolStartVisPos);
        if RectsIntersected(GridClientRect, SpanDrawRect) then
          DrawSpanItem(SpanItem, SpanDrawRect);
      end;
    finally
      RestoreClipRectangleEh(Canvas, RestRgn);
    end;
  end;
end;

procedure TPlannerWeekViewEh.DrawAllDaySpanItems;
var
  SpanItem: TTimeSpanDisplayItemEh;
  i: Integer;
  SpanDrawRect: TRect;
  AllDayRow: Integer;
  GridClientRect: TRect;
begin
  GridClientRect := ClientRect;
  if AllDayListCount > 0 then
  begin
//    RestRect := Rect(HorzAxis.FixedBoundary, VertAxis.FixedBoundary, HorzAxis.ContraStart, VertAxis.ContraStart);
//    RestRgn := SelectClipRectangleEh(Canvas, RestRect);
    for i := 0 to AllDayListCount-1 do
    begin
      SpanItem := AllDayList[i];
      SpanDrawRect := SpanItem.BoundRect;
      if FDayNameBarPos >= 0
        then AllDayRow := FDayNameBarPos + 1
        else AllDayRow := FResourceAxisPos + 1;


      OffsetRect(SpanDrawRect, -HorzAxis.RolStartVisPos, VertAxis.GetFixedCelPos(AllDayRow));
      if RectsIntersected(GridClientRect, SpanDrawRect) then
        DrawSpanItem(SpanItem, SpanDrawRect);
    end;
//    RestoreClipRectangleEh(Canvas, RestRgn);
  end;
end;

procedure TPlannerWeekViewEh.DayNamesCellClick(const Cell: TGridCoord;
  Button: TMouseButton; Shift: TShiftState; const ACellRect: TRect;
  const GridMousePos, CellMousePos: TPoint);
var
  DayNoDate: TDateTime;
begin
  DayNoDate := StartDate + (Cell.X-1) mod FBarsPerRes;
  if DateOf(DayNoDate) <> DateOf(CurrentTime) then
    CurrentTime := DayNoDate;

  PlannerControl.ViewMode := pdrmDayEh;
end;

function TPlannerWeekViewEh.CellToDateTime(ACol, ARow: Integer): TDateTime;
var
  DataCell: TGridCoord;
begin
  DataCell := GridCoordToDataCoord(GridCoord(ACol, ARow));

  Result := StartDate;
  if DataCell.X >= 0 then
    Result := NormalizeDateTime(IncDay(Result, DataCell.X mod FBarsPerRes * FColDaysLength ));
  if DataCell.Y >= 0 then
    Result := NormalizeDateTime(IncMinute(Result, DataCell.Y * FRowMinutesLength) + FGridStartWorkingTime);
end;

function TPlannerWeekViewEh.TimeToGridRolPos(AColStartTime, ATime: TDateTime): Integer;
var
  StartWorkingTimeGridLen: Integer;
begin
  if ATime < AColStartTime then
    Result := 0
  else if ATime >= AColStartTime + 1 then
    Result := VertAxis.RolLen
  else
    Result := Round(TimeOf(ATime) / EncodeTime(0, 30, 0, 0) * DefaultRowHeight);
  StartWorkingTimeGridLen := Round(FGridStartWorkingTime / EncodeTime(0, 30, 0, 0)) * DefaultRowHeight;
  Result := Result - StartWorkingTimeGridLen;
end;

procedure TPlannerWeekViewEh.UpdateDummySpanItemSize(MousePos: TPoint);
var
  FixedMousePos: TPoint;
  ACell: TGridCoord;
  ANewTime: TDateTime;
  ATimeLen: TDateTime;
  ResViewIdx: Integer;
  AResource: TPlannerResourceEh;
  OldInAllDays, NewInAllDays: Boolean;
begin
  FixedMousePos := MousePos;
  if FPlannerState in [psSpanMovingEh, psSpanTestMovingEh] then
  begin
    FixedMousePos.X := FixedMousePos.X + FTopLeftSpanShift.cx;
    FixedMousePos.Y := FixedMousePos.Y + FTopLeftSpanShift.cy;
  end;
  if FPlannerState = psSpanButtomSizingEh then
  begin
    ACell := CalcCoordFromPoint(FixedMousePos.X, FixedMousePos.Y+DefaultRowHeight div 2);
    ANewTime := CellToDateTime(ACell.X, ACell.Y);
    if (FDummyPlanItem.StartTime < ANewTime) and (FDummyPlanItem.EndTime <> ANewTime) then
    begin
      FDummyCheckPlanItem.Assign(FDummyPlanItem);
      FDummyCheckPlanItem.EndTime := ANewTime;
      CheckSetDummyPlanItem(FDummyPlanItemFor, FDummyCheckPlanItem);
      PlannerDataSourceChanged();
    end;
  end else if FPlannerState = psSpanTopSizingEh then
  begin
    ACell := CalcCoordFromPoint(FixedMousePos.X, FixedMousePos.Y+DefaultRowHeight div 2);
    ANewTime := CellToDateTime(ACell.X, ACell.Y);
    if (FDummyPlanItem.EndTime > ANewTime) and (FDummyPlanItem.StartTime <> ANewTime) then
    begin
      FDummyCheckPlanItem.Assign(FDummyPlanItem);
      FDummyCheckPlanItem.StartTime := ANewTime;
      CheckSetDummyPlanItem(FDummyPlanItemFor, FDummyCheckPlanItem);
      PlannerDataSourceChanged();
    end;
  end else if FPlannerState in [psSpanMovingEh, psSpanTestMovingEh] then
  begin
    ACell := CalcCoordFromPoint(FixedMousePos.X, FixedMousePos.Y);
    ANewTime := CellToDateTime(ACell.X, ACell.Y);
    ResViewIdx := GetResourceViewAtCell(ACell.X, ACell.Y);
    if ResViewIdx >= 0
      then AResource := FResourcesView[ResViewIdx].Resource
      else AResource := FDummyPlanItem.Resource;

    if ACell.Y < FixedRowCount
      then NewInAllDays := True
      else NewInAllDays := False;
    if IsPlanItemHitAllGridArea(FDummyPlanItem)
      then OldInAllDays := True
      else OldInAllDays := False;
    if (NewInAllDays = OldInAllDays) and (NewInAllDays = False) then
    begin
      if (FDummyPlanItem.StartTime <> ANewTime) or
         (AResource <> FDummyPlanItem.Resource) then
      begin
        ATimeLen :=  FDummyPlanItem.EndTime - FDummyPlanItem.StartTime;
        FDummyCheckPlanItem.Assign(FDummyPlanItem);
        FDummyCheckPlanItem.StartTime := ANewTime;
        FDummyCheckPlanItem.EndTime := FDummyCheckPlanItem.StartTime + ATimeLen;
        FDummyCheckPlanItem.Resource := AResource;
        CheckSetDummyPlanItem(FDummyPlanItemFor, FDummyCheckPlanItem);

        PlannerDataSourceChanged;
      end;
    end else if (NewInAllDays = OldInAllDays) and (NewInAllDays = True) then
    begin
      if (FDummyPlanItem.StartTime <> ANewTime) or
         (AResource <> FDummyPlanItem.Resource) then
      begin
        ATimeLen :=  FDummyPlanItem.EndTime - FDummyPlanItem.StartTime;
        FDummyPlanItem.StartTime := ANewTime;
        FDummyPlanItem.EndTime := FDummyPlanItem.StartTime + ATimeLen;
        FDummyPlanItem.Resource := AResource;
        PlannerDataSourceChanged;
      end;
    end else if (NewInAllDays = True) and (OldInAllDays = False) then
    begin
      if IsPlanItemHitAllGridArea(FDummyPlanItemFor) then
      begin
        FDummyPlanItem.StartTime := FDummyPlanItemFor.StartTime;
        FDummyPlanItem.EndTime := FDummyPlanItemFor.EndTime;
        FDummyPlanItem.Resource := AResource;
        PlannerDataSourceChanged;
      end else
      begin
        FDummyPlanItem.StartTime := FDummyPlanItemFor.StartTime;
        FDummyPlanItem.EndTime := FDummyPlanItemFor.EndTime;
        FDummyPlanItem.Resource := AResource;
        FDummyPlanItem.AllDay := True;
        PlannerDataSourceChanged;
      end;
    end else if (NewInAllDays = False) and (OldInAllDays = True) then
    begin
      if IsPlanItemHitAllGridArea(FDummyPlanItemFor) then
      begin
        FDummyPlanItem.StartTime := ANewTime;
        FDummyPlanItem.EndTime := IncHour(FDummyPlanItem.StartTime, 1);
        FDummyPlanItem.AllDay := False;
        FDummyPlanItem.Resource := AResource;
        PlannerDataSourceChanged;
      end else
      begin
        ATimeLen :=  FDummyPlanItem.EndTime - FDummyPlanItem.StartTime;
        FDummyPlanItem.StartTime := ANewTime;
        FDummyPlanItem.EndTime := FDummyPlanItem.StartTime + ATimeLen;
        FDummyPlanItem.Resource := AResource;
        FDummyPlanItem.AllDay := False;
        PlannerDataSourceChanged;
      end;
    end;
    ShowMoveHintWindow(FDummyPlanItem, MousePos);
  end;
end;

function TPlannerWeekViewEh.GetColsStartTime(ADayCol: Integer): TDateTime;
begin
  Result := StartDate + ADayCol;
end;

function TPlannerWeekViewEh.GetCoveragePeriodType: TPlannerCoveragePeriodTypeEh;
begin
  Result := pcpWeekEh;
end;

function TPlannerWeekViewEh.GetPeriodCaption: String;
var
  StartDate, EndDate: TDateTime;
begin
  GetViewPeriod(StartDate, EndDate);
  Result := DateRangeToStr(StartDate, EndDate);
end;

procedure TPlannerWeekViewEh.SetGridShowHours;
var
  h, m, s, ms: Word;
begin
  if ShowWorkingTimeOnly then
  begin
    DecodeTime(FWorkingTimeStart, h, m, s, ms);
    FGridStartWorkingTime := EncodeTime(h, 0, 0, 0);

    if FWorkingTimeEnd > 0 then
    begin
      DecodeTime(FWorkingTimeEnd, h, m, s, ms);
      FGridWorkingTimeLength := EncodeTime(h, 0, 0, 0) - FGridStartWorkingTime;
    end else
      FGridWorkingTimeLength := IncHour(0, 24) - FGridStartWorkingTime;
  end else
  begin
    FGridStartWorkingTime := EncodeTime(0, 0, 0, 0);
    FGridWorkingTimeLength := IncHour(0, 24);
  end;
end;

function TPlannerWeekViewEh.GetResourceAtCell(ACol,
  ARow: Integer): TPlannerResourceEh;
begin
  if (ACol-FDataColsOffset >= 0) and
     (PlannerDataSource <> nil) and
     (PlannerDataSource.Resources.Count > 0) and
     ((ACol-FDataColsOffset) div FBarsPerRes < PlannerDataSource.Resources.Count)
  then
    Result := PlannerDataSource.Resources[(ACol-FDataColsOffset) div FBarsPerRes]
  else
    Result := nil;
end;

function TPlannerWeekViewEh.GetResourceViewAtCell(ACol, ARow: Integer): Integer;
begin
  if ACol < FDataColsOffset then
    Result := -1
  else if IsInterResourceCell(ACol, ARow) then
    Result := -1
  else
    Result := (ACol-FDataColsOffset) div FBarsPerRes;
end;

function TPlannerWeekViewEh.IsDayNameAreaNeedVisible: Boolean;
begin
  Result := True;
end;

function TPlannerWeekViewEh.IsInterResourceCell(ACol,
  ARow: Integer): Boolean;
begin
  if (ACol-FDataColsOffset >= 0) and
     (Length(FResourcesView) > 1)
  then
    Result := (ACol-FDataColsOffset) mod FBarsPerRes = FBarsPerRes - 1
  else
    Result := False;
end;

procedure TPlannerWeekViewEh.SetShowWorkingTimeOnly(const Value: Boolean);
begin
  if FShowWorkingTimeOnly <> Value then
  begin
    FShowWorkingTimeOnly := Value;
    PlannerDataSourceChanged;
  end;
end;

procedure TPlannerWeekViewEh.SetWorkingTimeEnd(const Value: TTime);
begin
  if FWorkingTimeEnd <> Value then
  begin
    FWorkingTimeEnd := Value;
    PlannerDataSourceChanged;
  end;
end;

procedure TPlannerWeekViewEh.SetWorkingTimeStart(const Value: TTime);
begin
  if FWorkingTimeStart <> Value then
  begin
    FWorkingTimeStart := Value;
    PlannerDataSourceChanged;
  end;
end;

procedure TPlannerWeekViewEh.CalcRolRows;
begin
  FRolRowCount := Round(FGridWorkingTimeLength / EncodeTime(0, 30, 0, 0));
end;

function TPlannerWeekViewEh.DefaultHoursBarSize: Integer;
begin
  if HandleAllocated then
  begin
    Canvas.Font := HoursBarArea.Font;
    Result := Canvas.TextWidth('11111');
  end else
    Result := 45;
end;

function TPlannerWeekViewEh.CreateHoursBarArea: THoursBarAreaEh;
begin
  Result := THoursVertBarAreaEh.Create(Self);
end;

function TPlannerWeekViewEh.GetHoursColArea: THoursVertBarAreaEh;
begin
  Result := THoursVertBarAreaEh(inherited HoursBarArea);
end;

procedure TPlannerWeekViewEh.SetHoursColArea(const Value: THoursVertBarAreaEh);
begin
  inherited HoursBarArea := Value;
end;

function TPlannerWeekViewEh.CreateDayNameArea: TDayNameAreaEh;
begin
  Result := TDayNameVertAreaEh.Create(Self);
end;

function TPlannerWeekViewEh.GetDayNameArea: TDayNameVertAreaEh;
begin
  Result := TDayNameVertAreaEh(inherited DayNameArea);
end;

procedure TPlannerWeekViewEh.SetDayNameArea(const Value: TDayNameVertAreaEh);
begin
  inherited DayNameArea := Value;
end;

function TPlannerWeekViewEh.GetDataBarsArea: TDataBarsVertAreaEh;
begin
  Result := TDataBarsVertAreaEh(inherited DataBarsArea);
end;

procedure TPlannerWeekViewEh.SetDataBarsArea(const Value: TDataBarsVertAreaEh);
begin
  inherited DataBarsArea := Value;
end;

function TPlannerWeekViewEh.GetResourceCaptionArea: TResourceVertCaptionAreaEh;
begin
  Result := TResourceVertCaptionAreaEh(inherited ResourceCaptionArea);
end;

procedure TPlannerWeekViewEh.SetResourceCaptionArea(
  const Value: TResourceVertCaptionAreaEh);
begin
  inherited ResourceCaptionArea :=  Value;
end;

function TPlannerWeekViewEh.CreateResourceCaptionArea: TResourceCaptionAreaEh;
begin
  Result := TResourceVertCaptionAreaEh.Create(Self);
end;

{ TPlannerDayViewEh }

constructor TPlannerDayViewEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FDayCols := 1;
  HorzScrollBar.VisibleMode := sbNeverShowEh;
end;

destructor TPlannerDayViewEh.Destroy;
begin
  inherited Destroy;
end;

function TPlannerDayViewEh.GetCoveragePeriodType: TPlannerCoveragePeriodTypeEh;
begin
  Result := pcpDayEh;
end;

function TPlannerDayViewEh.IsDayNameAreaNeedVisible: Boolean;
begin
  Result := not IsResourceCaptionNeedVisible;
end;

procedure TPlannerDayViewEh.StartDateChanged;
begin
  inherited StartDateChanged;
  FFirstWeekDayNum := DayOfWeek(StartDate);
end;

{ TPlannerMonthViewEh }

constructor TPlannerMonthViewEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FDataColsOffset := 1;
  FDataRowsOffset := 1;
  HorzScrollBar.VisibleMode := sbNeverShowEh;
  VertScrollBar.VisibleMode := sbNeverShowEh;
  FMinDayColWidth := 50;
  FDataColsFor1Res := 7;
  FSortedSpans := TList.Create;
  FWeekColArea := TWeekBarAreaEh.Create(Self);
end;

destructor TPlannerMonthViewEh.Destroy;
begin
  FreeAndNil(FSortedSpans);
  FreeAndNil(FWeekColArea);
  inherited Destroy;
end;

procedure TPlannerMonthViewEh.BuildGridData;
begin
  inherited BuildGridData;
  BuildMonthGridMode;
  RealignGridControls;
end;

procedure TPlannerMonthViewEh.BuildMonthGridMode;
var
  ColWidth, FitGap: Integer;
  RowHeight: Integer;
  i: Integer;
  Groups: Integer;
  ColGroups: Integer;
  ic{, h}: Integer;
  ARecNoColWidth: Integer;
  ADataColCount: Integer;
  AFixedRowCount: Integer;
begin
  ClearGridCells;
  FHoursBarIndex := -1;
  if FWeekColArea = nil then Exit;

  AFixedRowCount := TopGridLineCount;

  if ResourceCaptionArea.Visible then
  begin
    if (PlannerDataSource <> nil)
      then ColGroups := PlannerDataSource.Resources.Count
      else ColGroups := 0;
    if FShowUnassignedResource then
      Inc(ColGroups);
    if ColGroups = 0 then
      ColGroups := 1;
    Groups := ColGroups;
    FResourceAxisPos := AFixedRowCount;
    FDataRowsOffset := AFixedRowCount + 1;
  end else
  begin
    if (PlannerDataSource <> nil)
      then ColGroups := PlannerDataSource.Resources.Count
      else ColGroups := 0;
    if FShowUnassignedResource then
      Inc(ColGroups);
    if ColGroups = 0 then
      ColGroups := 1;
    Groups := ColGroups;
    FResourceAxisPos := -1;
    FDataRowsOffset := AFixedRowCount;
  end;

  if DayNameArea.Visible then
  begin
    Inc(FDataRowsOffset);
    FDayNameBarPos := FDataRowsOffset-1;
  end else
    FDayNameBarPos := -1;

  if WeekColArea.Visible then
  begin
    FixedColCount := 1;
    FDataColsOffset := 1;
    FWeekColIndex := 0;
  end else
  begin
    FixedColCount := 0;
    FDataColsOffset := 0;
    FWeekColIndex := -1;
  end;

  FixedRowCount := FDataRowsOffset;

  ADataColCount := 7 * ColGroups + (ColGroups-1);
  ColCount := ADataColCount + FDataColsOffset;
  RowCount := FixedRowCount + 6;

  SetGridSize(ColCount, RowCount);
  if HandleAllocated then
  begin
    if TopGridLineCount > 0 then
      RowHeights[FTopGridLineIndex] := 1;

    if FResourceAxisPos >= 0 then
      RowHeights[FResourceAxisPos] := ResourceCaptionArea.GetActualSize;

    if FDayNameBarPos >= 0 then
      RowHeights[FDayNameBarPos] := DayNameArea.GetActualSize;
  end;

  if FWeekColIndex >= 0 then
  begin
    ARecNoColWidth := WeekColArea.GetActualSize;
    ColWidths[FWeekColIndex] := ARecNoColWidth;
    FShowWeekNoCaption := CalcShowKeekNoCaption((GridClientHeight - VertAxis.FixedBoundary) div 6);
  end else
  begin
    ARecNoColWidth := 0;
  end;

  if Groups > 0
    then FBarsPerRes := 8
    else FBarsPerRes := 7;

  ColWidth := (GridClientWidth - ARecNoColWidth - (ColGroups-1)*3) div (ADataColCount-(ColGroups-1));
  if ColWidth < MinDayColWidth then
  begin
    ColWidth := MinDayColWidth;
    HorzScrollBar.VisibleMode := sbAutoShowEh;
    FitGap := 0;
  end else
  begin
    HorzScrollBar.VisibleMode := sbNeverShowEh;
    FitGap := (GridClientWidth - ARecNoColWidth - (ColGroups-1)*3) mod (ADataColCount-(ColGroups-1));
  end;

  for i := FDataColsOffset to ColCount-1 do
  begin
    if IsInterResourceCell(i, 0) then
    begin
      ColWidths[i] := 3;
    end else
    begin
      if FitGap > 0
        then ColWidths[i] := ColWidth + 1
        else ColWidths[i] := ColWidth;
      Dec(FitGap);
    end;
  end;

  RowHeight := (GridClientHeight - VertAxis.FixedBoundary) div 6;
  FitGap := (GridClientHeight - VertAxis.FixedBoundary) mod 6;
  for i := FixedRowCount to RowCount-1 do
  begin
    if FitGap > 0
      then RowHeights[i] := RowHeight + 1
      else RowHeights[i] := RowHeight;
    Dec(FitGap);
  end;

  while DayOfWeek(TDateTime(FStartDate)) <> FFirstWeekDayNum do
    FStartDate := FStartDate - 1;

  if FixedColCount >= 0 then
    MergeCells(0, 0, FixedColCount-1, FixedRowCount-1);

  if FResourceAxisPos >= 0 then
  begin
    for i := 0 to ColGroups-1 do
    begin
      ic := FBarsPerRes * i;
      MergeCells(ic+FDataColsOffset, FResourceAxisPos, FBarsPerRes-2, 0);
    end;
  end;

end;

function TPlannerMonthViewEh.AdjustDate(const Value: TDateTime): TDateTime;
var
  y, m, d: Word;
begin
  DecodeDate(Value, y, m, d);
  Result := StartOfTheWeek(EncodeDate(y, m, 1));
end;

procedure TPlannerMonthViewEh.CellMouseClick(const Cell: TGridCoord;
  Button: TMouseButton; Shift: TShiftState; const ACellRect: TRect;
  const GridMousePos, CellMousePos: TPoint);
begin
  inherited CellMouseClick(Cell, Button, Shift, ACellRect, GridMousePos, CellMousePos);
  if (Cell.X < FDataColsOffset) and (Cell.Y >= FDataRowsOffset) then
    WeekNoCellClick(Cell, Button, Shift, ACellRect, GridMousePos, CellMousePos);
end;

procedure TPlannerMonthViewEh.WeekNoCellClick(const Cell: TGridCoord;
  Button: TMouseButton; Shift: TShiftState; const ACellRect: TRect;
  const GridMousePos, CellMousePos: TPoint);
//var
//  WeekNoDate: TDateTime;
//  NewDateToView: TDateTime;
begin
//  WeekNoDate := StartDate + 7 * (Cell.Y-FDataRowsOffset);
//  if (WeekNoDate <= CurrentTime) and (WeekNoDate + 7 > CurrentTime) then
//   
//  else
//  begin
//    if WeekNoDate > CurrentTime
//      then NewDateToView := IncWeek(CurrentTime, Trunc((WeekNoDate + 6 - CurrentTime)) div 7)
//      else NewDateToView := IncWeek(CurrentTime, Trunc((WeekNoDate - CurrentTime)) div 7);
//    CurrentTime := NewDateToView;
//  end;
//
//  PlannerControl.ViewMode := pdrmWeekEh;
end;

function TPlannerMonthViewEh.CellToDateTime(ACol, ARow: Integer): TDateTime;
var
  DataCell: TGridCoord;
begin
  DataCell := GridCoordToDataCoord(GridCoord(ACol, ARow));
  Result := StartDate + DataCell.Y * (FDataColsFor1Res) + DataCell.X mod FBarsPerRes;
end;

procedure TPlannerMonthViewEh.CheckDrawCellBorder(ACol, ARow: Integer;
  BorderType: TGridCellBorderTypeEh; var IsDraw: Boolean;
  var BorderColor: TColor; var IsExtent: Boolean);
var
  DataCell: TGridCoord;
  DataRow: Integer;
  Resource: TPlannerResourceEh;
begin
  inherited CheckDrawCellBorder(ACol, ARow, BorderType, IsDraw, BorderColor, IsExtent);

  Resource := nil;
  DataCell := GridCoordToDataCoord(GridCoord(ACol, ARow));

  if ((ACol-FDataColsOffset >= 0) or ((ACol-FDataColsOffset = -1 ) and (BorderType = cbtRightEh))) and
     (PlannerDataSource <> nil) and
     (PlannerDataSource.Resources.Count > 0) then
  begin
    Resource := GetResourceAtCell(ACol, ARow);
    if IsInterResourceCell(ACol, ARow) then
    begin
      if BorderType in [cbtTopEh, cbtBottomEh]  then
        IsDraw := False;
    end;
    if (Resource <> nil) and (Resource.DarkLineColor <> clDefault) then
        BorderColor := Resource.DarkLineColor;
  end;

  if IsDraw and (DataCell.X >= 0) and (DataCell.Y >= FixedRowCount) then
  begin
    DataRow := DataCell.Y - FixedRowCount;
    if (BorderType in [cbtTopEh, cbtBottomEh]) and (DataRow mod 2 = 1) then
      if (Resource <> nil) and (Resource.BrightLineColor <> clDefault)
        then BorderColor := Resource.BrightLineColor
        else BorderColor := ApproximateColor(BorderColor, Color, 255 div 2);
  end;
end;

procedure TPlannerMonthViewEh.DrawCell(ACol, ARow: Integer; ARect: TRect;
  State: TGridDrawState);
begin
  inherited DrawCell(ACol, ARow, ARect, State);
end;

procedure TPlannerMonthViewEh.DrawDataCell(ACol, ARow: Integer; ARect: TRect;
  State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
begin
  DrawMonthDayNum(ACol, ARow, ARect, State, DrawArgs);
end;

procedure TPlannerMonthViewEh.GetCellType(ACol, ARow: Integer;
  var CellType: TPlannerViewCellTypeEh; var ALocalCol, ALocalRow: Integer);
begin
  if ACol < FDataColsOffset then
  begin
    if ARow < FDataRowsOffset then
    begin
      CellType := pctTopLeftCellEh;
      ALocalCol := ACol;
      ALocalRow := ARow;
    end else
    begin
      CellType := pctWeekNoCell;
      ALocalCol := ACol;
      ALocalRow := ARow-FixedRowCount;
    end;
  end else if ARow = FResourceAxisPos then
  begin
    if IsInterResourceCell(ACol, ARow)
      then CellType := pctInterResourceSpaceEh
      else CellType := pctResourceCaptionCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end else if ARow = FDayNameBarPos then
  begin
    if IsInterResourceCell(ACol, ARow)
      then CellType := pctInterResourceSpaceEh
      else CellType := pctDayNameCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end else
  begin
    if IsInterResourceCell(ACol, ARow)
      then CellType := pctInterResourceSpaceEh
      else CellType := pctDataCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end;
end;

procedure TPlannerMonthViewEh.DrawMonthDayNum(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; DrawArgs: TPlannerViewCellDrawArgsEh);
begin
  Canvas.Font.Name := DrawArgs.FontName;
  Canvas.Font.Size := DrawArgs.FontSize;
  Canvas.Font.Style := DrawArgs.FontStyle;
  Canvas.Font.Color := DrawArgs.FontColor;
  Canvas.Brush.Color := DrawArgs.BackColor;

  WriteTextEh(Canvas, ARect, True, DrawArgs.HorzMargin, DrawArgs.VertMargin,
    DrawArgs.Text, DrawArgs.Alignment, DrawArgs.Layout, DrawArgs.WordWrap,
    False, 0, 0, False, True);
end;

procedure TPlannerMonthViewEh.DrawWeekNoCell(ACol, ARow: Integer; ARect: TRect;
  State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
begin
  if ARow < FDataRowsOffset then
  begin
    Canvas.Brush.Color := Color;
    Canvas.FillRect(ARect);
  end else
  begin
    Canvas.Font.Name := DrawArgs.FontName;
    Canvas.Font.Size := DrawArgs.FontSize;
    Canvas.Font.Style := DrawArgs.FontStyle;
    WriteTextVerticalEh(Canvas, ARect, True, 0, 0, DrawArgs.Text,
      taCenter, tlCenter, False, False);
  end;
end;

function TPlannerMonthViewEh.DrawMonthDayWithWeekDayName: Boolean;
begin
  Result := False;
end;

procedure TPlannerMonthViewEh.GetWeekNoCellDrawParams(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
var
  WeekNoDate: TDateTime;
  s: String;
begin
  if ARow >= FDataRowsOffset then
  begin
    DrawArgs.FontName := WeekColArea.Font.Name;
    DrawArgs.FontColor := WeekColArea.Font.Color;
    DrawArgs.FontSize := WeekColArea.Font.Size;
    DrawArgs.FontStyle := WeekColArea.Font.Style;
    DrawArgs.BackColor := WeekColArea.GetActualColor;

    WeekNoDate := StartDate + 7 * (ARow-FDataRowsOffset);
    s := '';
    if FShowWeekNoCaption then
      s := SPlannerWeekTextEh + ' ';
    s := s + IntToStr(WeekOfTheYear(WeekNoDate));
    DrawArgs.Value := WeekNoDate;
    DrawArgs.Text := s;
  end;
end;

procedure TPlannerMonthViewEh.GetDataCellDrawParams(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
var
  s: String;
  AddDats: Integer;
  InResCol: Integer;

  function CheckDrawMon: Boolean;
  var
    AYear, AMonth, ADay: Word;
  begin
    Result := False;
    DecodeDate(FStartDate + AddDats, AYear, AMonth, ADay);
    if  (ACol = 0) and (ARow = 1) then
      Result := True
    else if ADay = 1 then
      Result := True;
  end;

begin
  InResCol := (ACol-FDataColsOffset) mod FBarsPerRes;
  AddDats := 7 * (ARow-FDataRowsOffset) + InResCol;
  if CheckDrawMon then
    s := FormatDateTime('D MMM', FStartDate + AddDats)
  else
    s := FormatDateTime('D', FStartDate + AddDats);
  DrawArgs.BackColor := StyleServices.GetSystemColor(Color);
  DrawArgs.FontName := WeekColArea.Font.Name;
  DrawArgs.FontColor := WeekColArea.Font.Color;
  DrawArgs.FontSize := WeekColArea.Font.Size;
  DrawArgs.FontStyle := WeekColArea.Font.Style;

  if FStartDate + AddDats = Date then
    Canvas.Font.Style := [fsBold];

  if not (gdCurrent in State) then
    if IsWorkingDay(CellToDateTime(ACol, ARow))
      then DrawArgs.BackColor := StyleServices.GetSystemColor(Color)
      else DrawArgs.BackColor := ApproximateColor(GridLineColors.GetBrightColor,
        StyleServices.GetSystemColor(Color), 255 div 10 * 9);

  DrawArgs.Value := FStartDate + AddDats;
  DrawArgs.Text := s;
  DrawArgs.HorzMargin := 5;
  DrawArgs.VertMargin := 5;
  DrawArgs.Alignment := taRightJustify;
  DrawArgs.Layout := tlTop;
  DrawArgs.WordWrap := False;
end;


procedure TPlannerMonthViewEh.GetViewPeriod(var AStartDate,
  AEndDate: TDateTime);
begin
  AStartDate := StartDate;
  AEndDate := AStartDate + 6 * 7;
end;

function TPlannerMonthViewEh.IsDayNameAreaNeedVisible: Boolean;
begin
  Result := True;
end;

function TPlannerMonthViewEh.IsInterResourceCell(ACol, ARow: Integer): Boolean;
begin
  if (ACol-FDataColsOffset >= 0) and
     (PlannerDataSource <> nil) and
     (PlannerDataSource.Resources.Count > 0)
  then
    Result := (ACol-FDataColsOffset) mod FBarsPerRes = FBarsPerRes - 1
  else
    Result := False;
end;

function TPlannerMonthViewEh.NextDate: TDateTime;
var
  y,m,d: Word;
begin
  Result := IncMonth(StartDate + 7, 1);
  DecodeDate(Result, y,m,d);
  Result := EncodeDate(y,m,1);
end;

function TPlannerMonthViewEh.PriorDate: TDateTime;
var
  y,m,d: Word;
begin
  Result := IncMonth(StartDate + 7, -1);
  DecodeDate(Result, y,m,d);
  Result := EncodeDate(y,m,1);
end;

function TPlannerMonthViewEh.AppendPeriod(ATime: TDateTime;
  Periods: Integer): TDateTime;
begin
  Result := IncMonth(ATime, Periods);
end;

procedure TPlannerMonthViewEh.SetDisplayPosesSpanItemsForResource(AResource: TPlannerResourceEh; Index: Integer);
var
  ADCol, ADRow: Integer;
  SpanItem: TTimeSpanDisplayItemEh;
  i: Integer;
  ASpanRect: TRect;
  ACellDate: TDateTime;
  AStartCellPos, ACellStartPos: TPoint;
  ASpanWidth: Integer;
  InRowStartDate: TDateTime;
  InRowBoundDate: TDateTime;
  LeftIndent: Integer;
  AStartViewDate, AEndViewDate: TDateTime;
  ASpanBound: Integer;

  function CalcSpanWidth(AStartCol, ALenght: Integer): Integer;
  var
    i: Integer;
  begin
    Result := 0;
    for i:= AStartCol to AStartCol+ALenght-1 do
      Result := Result + ColWidths[i];
  end;

begin
//  AStartCellPos := CellRect(FDataColsOffset, FDataRowsOffset).TopLeft;
  AStartCellPos := Point(HorzAxis.FixedBoundary, VertAxis.FixedBoundary);
  ACellStartPos := AStartCellPos;
  GetViewPeriod(AStartViewDate, AEndViewDate);
  for ADCol := 0 to FDataColsFor1Res - 1 do
  begin
    ACellStartPos.Y := AStartCellPos.Y;
    for ADRow := 0 to RowCount - FDataRowsOffset - 1 do
    begin
      InRowStartDate := CellToDateTime(FDataColsOffset, ADRow + FDataRowsOffset);
      InRowBoundDate := IncDay(InRowStartDate, FDataColsFor1Res);
      for i := 0 to SpanItemsCount-1 do
      begin
        SpanItem := SortedSpan[i];
        if SpanItem.PlanItem.Resource <> AResource then Continue;

        ACellDate := CellToDateTime(ADCol + FDataColsOffset, ADRow + FDataRowsOffset);
        if (ACellDate <= SpanItem.StartTime) and
           (ACellDate + 1 > SpanItem.StartTime) then
        begin
          if SpanItem.PlanItem.StartTime < ACellDate then
          begin
            LeftIndent := 0;
            if (SpanItem.PlanItem.StartTime < StartDate) and (SpanItem.StartTime = StartDate) then
              SpanItem.FDrawBackOutInfo := True;
          end else
          begin
            LeftIndent := 5;
            SpanItem.FAllowedInteractiveChanges := SpanItem.FAllowedInteractiveChanges + [sichSpanLeftSizingEh];
          end;
          ASpanRect.Left := ACellStartPos.X + LeftIndent;
          ASpanRect.Top :=  FDataDayNumAreaHeight + ACellStartPos.Y +
            SpanItem.InCellFromCol * FDefaultLineHeight;
          ASpanWidth := CalcSpanWidth(ADCol + FDataColsOffset, SpanItem.FStopGridRolPos - SpanItem.FStartGridRollPos);
          ASpanRect.Right := ASpanRect.Left + ASpanWidth - LeftIndent;
          if SpanItem.PlanItem.EndTime < InRowBoundDate then
          begin
            ASpanRect.Right := ASpanRect.Right - 5;
            SpanItem.FAllowedInteractiveChanges := SpanItem.FAllowedInteractiveChanges + [sichSpanRightSizingEh];
          end else
          begin
            if (SpanItem.PlanItem.EndTime > AEndViewDate) and (SpanItem.EndTime = AEndViewDate) then
              SpanItem.FDrawForwardOutInfo := True;
          end;
          ASpanRect.Bottom := ASpanRect.Top + FDefaultLineHeight;

          ASpanBound := ACellStartPos.Y + RowHeights[ADRow + FDataRowsOffset];
          if ASpanRect.Bottom > ASpanBound then
            SpanItem.FBoundRect := EmptyRect
          else
          begin
            OffsetRect(ASpanRect, -HorzAxis.FixedBoundary, -VertAxis.FixedBoundary);
            if ResourcesCount > 0 then
              OffsetRect(ASpanRect, FResourcesView[Index].GridOffset, 0);
            SpanItem.FBoundRect := ASpanRect;
          end;
        end;
      end;
      ACellStartPos.Y := ACellStartPos.Y + RowHeights[ADRow + FDataRowsOffset];
    end;
    ACellStartPos.X := ACellStartPos.X + ColWidths[ADCol + FDataColsOffset];
  end;
end;

procedure TPlannerMonthViewEh.SetDisplayPosesSpanItems;
var
  i: Integer;
begin
  SetResOffsets;
  if ResourcesCount > 0 then
  begin
    for i := 0 to PlannerDataSource.Resources.Count-1 do
      SetDisplayPosesSpanItemsForResource(PlannerDataSource.Resources[i], i);
    if FShowUnassignedResource then
      SetDisplayPosesSpanItemsForResource(nil, PlannerDataSource.Resources.Count);
  end else
    SetDisplayPosesSpanItemsForResource(nil, -1);
end;

procedure TPlannerMonthViewEh.SetGroupPosesSpanItems(Resource: TPlannerResourceEh);
var
  i: Integer;
  SpanItem: TTimeSpanDisplayItemEh;
  CurStack: TList;
  CurList: TList;
  CurColbarCount: Integer;
  FullEmpty: Boolean;

  procedure CheckPushOutStack(ABoundPos: Integer);
  var
    i: Integer;
    StackSpanItem: TTimeSpanDisplayItemEh;
  begin
    for i := 0 to CurStack.Count-1 do
    begin
      StackSpanItem := TTimeSpanDisplayItemEh(CurStack[i]);
      if (StackSpanItem <> nil) and (StackSpanItem.FStopGridRolPos <= ABoundPos) then
      begin
        CurList.Add(CurStack[i]);
        CurStack[i] := nil;
      end;
    end;

    FullEmpty := True;
    for i := 0 to CurStack.Count-1 do
      if CurStack[i] <> nil then
      begin
        FullEmpty := False;
        Break;
      end;

    if FullEmpty then
    begin
      CurColbarCount := 1;
      CurList.Clear;
      CurStack.Clear;
    end;
  end;

  procedure PushInStack(ASpanItem: TTimeSpanDisplayItemEh);
  var
    i: Integer;
    StackSpanItem: TTimeSpanDisplayItemEh;
    PlaceFound: Boolean;
  begin
    PlaceFound := False;
    if CurStack.Count > 0 then
    begin
      for i := 0 to CurStack.Count-1 do
      begin
        StackSpanItem := TTimeSpanDisplayItemEh(CurStack[i]);
        if StackSpanItem = nil then
        begin
          ASpanItem.FInCellCols := CurColbarCount;
          ASpanItem.FInCellFromCol := i;
          ASpanItem.FInCellToCol := i;
          CurStack[i] := ASpanItem;
          PlaceFound := True;
          Break;
        end;
      end;
    end;
    if not PlaceFound then
    begin
      if CurStack.Count > 0 then
      begin
        CurColbarCount := CurColbarCount + 1;
        for i := 0 to CurStack.Count-1 do
          TTimeSpanDisplayItemEh(CurStack[i]).FInCellCols := CurColbarCount;
        for i := 0 to CurList.Count-1 do
          TTimeSpanDisplayItemEh(CurList[i]).FInCellCols := CurColbarCount;
      end;
      ASpanItem.FInCellCols := CurColbarCount;
      ASpanItem.FInCellFromCol := CurColbarCount-1;
      ASpanItem.FInCellToCol := CurColbarCount-1;
      CurStack.Add(ASpanItem);
    end;
  end;

begin
  CurStack := TList.Create;
  CurList := TList.Create;
  CurColbarCount := 1;

  for i := 0 to SpanItemsCount-1 do
  begin
    SpanItem := SortedSpan[i];
    if SpanItem.PlanItem.Resource = Resource then
    begin
      CheckPushOutStack(SpanItem.FStartGridRollPos);
      PushInStack(SpanItem);
    end;
  end;
  CurStack.Free;
  CurList.Free;
end;

procedure TPlannerMonthViewEh.SetMinDayColWidth(const Value: Integer);
begin
  if FMinDayColWidth <> Value then
  begin
    FMinDayColWidth := Value;
    GridLayoutChanged;
  end;
end;

procedure TPlannerMonthViewEh.SetResOffsets;
var
  i: Integer;
begin
  for i := 0 to Length(FResourcesView)-1 do
    if i*FBarsPerRes < HorzAxis.RolCelCount then
    begin
      if i < ResourcesCount
        then FResourcesView[i].Resource := PlannerDataSource.Resources[i]
        else FResourcesView[i].Resource := nil;
      FResourcesView[i].GridOffset := HorzAxis.RolLocCelPosArr[i*FBarsPerRes];
      FResourcesView[i].GridStartAxisBar := i*FBarsPerRes;
    end;
end;

function TPlannerMonthViewEh.WeekNoColWidth: Integer;
begin
  Result := 0;
  if not HandleAllocated then Exit;
  Canvas.Font := WeekColArea.Font;
  Result := Canvas.TextHeight('Wg');
  Result := Result + 10;
end;

function TPlannerMonthViewEh.DefaultHoursBarSize: Integer;
begin
  Result := WeekNoColWidth;
end;

function TPlannerMonthViewEh.CalcShowKeekNoCaption(RowHeight: Integer): Boolean;
var
  MaxTextHeight: Integer;
begin
  Result := False;
  if not HandleAllocated then Exit;
  Canvas.Font := Font;
  Canvas.Font.Style := [fsBold];
  MaxTextHeight := Canvas.TextWidth('  WEEK 00  ');
  Result := MaxTextHeight < RowHeight;
end;

procedure TPlannerMonthViewEh.Resize;
begin
  inherited Resize;
  if HandleAllocated and ActiveMode then
  begin
    Canvas.Font := Font;
    FDataDayNumAreaHeight := Canvas.TextHeight('Wg') + 5;

    Canvas.Font := GetPlannerControl.TimeSpanParams.Font;
    FDefaultLineHeight := Canvas.TextHeight('Wg') + 3;

    SetDisplayPosesSpanItems;
  end;
  ResetDayNameFormat(2, 0);
end;

function TPlannerMonthViewEh.GetCoveragePeriodType: TPlannerCoveragePeriodTypeEh;
begin
  Result := pcpMonthEh;
end;

function TPlannerMonthViewEh.GetPeriodCaption: String;
begin
  Result :=
    FormatDateTime('mmmm', StartDate + 7) + ' ' +
    FormatDateTime('yyyy', StartDate + 7);
end;

function TPlannerMonthViewEh.GetResourceAtCell(ACol,
  ARow: Integer): TPlannerResourceEh;
var
  ResIdx: Integer;
begin
  if (ACol-FDataColsOffset >= 0) and
     (PlannerDataSource <> nil) and
     (PlannerDataSource.Resources.Count > 0)
  then
  begin
    ResIdx := (ACol-FDataColsOffset) div FBarsPerRes;
    if ResIdx < PlannerDataSource.Resources.Count then
      Result := PlannerDataSource.Resources[(ACol-FDataColsOffset) div FBarsPerRes]
    else
      Result := nil;
  end else
    Result := nil;
end;

procedure TPlannerMonthViewEh.CalcPosByPeriod(AStartTime, AEndTime: TDateTime;
  var AStartGridPos, AStopGridPos: Integer);
begin
  AStartGridPos := TimeToGridLineRolPos(AStartTime);
  AStopGridPos := TimeToGridLineRolPos(AEndTime);
  if DateOf(AEndTime) <> AEndTime then
    Inc(AStopGridPos);
end;

function TPlannerMonthViewEh.TimeToGridLineRolPos(ADateTime: TDateTime): Integer;
begin
  Result := DaysBetween(StartDate, ADateTime);
end;

procedure TPlannerMonthViewEh.InitSpanItem(
  ASpanItem: TTimeSpanDisplayItemEh);
begin
  CalcPosByPeriod(
    ASpanItem.StartTime, ASpanItem.EndTime,
    ASpanItem.FStartGridRollPos, ASpanItem.FStopGridRolPos);
  ASpanItem.FGridColNum := 0;
  ASpanItem.FAllowedInteractiveChanges := [sichSpanMovingEh];
  ASpanItem.FTimeOrientation := toHorizontalEh;
end;

procedure TPlannerMonthViewEh.SortPlanItems;
var
  i: Integer;
begin
  if FSortedSpans = nil then Exit;

  FSortedSpans.Clear;
  for i := 0 to SpanItemsCount-1 do
    FSortedSpans.Add(SpanItems[i]);

  FSortedSpans.Sort(CompareSpanItemFuncBySpan);
end;

function TPlannerMonthViewEh.GetSortedSpan(Index: Integer): TTimeSpanDisplayItemEh;
begin
  Result := TTimeSpanDisplayItemEh(FSortedSpans[Index]);
end;

procedure TPlannerMonthViewEh.ReadDivByWeekPlanItem(StartDate, BoundDate: TDateTime;
  APlanItem: TPlannerDataItemEh);
var
  SpanItem: TTimeSpanDisplayItemEh;
begin
  SpanItem := AddSpanItem(APlanItem);
  SpanItem.FPlanItem := APlanItem;
  SpanItem.FHorzLocating := brrlGridRolAreaEh;
  SpanItem.FVertLocating := brrlGridRolAreaEh;
  if APlanItem.StartTime < StartDate then
    SpanItem.StartTime := DateOf(StartDate)
  else
    SpanItem.StartTime := DateOf(APlanItem.StartTime);
  if APlanItem.EndTime > BoundDate then
    SpanItem.EndTime :=  DateOf(BoundDate)
  else if DateOf(APlanItem.EndTime) = APlanItem.EndTime then
    SpanItem.EndTime := DateOf(APlanItem.EndTime)
  else
    SpanItem.EndTime := DateOf(APlanItem.EndTime) + 1;
  InitSpanItem(SpanItem);
end;

procedure TPlannerMonthViewEh.ReadPlanItem(APlanItem: TPlannerDataItemEh);
var
  i: Integer;
  WeekBoundDate: TDateTime;
begin
  for i := 0 to 5 do
  begin
    WeekBoundDate := StartDate + i*7 + 7;
    if (APlanItem.StartTime < WeekBoundDate) and (APlanItem.EndTime > WeekBoundDate - 7) then
      ReadDivByWeekPlanItem(WeekBoundDate - 7, WeekBoundDate, APlanItem);
  end;
end;

procedure TPlannerMonthViewEh.InitSpanItemMoving(
  SpanItem: TTimeSpanDisplayItemEh; MousePos: TPoint);
var
  ACell: TGridCoord;
  ACellTime: TDateTime;
begin
  ACell := MouseCoord(MousePos.X, MousePos.Y);
  ACellTime := CellToDateTime(ACell.X, ACell.Y);
  FMovingDaysShift := DaysBetween(DateOf(SpanItem.PlanItem.StartTime), DateOf(ACellTime));
end;

procedure TPlannerMonthViewEh.UpdateDummySpanItemSize(MousePos: TPoint);
var
  ACell: TGridCoord;
  ANewTime: TDateTime;
  ATimeLen: TDateTime;
  ResViewIdx: Integer;
  AResource: TPlannerResourceEh;
begin
  if FPlannerState = psSpanRightSizingEh then
  begin
    ACell := MouseCoord(MousePos.X, MousePos.Y);
    ANewTime := CellToDateTime(ACell.X, ACell.Y);
    ANewTime := IncDay(ANewTime);
    ResViewIdx := GetResourceViewAtCell(ACell.X, ACell.Y);
    if ResViewIdx >= 0
      then AResource := FResourcesView[ResViewIdx].Resource
      else AResource := FDummyPlanItem.Resource;
    if (FDummyPlanItem.EndTime <> ANewTime) and
       (FDummyPlanItem.StartTime < ANewTime) and
       (AResource = FDummyPlanItem.Resource) then
    begin
      FDummyCheckPlanItem.Assign(FDummyPlanItem);
      FDummyCheckPlanItem.EndTime := ANewTime;
      CheckSetDummyPlanItem(FDummyPlanItemFor, FDummyCheckPlanItem);
      PlannerDataSourceChanged;
    end;
  end else if FPlannerState = psSpanLeftSizingEh then
  begin
    ACell := MouseCoord(MousePos.X, MousePos.Y);
    ANewTime := CellToDateTime(ACell.X, ACell.Y);
    ResViewIdx := GetResourceViewAtCell(ACell.X, ACell.Y);
    if ResViewIdx >= 0
      then AResource := FResourcesView[ResViewIdx].Resource
      else AResource := FDummyPlanItem.Resource;
    if (FDummyPlanItem.StartTime <> ANewTime) and
       (FDummyPlanItem.EndTime > ANewTime) and
       (AResource = FDummyPlanItem.Resource) then
    begin
      FDummyCheckPlanItem.Assign(FDummyPlanItem);
      FDummyCheckPlanItem.StartTime := ANewTime;
      CheckSetDummyPlanItem(FDummyPlanItemFor, FDummyCheckPlanItem);
      PlannerDataSourceChanged;
    end;
  end else if FPlannerState in [psSpanMovingEh, psSpanTestMovingEh] then
  begin
    ACell := MouseCoord(MousePos.X, MousePos.Y);
    ANewTime := CellToDateTime(ACell.X, ACell.Y) - FMovingDaysShift;
    ResViewIdx := GetResourceViewAtCell(ACell.X, ACell.Y);
    if ResViewIdx >= 0
      then AResource := FResourcesView[ResViewIdx].Resource
      else AResource := FDummyPlanItem.Resource;

    if (FDummyPlanItem.StartTime <> ANewTime) or
       (AResource <> FDummyPlanItem.Resource) then
    begin
      ATimeLen :=  FDummyPlanItem.EndTime - FDummyPlanItem.StartTime;
      ANewTime := DateOf(ANewTime) + TimeOf(FDummyPlanItem.StartTime);

      FDummyCheckPlanItem.Assign(FDummyPlanItem);
      FDummyCheckPlanItem.StartTime := ANewTime;
      FDummyCheckPlanItem.EndTime := FDummyCheckPlanItem.StartTime + ATimeLen;
      FDummyCheckPlanItem.Resource := AResource;
      CheckSetDummyPlanItem(FDummyPlanItemFor, FDummyCheckPlanItem);

      PlannerDataSourceChanged;
    end;
    ShowMoveHintWindow(FDummyPlanItem, MousePos);
  end;
end;

function TPlannerMonthViewEh.GetResourceViewAtCell(ACol, ARow: Integer): Integer;
begin
  if ACol < FDataColsOffset then
    Result := -1
  else if IsInterResourceCell(ACol, ARow) then
    Result := -1
  else
    Result := (ACol-FDataColsOffset) div FBarsPerRes;
end;

function TPlannerMonthViewEh.NewItemParams(var StartTime, EndTime: TDateTime;
  var Resource: TPlannerResourceEh): Boolean;
begin
  StartTime := CellToDateTime(Col, Row);
  EndTime := StartTime + EncodeTime(0, 30, 0, 0);
  Resource :=  GetResourceAtCell(Col, Row);
  Result := True;
end;

procedure TPlannerMonthViewEh.SetWeekColArea(const Value: TWeekBarAreaEh);
begin
  FWeekColArea.Assign(Value);
end;

function TPlannerMonthViewEh.GetDayNameArea: TDayNameVertAreaEh;
begin
  Result := TDayNameVertAreaEh(inherited DayNameArea);
end;

procedure TPlannerMonthViewEh.SetDayNameArea(const Value: TDayNameVertAreaEh);
begin
  inherited DayNameArea := Value;
end;

function TPlannerMonthViewEh.CreateDayNameArea: TDayNameAreaEh;
begin
  Result := TDayNameVertAreaEh.Create(Self);
end;

function TPlannerMonthViewEh.GetDataBarsArea: TDataBarsVertAreaEh;
begin
  Result := TDataBarsVertAreaEh(inherited DataBarsArea);
end;

procedure TPlannerMonthViewEh.SetDataBarsArea(const Value: TDataBarsVertAreaEh);
begin
  inherited DataBarsArea := Value;
end;

procedure TPlannerMonthViewEh.CMFontChanged(var Message: TMessage);
begin
  inherited;
  WeekColArea.RefreshDefaultFont;
end;

function TPlannerMonthViewEh.CreateResourceCaptionArea: TResourceCaptionAreaEh;
begin
  Result := TResourceVertCaptionAreaEh.Create(Self);
end;

function TPlannerMonthViewEh.CreateHoursBarArea: THoursBarAreaEh;
begin
  Result := THoursVertBarAreaEh.Create(Self);
end;

{ TAxisTimeBandPlannerViewEh }

constructor TPlannerAxisTimelineViewEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FSortedSpans := TList.Create;
  FDatesBarArea := CreateDatesBarArea;
  FBandOreintation := atboVerticalEh;
  FTimeAxis := VertAxis;
  FResourceAxis := HorzAxis;
  FRolLenInSecs := 1;
  FMasterGroupLineColor :=  ChangeRelativeColorLuminance(GridLineColors.GetDarkColor, -10);
end;

destructor TPlannerAxisTimelineViewEh.Destroy;
begin
  FreeAndNil(FSortedSpans);
  FreeAndNil(FDatesBarArea);
  inherited Destroy;
end;

procedure TPlannerAxisTimelineViewEh.SetRandeKind(const Value: TTimePlanRangeKind);
var
  AdjustedDate: TDateTime;
begin
  if FRandeKind <> Value then
  begin
    FRandeKind := Value;

    AdjustedDate := AdjustDate(CurrentTime);
    FStartDate := AdjustedDate;
    CalcLayouts;
    GridLayoutChanged;
    ResetLoadedTimeRange;
    StartDateChanged;
  end;
end;

procedure TPlannerAxisTimelineViewEh.SetBandOreintation(
  const Value: TAxisTimeBandOreintationEh);
begin
  if FBandOreintation <> Value then
  begin
    FBandOreintation := Value;
    if FBandOreintation = atboVerticalEh then
    begin
      FTimeAxis := VertAxis;
      FResourceAxis := HorzAxis;
    end else
    begin
      FTimeAxis := HorzAxis;
      FResourceAxis := VertAxis;
    end;

    CalcLayouts;
    GridLayoutChanged;
    ResetLoadedTimeRange;
    StartDateChanged;
  end;
end;

function TPlannerAxisTimelineViewEh.CreateDatesBarArea: TDatesBarAreaEh;
begin
 Result := nil;
end;

procedure TPlannerAxisTimelineViewEh.SetDatesBarArea(const Value: TDatesBarAreaEh);
begin
  FDatesBarArea.Assign(Value);
end;

//procedure TPlannerAxisTimelineViewEh.SetDatesBarAreaDaySltMode(
//  const Value: TDatesBarAreaDaySltModeEh);
//begin
//  FDatesBarAreaDaySltMode.Assign(Value);
//end;

procedure TPlannerAxisTimelineViewEh.CalcLayouts;
begin
  FBarsInBand := -1;

  case RangeKind of
    rkDayByHoursEh:
      begin
        FBarsInBand := 24;
//        FColsInTimeGroup := 24 * 2;
        FCellsInBand := FBarsInBand * 2;
        FRolLenInSecs := SecondsBetween(StartDate, AppendPeriod(StartDate, 1));
        FDataColsOffset := 1;
        FDataRowsOffset := 2;
      end;
    rkWeekByHoursEh:
      begin
        FBarsInBand := 24 * 7;
//        FColsInTimeGroup := 24 * 2;
        FCellsInBand := FBarsInBand * 2;
        FRolLenInSecs := SecondsBetween(StartDate, AppendPeriod(StartDate, 1));
        FDataColsOffset := 1;
        FDataRowsOffset := 2;
      end;
    rkWeekByDaysEh:
      begin
        FBarsInBand := 7;
//        FColsInTimeGroup := 1;
        FCellsInBand := FBarsInBand;
        FRolLenInSecs := SecondsBetween(StartDate, AppendPeriod(StartDate, 1));
        FDataColsOffset := 1;
        FDataRowsOffset := 1;
      end;
    rkMonthByDaysEh:
      begin
        FBarsInBand := DaysInMonth(StartDate);
//        FColsInTimeGroup := 1;
        FCellsInBand := FBarsInBand;
        FRolLenInSecs := SecondsBetween(StartDate, AppendPeriod(StartDate, 1));
        FDataColsOffset := 1;
        FDataRowsOffset := 1;
      end;
  end;
end;

function TPlannerAxisTimelineViewEh.NextDate: TDateTime;
begin
  Result := AppendPeriod(StartDate, 1);
end;

procedure TPlannerAxisTimelineViewEh.PlannerStateChanged(AOldState: TPlannerStateEh);
begin
  inherited PlannerStateChanged(AOldState);
  StopSpanMoveSliding;
end;

function TPlannerAxisTimelineViewEh.PriorDate: TDateTime;
begin
  Result := AppendPeriod(StartDate, -1);
end;

function TPlannerAxisTimelineViewEh.AppendPeriod(ATime: TDateTime;
  Periods: Integer): TDateTime;
begin
  case RangeKind of
    rkDayByHoursEh:
      Result := IncDay(ATime, Periods);
    rkWeekByHoursEh:
      Result := IncWeek(ATime, Periods);
    rkWeekByDaysEh:
      Result := IncWeek(ATime, Periods);
    rkMonthByDaysEh:
      Result := IncMonth(ATime, Periods);
  else
    raise Exception.Create('RangeKind is not supported (' +
      GetEnumName(TypeInfo(TTimePlanRangeKind), Ord(RangeKind)) + ')');
  end;
end;

function TPlannerAxisTimelineViewEh.GetPeriodCaption: String;
var
  StartDate, EndDate: TDateTime;
begin
  GetViewPeriod(StartDate, EndDate);
  Result := DateRangeToStr(StartDate, EndDate);
end;

function TPlannerAxisTimelineViewEh.GetResourceAtCell(ACol,
  ARow: Integer): TPlannerResourceEh;
var
  ResViewIdx: Integer;
begin
  ResViewIdx := GetResourceViewAtCell(ACol, ARow);
  if (ResViewIdx >= 0) and (PlannerDataSource <> nil) and (ResViewIdx < PlannerDataSource.Resources.Count) then
    Result := PlannerDataSource.Resources[ResViewIdx]
  else
    Result := nil;
end;

procedure TPlannerAxisTimelineViewEh.Resize;
begin
  inherited Resize;
  if HandleAllocated then
    SetDisplayPosesSpanItems;
end;

function TPlannerAxisTimelineViewEh.AdjustDate(
  const Value: TDateTime): TDateTime;
begin
  case RangeKind of
    rkDayByHoursEh:
      Result := DateOf(Value);
    rkWeekByHoursEh:
      Result := StartOfTheWeek(Value);
    rkWeekByDaysEh:
      Result := StartOfTheWeek(Value);
    rkMonthByDaysEh:
      Result := StartOfTheMonth(Value);
  else
    raise Exception.Create('RangeKind in not supported (' +
      GetEnumName(TypeInfo(TTimePlanRangeKind), Ord(RangeKind)) + ')');
  end;
end;

function TPlannerAxisTimelineViewEh.FullRedrawOnSroll: Boolean;
begin
  Result := True;
end;

procedure TPlannerAxisTimelineViewEh.BuildDaysGridData;
begin

end;

procedure TPlannerAxisTimelineViewEh.BuildHoursGridData;
begin

end;

function TPlannerAxisTimelineViewEh.DateTimeToGridRolPos(ADateTime: TDateTime): Integer;
begin
  case RangeKind of
    rkDayByHoursEh:
      Result := Round(Integer(FTimeAxis.RolLen) * SecondSpan(StartDate, ADateTime) / FRolLenInSecs);
    rkWeekByHoursEh:
      Result := Round(Integer(FTimeAxis.RolLen) * SecondSpan(StartDate, ADateTime) / FRolLenInSecs);
    rkWeekByDaysEh:
      Result := Round(Integer(FTimeAxis.RolLen) * SecondSpan(StartDate, ADateTime) / FRolLenInSecs);
    rkMonthByDaysEh:
      Result := Round(Integer(FTimeAxis.RolLen) * SecondSpan(StartDate, ADateTime) / FRolLenInSecs);
  else
    Result := 0;
  end;
end;

procedure TPlannerAxisTimelineViewEh.SetResOffsets;
begin

end;

procedure TPlannerAxisTimelineViewEh.BuildGridData;
begin
  inherited BuildGridData;
  CalcLayouts;
  if FBarsInBand <= 0 then Exit;
  ClearGridCells;
  if RangeKind in [rkDayByHoursEh, rkWeekByHoursEh] then
    BuildHoursGridData
  else
    BuildDaysGridData;
  RealignGridControls;
end;

procedure TPlannerAxisTimelineViewEh.CalcPosByPeriod(AStartTime, AEndTime: TDateTime;
  var AStartGridPos, AStopGridPos: Integer);
begin
  if AStartTime < StartDate then
    AStartTime := StartDate;
  if AEndTime > NextDate then
    AEndTime := NextDate;

  if RangeKind in [rkDayByHoursEh, rkWeekByHoursEh] then
  begin
    AStartGridPos := DateTimeToGridRolPos(AStartTime);
    AStopGridPos := DateTimeToGridRolPos(AEndTime);
  end else
  begin
    AStartGridPos := DateTimeToGridRolPos(DateOf(AStartTime));
    AStopGridPos := DateTimeToGridRolPos(DateOf(AEndTime+1));
  end;
end;

procedure TPlannerAxisTimelineViewEh.InitSpanItem(
  ASpanItem: TTimeSpanDisplayItemEh);
begin
  CalcPosByPeriod(
    ASpanItem.PlanItem.StartTime, ASpanItem.PlanItem.EndTime,
    ASpanItem.FStartGridRollPos, ASpanItem.FStopGridRolPos);
  ASpanItem.FGridColNum := 0;
end;

procedure TPlannerAxisTimelineViewEh.ReadPlanItem(APlanItem: TPlannerDataItemEh);
begin
  inherited ReadPlanItem(APlanItem);
end;

procedure TPlannerAxisTimelineViewEh.CheckDrawCellBorder(ACol, ARow: Integer;
  BorderType: TGridCellBorderTypeEh; var IsDraw: Boolean;
  var BorderColor: TColor; var IsExtent: Boolean);
begin
  inherited CheckDrawCellBorder(ACol, ARow, BorderType, IsDraw, BorderColor, IsExtent);
end;

function TPlannerAxisTimelineViewEh.GetDataBarsAreaDefaultBarSize: Integer;
begin
  if RangeKind in [rkDayByHoursEh, rkWeekByHoursEh] then
    Result := inherited GetDataBarsAreaDefaultBarSize
  else
  begin
    if HandleAllocated then
    begin
      Canvas.Font := Font;
      Result := Trunc(Canvas.TextHeight('Wg') * 10);
    end else
      Result := 16;
  end;
end;

function TPlannerAxisTimelineViewEh.GetDefaultDatesBarSize: Integer;
begin
  if HandleAllocated then
  begin
    Canvas.Font := DatesBarArea.Font;
    Result := Trunc(Canvas.TextHeight('Wg') * 1.5);
  end else
    Result := 10;
end;

function TPlannerAxisTimelineViewEh.GetDefaultDatesBarVisible: Boolean;
begin
  Result := True;
end;

function TPlannerAxisTimelineViewEh.GetGridOffsetForResource(
  Resource: TPlannerResourceEh): Integer;
var
  i: Integer;
begin
  Result := 0;
  if PlannerDataSource <> nil then
  begin
    if Resource <> nil then
    begin
      for i := 0 to PlannerDataSource.Resources.Count-1 do
        if PlannerDataSource.Resources[i] = Resource then
        begin
          Result := FResourcesView[i].GridOffset;
          Break;
        end
    end else if FShowUnassignedResource then
      Result := FResourcesView[Length(FResourcesView)-1].GridOffset;
  end;
end;

function TPlannerAxisTimelineViewEh.GetSortedSpan(Index: Integer): TTimeSpanDisplayItemEh;
begin
  Result := TTimeSpanDisplayItemEh(FSortedSpans[Index]);
end;

procedure TPlannerAxisTimelineViewEh.GetViewPeriod(var AStartDate,
  AEndDate: TDateTime);
begin
  AStartDate := StartDate;
  AEndDate := AppendPeriod(StartDate, 1);
end;

function TPlannerAxisTimelineViewEh.IsDayNameAreaNeedVisible: Boolean;
begin
  Result := False;//not IsResourceCaptionNeedVisible;
end;

function TPlannerAxisTimelineViewEh.IsInterResourceCell(ACol,
  ARow: Integer): Boolean;
begin
  Result := False;
end;

procedure TPlannerAxisTimelineViewEh.RangeModeChanged;
begin
  inherited RangeModeChanged;
  HorzScrollBar.SmoothStep := True;
end;

procedure TPlannerAxisTimelineViewEh.SetGroupPosesSpanItems(
  Resource: TPlannerResourceEh);
var
  i: Integer;
  SpanItem: TTimeSpanDisplayItemEh;
  CurStack: TList;
  CurList: TList;
  CurColbarCount: Integer;
  FullEmpty: Boolean;

  procedure CheckPushOutStack(ABoundPos: Integer);
  var
    i: Integer;
    StackSpanItem: TTimeSpanDisplayItemEh;
  begin
    for i := 0 to CurStack.Count-1 do
    begin
      StackSpanItem := TTimeSpanDisplayItemEh(CurStack[i]);
      if (StackSpanItem <> nil) and (StackSpanItem.FStopGridRolPos <= ABoundPos) then
      begin
        CurList.Add(CurStack[i]);
        CurStack[i] := nil;
      end;
    end;

    FullEmpty := True;
    for i := 0 to CurStack.Count-1 do
      if CurStack[i] <> nil then
      begin
        FullEmpty := False;
        Break;
      end;

    if FullEmpty then
    begin
      CurColbarCount := 1;
      CurList.Clear;
      CurStack.Clear;
    end;
  end;

  procedure PushInStack(ASpanItem: TTimeSpanDisplayItemEh);
  var
    i: Integer;
    StackSpanItem: TTimeSpanDisplayItemEh;
    PlaceFound: Boolean;
  begin
    PlaceFound := False;
    if CurStack.Count > 0 then
    begin
      for i := 0 to CurStack.Count-1 do
      begin
        StackSpanItem := TTimeSpanDisplayItemEh(CurStack[i]);
        if StackSpanItem = nil then
        begin
          if BandOreintation = atboVerticalEh then
          begin
            ASpanItem.FInCellCols := CurColbarCount;
            ASpanItem.FInCellFromCol := i;
            ASpanItem.FInCellToCol := i;
          end else
          begin
            ASpanItem.FInCellRows := CurColbarCount;
            ASpanItem.FInCellFromRow := i;
            ASpanItem.FInCellToRow := i;
          end;
          CurStack[i] := ASpanItem;
          PlaceFound := True;
          Break;
        end;
      end;
    end;
    if not PlaceFound then
    begin
      if CurStack.Count > 0 then
      begin
        CurColbarCount := CurColbarCount + 1;
        for i := 0 to CurStack.Count-1 do
          if BandOreintation = atboVerticalEh then
            TTimeSpanDisplayItemEh(CurStack[i]).FInCellCols := CurColbarCount
          else
            TTimeSpanDisplayItemEh(CurStack[i]).FInCellRows := CurColbarCount;
        for i := 0 to CurList.Count-1 do
          if BandOreintation = atboVerticalEh then
            TTimeSpanDisplayItemEh(CurList[i]).FInCellCols := CurColbarCount
          else
            TTimeSpanDisplayItemEh(CurList[i]).FInCellRows := CurColbarCount;
      end;
      if BandOreintation = atboVerticalEh then
      begin
        ASpanItem.FInCellCols := CurColbarCount;
        ASpanItem.FInCellFromCol := CurColbarCount-1;
        ASpanItem.FInCellToCol := CurColbarCount-1;
      end else
      begin
        ASpanItem.FInCellRows := CurColbarCount;
        ASpanItem.FInCellFromRow := CurColbarCount-1;
        ASpanItem.FInCellToRow := CurColbarCount-1;
      end;
      CurStack.Add(ASpanItem);
    end;
  end;

begin
  CurStack := TList.Create;
  CurList := TList.Create;
  CurColbarCount := 1;

  for i := 0 to SpanItemsCount-1 do
  begin
    SpanItem := SortedSpan[i];
    if SpanItem.PlanItem.Resource = Resource then
    begin
      CheckPushOutStack(SpanItem.FStartGridRollPos);
      PushInStack(SpanItem);
    end;
  end;
  CurStack.Free;
  CurList.Free;
end;

procedure TPlannerAxisTimelineViewEh.SortPlanItems;
var
  i: Integer;
begin
  if FSortedSpans = nil then Exit;

  FSortedSpans.Clear;
  for i := 0 to SpanItemsCount-1 do
    FSortedSpans.Add(SpanItems[i]);

  FSortedSpans.Sort(CompareSpanItemFuncByPlan);
end;

procedure TPlannerAxisTimelineViewEh.StartDateChanged;
begin
  if RangeKind in [rkMonthByDaysEh] then
    GridLayoutChanged;
  inherited StartDateChanged;
end;

function TPlannerAxisTimelineViewEh.NewItemParams(var StartTime,
  EndTime: TDateTime; var Resource: TPlannerResourceEh): Boolean;
begin
  StartTime := CellToDateTime(Col, Row);
  EndTime := StartTime + EncodeTime(0, 30, 0, 0);
  Resource :=  GetResourceAtCell(Col, Row);
  Result := True;
end;

function TPlannerAxisTimelineViewEh.GetCoveragePeriodType: TPlannerCoveragePeriodTypeEh;
begin
  case RangeKind of
    rkDayByHoursEh: Result := pcpDayEh;
    rkWeekByHoursEh: Result := pcpWeekEh;
    rkWeekByDaysEh: Result := pcpWeekEh;
    rkMonthByDaysEh: Result := pcpMonthEh;
  else
    Result := pcpDayEh;
  end;
end;

procedure TPlannerAxisTimelineViewEh.CoveragePeriod(var AFromTime,
  AToTime: TDateTime);
begin
  AFromTime := StartDate;
  case CoveragePeriodType of
    pcpDayEh: AToTime := StartDate + 1;
    pcpWeekEh: AToTime := IncWeek(StartDate);
    pcpMonthEh: AToTime := IncMonth(StartDate, 1);
  end;
end;

{ TPlannerVertTimelineViewEh }

constructor TPlannerVertTimelineViewEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FMinDataColWidth := -1;
end;

destructor TPlannerVertTimelineViewEh.Destroy;
begin
  inherited Destroy;
end;

procedure TPlannerVertTimelineViewEh.DrawCell(ACol, ARow: Integer; ARect: TRect;
  State: TGridDrawState);
begin
  inherited DrawCell(ACol, ARow, ARect, State);
end;

procedure TPlannerVertTimelineViewEh.GetCellType(ACol, ARow: Integer;
  var CellType: TPlannerViewCellTypeEh; var ALocalCol, ALocalRow: Integer);
begin
  if (ARow >= FixedRowCount) and (ACol = FDayGroupCol) then
  begin
    CellType := pctDateBarEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end else if (ARow >= FixedRowCount) and (ACol = FHoursBarIndex) then
  begin
    CellType := pctTimeCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end else if (ARow >= FixedRowCount) and (ACol = FDaySplitModeCol) then
  begin
    CellType := pctDateCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end else if (ARow = FDayNameBarPos) and (ACol >= FixedColCount) then
  begin
    CellType := pctDayNameCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end else if (ARow = 0) and (ACol < FixedRowCount) then
  begin
    CellType := pctTopLeftCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end else if (ARow = FResourceAxisPos) and (ACol >= FixedColCount) then
  begin
     if IsInterResourceCell(ACol, ARow)
      then CellType := pctInterResourceSpaceEh
      else CellType := pctResourceCaptionCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end else
  begin
    CellType := pctDataCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end;
end;

procedure TPlannerVertTimelineViewEh.DrawDataCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
var
  NextCellDateTime, CellDateTime: TDateTime;
begin

  if RangeKind in [rkDayByHoursEh, rkWeekByHoursEh] then
  begin
    CellDateTime := CellToDateTime(ACol, ARow);
    if ARow < RowCount-1
      then NextCellDateTime := CellToDateTime(ACol, ARow+1)
      else NextCellDateTime := CellDateTime;

    if DateOf(CellDateTime) <> DateOf(NextCellDateTime) then
    begin

      Canvas.Pen.Color := FMasterGroupLineColor;
      DrawPolyline([Point(ARect.Left, ARect.Bottom-1), Point(ARect.Right, ARect.Bottom-1)]);
      DrawPolyline([Point(ARect.Left, ARect.Bottom-2), Point(ARect.Right, ARect.Bottom-2)]);

      ARect.Bottom := ARect.Bottom - 2;
    end;
  end;

  inherited DrawDataCell(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
end;

procedure TPlannerVertTimelineViewEh.DrawDateBar(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
begin
  DrawTimeGroupCell(ACol, ARow, ARect, State, DrawArgs);
end;

procedure TPlannerVertTimelineViewEh.DrawDateCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
begin
  DrawDaySplitModeDateCell(ACol, ARow, ARect, State, ARow-FixedRowCount, DrawArgs);
end;

procedure TPlannerVertTimelineViewEh.DrawTimeCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
var
  NextCellDateTime, CellDateTime: TDateTime;
  s: String;
  w: Integer;
  wn: String;
begin
  if RangeKind in [rkDayByHoursEh, rkWeekByHoursEh] then
  begin
    CellDateTime := CellToDateTime(ACol, ARow);
    if ARow < RowCount-2
      then NextCellDateTime := CellToDateTime(ACol, ARow+2)
      else NextCellDateTime := CellDateTime;

    if DateOf(CellDateTime) <> DateOf(NextCellDateTime) then
    begin

      Canvas.Pen.Color := FMasterGroupLineColor;
      DrawPolyline([Point(ARect.Left, ARect.Bottom-1), Point(ARect.Right, ARect.Bottom-1)]);
      DrawPolyline([Point(ARect.Left, ARect.Bottom-2), Point(ARect.Right, ARect.Bottom-2)]);

      ARect.Bottom := ARect.Bottom - 2;
    end;

    inherited DrawTimeCell(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
  end else
  begin
    CellDateTime := CellToDateTime(ACol, ARow);
    Canvas.Font := Font;
    Canvas.Brush.Color := StyleServices.GetSystemColor(Color);
    wn := FormatSettings.ShortDayNames[DayOfWeek(CellDateTime)];
    s := wn + ', ' + FormatDateTime('D', CellDateTime);
    w := Canvas.TextWidth('0') div 2;
    WriteTextEh(Canvas, ARect, True, w, 2, s,
      taLeftJustify, tlTop, False, False, 0, 0, False, True);
  end;
end;

procedure TPlannerVertTimelineViewEh.DrawTimeGroupCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; DrawArgs: TPlannerViewCellDrawArgsEh);
var
  s: String;
begin
  Canvas.Pen.Color := StyleServices.GetSystemColor(FMasterGroupLineColor);
  DrawPolyline([Point(ARect.Left, ARect.Bottom-1), Point(ARect.Right, ARect.Bottom-1)]);
  DrawPolyline([Point(ARect.Left, ARect.Bottom-2), Point(ARect.Right, ARect.Bottom-2)]);

  ARect.Bottom := ARect.Bottom - 2;
  s := DrawArgs.Text;

  Canvas.Font.Name := DrawArgs.FontName;
  Canvas.Font.Size := DrawArgs.FontSize;
  Canvas.Font.Style := DrawArgs.FontStyle;
  Canvas.Font.Color := DrawArgs.FontColor;
  Canvas.Brush.Color := DrawArgs.BackColor;

  Canvas.FillRect(ARect);
  if ARect.Top < VertAxis.FixedBoundary then
    ARect.Top := VertAxis.FixedBoundary;
  if ARect.Bottom > VertAxis.ContraStart then
    ARect.Bottom := VertAxis.ContraStart;

  WriteTextVerticalEh(Canvas, ARect, True, 0, 0, s, taCenter, tlCenter, False, False);
end;

procedure TPlannerVertTimelineViewEh.GetDateBarDrawParams(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
var
  DrawDate: TDateTime;
begin
  DrawDate := CellToDateTime(ACol, ARow);
  DrawArgs.Text := FormatDateTime(FormatSettings.LongDateFormat, DrawDate) + ' ';
  DrawArgs.Value := DrawDate;

  DrawArgs.FontName := DatesBarArea.Font.Name;
  DrawArgs.FontColor := DatesBarArea.Font.Color;
  DrawArgs.FontSize := DatesBarArea.Font.Size;
  DrawArgs.FontStyle := DatesBarArea.Font.Style;
  DrawArgs.BackColor := DatesBarArea.GetActualColor;
  DrawArgs.Alignment := taCenter;
  DrawArgs.Layout := tlCenter;
  DrawArgs.Orientation := tohVertical;
end;

procedure TPlannerVertTimelineViewEh.DrawDaySplitModeDateCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ADataRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
begin

end;

procedure TPlannerVertTimelineViewEh.CalcLayouts;
begin
  inherited CalcLayouts;
  FDayGroupRows := -1;

  case RangeKind of
    rkDayByHoursEh:
      begin
        FDayGroupRows := 24 * 2;
      end;
    rkWeekByHoursEh:
      begin
        FDayGroupRows := 24 * 2;
      end;
    rkWeekByDaysEh:
      begin
        FDayGroupRows := 1;
      end;
    rkMonthByDaysEh:
      begin
        FDayGroupRows := 1;
      end;
  end;
end;

procedure TPlannerVertTimelineViewEh.CalcPosByPeriod(AStartTime,
  AEndTime: TDateTime; var AStartGridPos, AStopGridPos: Integer);
begin
  if AStartTime < StartDate then
    AStartTime := StartDate;
  if AEndTime > NextDate then
    AEndTime := NextDate;

  AStartGridPos := DateTimeToGridRolPos(AStartTime);
  AStopGridPos := DateTimeToGridRolPos(AEndTime);
end;

procedure TPlannerVertTimelineViewEh.InitSpanItem(
  ASpanItem: TTimeSpanDisplayItemEh);
var
  AStartTime, AEndTime: TDateTime;
begin
  if RangeKind in [rkDayByHoursEh, rkWeekByHoursEh] then
  begin
    CalcPosByPeriod(
      ASpanItem.PlanItem.StartTime, ASpanItem.PlanItem.EndTime,
      ASpanItem.FStartGridRollPos, ASpanItem.FStopGridRolPos);
  end else
  begin
    AStartTime := DateOf(ASpanItem.PlanItem.StartTime);
    AEndTime := DateOf(ASpanItem.PlanItem.EndTime);
    if AEndTime <> DayOf(ASpanItem.PlanItem.EndTime) then
      AEndTime := IncDay(AEndTime);
    CalcPosByPeriod(AStartTime, AEndTime,
      ASpanItem.FStartGridRollPos, ASpanItem.FStopGridRolPos);
    ASpanItem.StartTime := AStartTime;
    ASpanItem.EndTime := AEndTime;
  end;
  ASpanItem.FGridColNum := 0;
  ASpanItem.FAllowedInteractiveChanges :=
    [sichSpanMovingEh, sichSpanTopSizingEh, sichSpanButtomSizingEh];
  ASpanItem.FTimeOrientation := toVerticalEh;
end;

procedure TPlannerVertTimelineViewEh.SetGroupPosesSpanItems(Resource: TPlannerResourceEh);
begin
  if RangeKind in [rkDayByHoursEh, rkWeekByHoursEh] then
    inherited SetGroupPosesSpanItems(Resource)
  else
    SetGroupPosesSpanItemsForDayStep(Resource);
end;

procedure TPlannerVertTimelineViewEh.SetGroupPosesSpanItemsForDayStep(
  Resource: TPlannerResourceEh);
var
  i: Integer;
  SpanItem: TTimeSpanDisplayItemEh;
  CurStack: TList;
  CurList: TList;
  CurColbarCount: Integer;
  FullEmpty: Boolean;
  InDayPos: Integer;
  InDayList: TList;

  procedure ClearInDay();
  var
    i: Integer;
    InDaySpanItem: TTimeSpanDisplayItemEh;
  begin
    for i := 0 to InDayList.Count-1 do
    begin
      InDaySpanItem := TTimeSpanDisplayItemEh(InDayList[i]);
      InDaySpanItem.FInCellRows := InDayList.Count;
      InDaySpanItem.FInCellFromRow := i;
      InDaySpanItem.FInCellToRow := i;
    end;
    InDayList.Clear;
  end;

  procedure AddInDay(ASpanItem: TTimeSpanDisplayItemEh);
  var
    InStackDayItem: TTimeSpanDisplayItemEh;
  begin
    InStackDayItem := TTimeSpanDisplayItemEh(CurStack[InDayPos]);
    ASpanItem.FInCellCols := InStackDayItem.InCellCols;
    ASpanItem.FInCellFromCol := InDayPos;
    ASpanItem.FInCellToCol := InDayPos;
    InDayList.Add(ASpanItem);
  end;

  procedure CheckPushOutStack(ABoundPos: Integer);
  var
    i: Integer;
    StackSpanItem: TTimeSpanDisplayItemEh;
  begin
    for i := 0 to CurStack.Count-1 do
    begin
      StackSpanItem := TTimeSpanDisplayItemEh(CurStack[i]);
      if (StackSpanItem <> nil) and (StackSpanItem.FStopGridRolPos <= ABoundPos) then
      begin
        CurList.Add(CurStack[i]);
        CurStack[i] := nil;
        if i = InDayPos then
        begin
          ClearInDay();
          InDayPos := -1;
        end;
      end;
    end;

    FullEmpty := True;
    for i := 0 to CurStack.Count-1 do
      if CurStack[i] <> nil then
      begin
        FullEmpty := False;
        Break;
      end;

    if FullEmpty then
    begin
      CurColbarCount := 1;
      CurList.Clear;
      CurStack.Clear;
    end;
  end;

  procedure PushInStack(ASpanItem: TTimeSpanDisplayItemEh);
  var
    i: Integer;
    StackSpanItem: TTimeSpanDisplayItemEh;
    PlaceFound: Boolean;
  begin
    PlaceFound := False;
    if CurStack.Count > 0 then
    begin
      if ASpanItem.PlanItem.InsideDayRange and (InDayPos >= 0) then
      begin
        AddInDay(ASpanItem);
        PlaceFound := True;
      end else
      begin
        for i := 0 to CurStack.Count-1 do
        begin
          StackSpanItem := TTimeSpanDisplayItemEh(CurStack[i]);
          if StackSpanItem = nil then
          begin
            if BandOreintation = atboVerticalEh then
            begin
              ASpanItem.FInCellCols := CurColbarCount;
              ASpanItem.FInCellFromCol := i;
              ASpanItem.FInCellToCol := i;
            end else
            begin
              ASpanItem.FInCellRows := CurColbarCount;
              ASpanItem.FInCellFromRow := i;
              ASpanItem.FInCellToRow := i;
            end;
            CurStack[i] := ASpanItem;
            PlaceFound := True;
            if ASpanItem.PlanItem.InsideDayRange then
            begin
              InDayPos := i;
              InDayList.Add(ASpanItem);
            end;
            Break;
          end;
        end;
      end;
    end;
    if not PlaceFound then
    begin
      if CurStack.Count > 0 then
      begin
        CurColbarCount := CurColbarCount + 1;
        for i := 0 to CurStack.Count-1 do
          if BandOreintation = atboVerticalEh then
            TTimeSpanDisplayItemEh(CurStack[i]).FInCellCols := CurColbarCount
          else
            TTimeSpanDisplayItemEh(CurStack[i]).FInCellRows := CurColbarCount;
        for i := 0 to CurList.Count-1 do
          if BandOreintation = atboVerticalEh then
            TTimeSpanDisplayItemEh(CurList[i]).FInCellCols := CurColbarCount
          else
            TTimeSpanDisplayItemEh(CurList[i]).FInCellRows := CurColbarCount;
        if InDayPos >= 0 then
          for i := 0 to InDayList.Count-1 do
            TTimeSpanDisplayItemEh(InDayList[i]).FInCellCols := CurColbarCount
      end;
      if BandOreintation = atboVerticalEh then
      begin
        ASpanItem.FInCellCols := CurColbarCount;
        ASpanItem.FInCellFromCol := CurColbarCount-1;
        ASpanItem.FInCellToCol := CurColbarCount-1;
      end else
      begin
        ASpanItem.FInCellRows := CurColbarCount;
        ASpanItem.FInCellFromRow := CurColbarCount-1;
        ASpanItem.FInCellToRow := CurColbarCount-1;
      end;
      CurStack.Add(ASpanItem);
      if ASpanItem.PlanItem.InsideDayRange then
      begin
        InDayPos := CurStack.Count-1;
        InDayList.Add(ASpanItem);
      end;
    end;
  end;

begin
  CurStack := TList.Create;
  CurList := TList.Create;
  InDayList := TList.Create;
  CurColbarCount := 1;
  InDayPos := -1;

  for i := 0 to SpanItemsCount-1 do
  begin
    SpanItem := SortedSpan[i];
    if SpanItem.PlanItem.Resource = Resource then
    begin
      CheckPushOutStack(SpanItem.FStartGridRollPos);
      PushInStack(SpanItem);
    end;
  end;
  ClearInDay();
  CurStack.Free;
  CurList.Free;
  InDayList.Free;
end;

procedure TPlannerVertTimelineViewEh.SetResOffsets;
var
  i: Integer;
begin
  for i := 0 to Length(FResourcesView)-1 do
    if i*FBarsPerRes < HorzAxis.RolCelCount then
    begin
      if (PlannerDataSource <> nil) and (i < PlannerDataSource.Resources.Count)
        then FResourcesView[i].Resource := PlannerDataSource.Resources[i]
        else FResourcesView[i].Resource := nil;
      FResourcesView[i].GridOffset := HorzAxis.RolLocCelPosArr[i*FBarsPerRes];
      FResourcesView[i].GridStartAxisBar := i*FBarsPerRes;
    end;
end;

procedure TPlannerVertTimelineViewEh.SetDisplayPosesSpanItems;
var
  ASpanRect: TRect;
  SpanItem: TTimeSpanDisplayItemEh;
  i: Integer;
  StartX: Integer;
  ResourceOffset: Integer;
  AStartViewDate, AEndViewDate: TDateTime;
begin
  SetResOffsets;
  StartX := 0;{HorzAxis.FixedBoundary;}
  GetViewPeriod(AStartViewDate, AEndViewDate);

  for i := 0 to SpanItemsCount-1 do
  begin
    SpanItem := SpanItems[i];
//    if SpanItem.GridColNum = FDataColsOffset then
    begin
      ASpanRect.Left := StartX + 10;
      ASpanRect.Right := ASpanRect.Left + ColWidths[FDataColsOffset] - 20;
      ASpanRect.Top := {VertAxis.FixedBoundary + }SpanItem.StartGridRollPos;
      ASpanRect.Bottom := {VertAxis.FixedBoundary + }SpanItem.StopGridRolPosl;
      ResourceOffset := GetGridOffsetForResource(SpanItem.PlanItem.Resource);
      OffsetRect(ASpanRect, ResourceOffset, 0);

      CalcRectForInCellCols(SpanItem, ASpanRect);
      CalcRectForInCellRows(SpanItem, ASpanRect);

      SpanItem.FBoundRect := ASpanRect;

      if (SpanItem.PlanItem.StartTime < AStartViewDate) then
        SpanItem.FDrawBackOutInfo := True;
      if (SpanItem.PlanItem.EndTime > AEndViewDate) then
        SpanItem.FDrawForwardOutInfo := True;
    end;
  end;

end;

function TPlannerVertTimelineViewEh.CellToDateTime(ACol, ARow: Integer): TDateTime;
begin
  case RangeKind of
    rkDayByHoursEh:
      Result := IncMinute(StartDate, 30 * (ARow - FixedRowCount));
    rkWeekByHoursEh:
      Result := IncMinute(StartDate, 30 * (ARow - FixedRowCount));
    rkWeekByDaysEh:
      Result := IncDay(StartDate, (ARow - FixedRowCount));
    rkMonthByDaysEh:
      Result := IncDay(StartDate, (ARow - FixedRowCount));
  else
    Result := 0;
  end;
end;

procedure TPlannerVertTimelineViewEh.CheckDrawCellBorder(ACol, ARow: Integer;
  BorderType: TGridCellBorderTypeEh; var IsDraw: Boolean;
  var BorderColor: TColor; var IsExtent: Boolean);
var
  DataCell: TGridCoord;
  DataRow: Integer;
  NextCellDateTime, CellDateTime: TDateTime;
  NextCellOffset: Integer;
  Resource: TPlannerResourceEh;
  ResNo: Integer;
begin
  inherited CheckDrawCellBorder(ACol, ARow, BorderType, IsDraw, BorderColor, IsExtent);

  if ((ACol-FDataColsOffset >= 0) or ((ACol-FDataColsOffset = -1 ) and (BorderType = cbtRightEh))) and
     (PlannerDataSource <> nil) and
     (PlannerDataSource.Resources.Count > 0) then
  begin
    Resource := nil;
    if IsInterResourceCell(ACol, ARow) then
    begin
      if BorderType in [cbtTopEh, cbtBottomEh]  then
        IsDraw := False;
      Resource := PlannerDataSource.Resources[(ACol-FDataColsOffset) div FBarsPerRes];
    end else
    begin
      ResNo := (ACol-FDataColsOffset) div FBarsPerRes;
      if (ResNo >= 0) and (ResNo < PlannerDataSource.Resources.Count) then
        Resource := PlannerDataSource.Resources[ResNo];
    end;
    if (Resource <> nil) and (Resource.DarkLineColor <> clDefault) then
        BorderColor := Resource.DarkLineColor;
  end;

  if RangeKind in [rkDayByHoursEh, rkWeekByHoursEh] then
  begin
    DataCell := GridCoordToDataCoord(GridCoord(ACol, ARow));

    CellDateTime := CellToDateTime(ACol, ARow);
    if (ACol = 1)
      then NextCellOffset := 2
      else NextCellOffset := 1;

    if ARow < RowCount-NextCellOffset
      then NextCellDateTime := CellToDateTime(ACol, ARow+NextCellOffset)
      else NextCellDateTime := CellDateTime;

    if ARow >= FixedRowCount then
    begin
      if (BorderType = cbtBottomEh) and
         ( (ACol = FDayGroupCol) or
           ( (DateOf(CellDateTime) <> DateOf(NextCellDateTime)) )
         )  then
    //    BorderColor := cl3DDkShadow;
        IsDraw := False
      else if IsDraw and (DataCell.X >= 0) and (DataCell.Y >= FixedRowCount) then
      begin
        DataRow := DataCell.Y - FixedRowCount;
        if (BorderType in [cbtTopEh, cbtBottomEh]) and (DataRow mod 2 = 0) then
          BorderColor := ApproximateColor(BorderColor, Color, 255 div 2);
      end;
    end;
  end;
end;

procedure TPlannerVertTimelineViewEh.BuildHoursGridData;
var
  i, g, ir: Integer;
  ColGroups: Integer;
  ARowCount: Integer;
  AFixedRowCount, AFixedColCount: Integer;
  AHoursBarWidth, ADayGroupColWidth: Integer;
  ColWidth, FitGap: Integer;
  DataColsWidth: Integer;
  ATimeGroups: Integer;
  InGroupStartRow: Integer;
  BarsInGroup: Integer;
begin

  ClearGridCells;


  AFixedRowCount := TopGridLineCount;

  if ResourceCaptionArea.Visible then
  begin
    ColGroups := ResourcesCount;
    if FShowUnassignedResource then
      Inc(ColGroups);
    if ColGroups = 0 then
      ColGroups := 1;

    FResourceAxisPos := AFixedRowCount;
    AFixedRowCount := AFixedRowCount + 1;
    ARowCount := FCellsInBand + AFixedRowCount;
  end else
  begin
    ColGroups := ResourcesCount;
    if FShowUnassignedResource then
      Inc(ColGroups);
    if ColGroups = 0 then
      ColGroups := 1;
//    AFixedRowCount := 0;
    ARowCount := FCellsInBand + AFixedRowCount;
    FResourceAxisPos := -1;
  end;

  if DayNameArea.Visible then
  begin
    FDayNameBarPos := AFixedRowCount;
    Inc(ARowCount);
    Inc(AFixedRowCount);
  end else
    FDayNameBarPos := -1;


  if DatesBarArea.Visible then
  begin
    AFixedColCount := 1;
    FDataColsOffset := 1;
    FDayGroupCol := 0;
  end else
  begin
    AFixedColCount := 0;
    FDataColsOffset := 0;
    FDayGroupCol := -1;
  end;

  if HoursBarArea.Visible then
  begin
    FHoursBarIndex := AFixedColCount;
    Inc(AFixedColCount);
    Inc(FDataColsOffset);
  end else
  begin
    FHoursBarIndex := -1;
  end;

  FixedRowCount := AFixedRowCount;
  FixedColCount := AFixedColCount;
  FDataRowsOffset := AFixedRowCount;
  FDaySplitModeCol := -1;

  ColCount := AFixedColCount + ColGroups + (ColGroups-1);
  RowCount := ARowCount;
  SetGridSize(FullColCount, FullRowCount);
  if ColGroups = 1
    then FBarsPerRes := 1
    else FBarsPerRes := 2;


  if FHoursBarIndex >= 0 then
  begin
    ColWidths[FHoursBarIndex] := HoursBarArea.GetActualSize;
    AHoursBarWidth := HoursBarArea.GetActualSize;
  end else
    AHoursBarWidth := 0;


  if FDayGroupCol >= 0 then
  begin
    ColWidths[FDayGroupCol] := DatesBarArea.GetActualSize;
    ADayGroupColWidth := DatesBarArea.GetActualSize;
  end else
    ADayGroupColWidth := 0;


  DataColsWidth := (GridClientWidth - AHoursBarWidth - ADayGroupColWidth - (ColGroups-1)*3);
  ColWidth := DataColsWidth div ColGroups;
  if ColWidth < MinDataColWidth then
  begin
    ColWidth := MinDataColWidth;
    HorzScrollBar.VisibleMode := sbAutoShowEh;
    FitGap := 0;
  end else
  begin
    HorzScrollBar.VisibleMode := sbNeverShowEh;
    FitGap := DataColsWidth mod ColGroups;
  end;

  for i := FixedColCount to ColCount-1 do
  begin
    if IsInterResourceCell(i, 0) then
    begin
      ColWidths[i] := 3;
    end else
    begin
      if FitGap > 0
        then ColWidths[i] := ColWidth + 1
        else ColWidths[i] := ColWidth;
      Dec(FitGap);
    end;
  end;


  if HandleAllocated then
  begin
    DefaultRowHeight := DataBarsArea.GetActualRowHeight;
    for i := 0 to RowCount-1 do
      RowHeights[i] := DefaultRowHeight;

    if TopGridLineCount > 0 then
      RowHeights[FTopGridLineIndex] := 1;

    if FResourceAxisPos >= 0 then
      RowHeights[FResourceAxisPos] := ResourceCaptionArea.GetActualSize;

    if FDayNameBarPos >= 0 then
      RowHeights[FDayNameBarPos] := DayNameArea.GetActualSize;
  end;


  MergeCells(0,0, FixedColCount-1,FixedRowCount-1);

  ATimeGroups := FCellsInBand div FDayGroupRows;
  for g := 0 to ATimeGroups - 1 do
  begin
    InGroupStartRow := g * FDayGroupRows + FixedRowCount;
    BarsInGroup := FDayGroupRows div 2;
    if FDayGroupCol >= 0 then
      MergeCells(FDayGroupCol, InGroupStartRow, 0, FDayGroupRows-1);

    if FHoursBarIndex >= 0 then
      for i := 0 to BarsInGroup - 1 do
      begin
        ir := InGroupStartRow + i*2;
        MergeCells(FHoursBarIndex, ir, 0, 2-1);
        Cell[FHoursBarIndex, ir].Value := FormatFloat('00', i) + ':00';
      end;
  end;
end;

function TPlannerVertTimelineViewEh.GetSortedSpan(Index: Integer): TTimeSpanDisplayItemEh;
begin
  Result := TTimeSpanDisplayItemEh(FSortedSpans[Index]);
end;

procedure TPlannerVertTimelineViewEh.GetViewPeriod(var AStartDate,
  AEndDate: TDateTime);
begin
  AStartDate := StartDate;
  AEndDate := AppendPeriod(StartDate, 1);
end;

function TPlannerVertTimelineViewEh.IsInterResourceCell(ACol,
  ARow: Integer): Boolean;
begin
  if (ACol-FDataColsOffset >= 0) and
     (Length(FResourcesView) > 1)
  then
    Result := (ACol-FDataColsOffset) mod FBarsPerRes = FBarsPerRes - 1
  else
    Result := False;
end;

procedure TPlannerVertTimelineViewEh.SetMinDataColWidth(const Value: Integer);
begin
  FMinDataColWidth := Value;
end;

procedure TPlannerVertTimelineViewEh.InitSpanItemMoving(
  SpanItem: TTimeSpanDisplayItemEh; MousePos: TPoint);
var
  ACell: TGridCoord;
  ACellTime: TDateTime;
begin
  ACell := MouseCoord(MousePos.X, MousePos.Y);
  ACellTime := CellToDateTime(ACell.X, ACell.Y);
  FMovingDaysShift := DaysBetween(DateOf(SpanItem.PlanItem.StartTime), DateOf(ACellTime));
  FMovingTimeShift := ACellTime - SpanItem.PlanItem.StartTime;
end;

procedure TPlannerVertTimelineViewEh.UpdateDummySpanItemSize(MousePos: TPoint);
var
  FixedMousePos: TPoint;
  ACell: TGridCoord;
  ANewTime: TDateTime;
  ATimeLen: TDateTime;
  ResViewIdx: Integer;
  AResource: TPlannerResourceEh;
begin
  FixedMousePos := MousePos;
  if FPlannerState in [psSpanMovingEh, psSpanTestMovingEh] then
  begin
    FixedMousePos.X := FixedMousePos.X + FTopLeftSpanShift.cx;
    FixedMousePos.Y := FixedMousePos.Y + FTopLeftSpanShift.cy;
  end;
  if FPlannerState = psSpanButtomSizingEh then
  begin
    ACell := CalcCoordFromPoint(FixedMousePos.X, FixedMousePos.Y+DefaultRowHeight div 2);
    ANewTime := CellToDateTime(ACell.X, ACell.Y);
    if (FDummyPlanItem.StartTime < ANewTime) and (FDummyPlanItem.EndTime <> ANewTime) then
    begin
      FDummyCheckPlanItem.Assign(FDummyPlanItem);
      FDummyCheckPlanItem.EndTime := ANewTime;
      CheckSetDummyPlanItem(FDummyPlanItemFor, FDummyCheckPlanItem);
      PlannerDataSourceChanged();
    end;
  end else if FPlannerState = psSpanTopSizingEh then
  begin
    ACell := MouseCoord(FixedMousePos.X, FixedMousePos.Y+DefaultRowHeight div 2);
    ANewTime := CellToDateTime(ACell.X, ACell.Y);
    if (FDummyPlanItem.EndTime > ANewTime) and (FDummyPlanItem.StartTime <> ANewTime) then
    begin
      FDummyCheckPlanItem.Assign(FDummyPlanItem);
      FDummyCheckPlanItem.StartTime := ANewTime;
      CheckSetDummyPlanItem(FDummyPlanItemFor, FDummyCheckPlanItem);
      PlannerDataSourceChanged();
    end;
  end else if FPlannerState in [psSpanMovingEh, psSpanTestMovingEh] then
  begin
    ACell := CalcCoordFromPoint(FixedMousePos.X, FixedMousePos.Y);
    ANewTime := CellToDateTime(ACell.X, ACell.Y);
    ResViewIdx := GetResourceViewAtCell(ACell.X, ACell.Y);
    if ResViewIdx >= 0
      then AResource := FResourcesView[ResViewIdx].Resource
      else AResource := FDummyPlanItem.Resource;

    if (FDummyPlanItem.StartTime <> ANewTime) or
       (AResource <> FDummyPlanItem.Resource) then
    begin
      ATimeLen :=  FDummyPlanItem.EndTime - FDummyPlanItem.StartTime;

      FDummyCheckPlanItem.Assign(FDummyPlanItem);
      FDummyCheckPlanItem.StartTime := ANewTime - FMovingTimeShift;
      FDummyCheckPlanItem.EndTime := FDummyCheckPlanItem.StartTime + ATimeLen;
      FDummyCheckPlanItem.Resource := AResource;
      CheckSetDummyPlanItem(FDummyPlanItemFor, FDummyCheckPlanItem);

      PlannerDataSourceChanged;
    end;
    ShowMoveHintWindow(FDummyPlanItem, MousePos);
  end;
end;

function TPlannerVertTimelineViewEh.GetResourceViewAtCell(ACol,
  ARow: Integer): Integer;
begin
  if ACol < FDataColsOffset then
    Result := -1
  else if IsInterResourceCell(ACol, ARow) then
    Result := -1
  else
    Result := (ACol-FDataColsOffset) div FBarsPerRes;
end;

procedure TPlannerVertTimelineViewEh.MouseMove(Shift: TShiftState; X, Y: Integer);
begin
  inherited MouseMove(Shift, X, Y);
end;

function TPlannerVertTimelineViewEh.DefaultHoursBarSize: Integer;
begin
  if HandleAllocated then
  begin
    Canvas.Font := HoursBarArea.Font;
    Result := Canvas.TextWidth('11111');
  end else
    Result := 45;
end;

function TPlannerVertTimelineViewEh.GetDayNameAreaDefaultSize: Integer;
var
  h: Integer;
begin
  if HandleAllocated then
  begin
    Canvas.Font := DayNameArea.Font;
    h := Canvas.TextHeight('Wg') + 4;
    if (h > DefaultRowHeight) or (RangeKind in [rkWeekByDaysEh, rkMonthByDaysEh])
      then Result := h
      else Result := DefaultRowHeight;
  end else
    Result := DefaultRowHeight;
end;

function TPlannerVertTimelineViewEh.GetResourceCaptionAreaDefaultSize: Integer;
var
  h: Integer;
begin
  if HandleAllocated then
  begin
    Canvas.Font := ResourceCaptionArea.Font;
    h := Canvas.TextHeight('Wg') + 4;
    if (h > DefaultRowHeight) or (RangeKind in [rkWeekByDaysEh, rkMonthByDaysEh])
      then Result := h
      else Result := DefaultRowHeight;
  end else
    Result := DefaultRowHeight;
end;

function TPlannerVertTimelineViewEh.CreateDayNameArea: TDayNameAreaEh;
begin
  Result := TDayNameVertAreaEh.Create(Self);
end;

function TPlannerVertTimelineViewEh.GetDayNameArea: TDayNameVertAreaEh;
begin
  Result := TDayNameVertAreaEh(inherited DayNameArea);
end;

procedure TPlannerVertTimelineViewEh.SetDayNameArea(const Value: TDayNameVertAreaEh);
begin
  inherited DayNameArea := Value;
end;

function TPlannerVertTimelineViewEh.GetDataBarsArea: TDataBarsVertAreaEh;
begin
  Result := TDataBarsVertAreaEh(inherited DataBarsArea);
end;

procedure TPlannerVertTimelineViewEh.SetDataBarsArea(
  const Value: TDataBarsVertAreaEh);
begin
  inherited DataBarsArea := Value;
end;

function TPlannerVertTimelineViewEh.GetDatesColArea: TDatesColAreaEh;
begin
  Result := TDatesColAreaEh(DatesBarArea);
end;

procedure TPlannerVertTimelineViewEh.SetDatesColArea(const Value: TDatesColAreaEh);
begin
  DatesBarArea := Value;
end;

function TPlannerVertTimelineViewEh.CreateResourceCaptionArea: TResourceCaptionAreaEh;
begin
  Result := TResourceVertCaptionAreaEh.Create(Self);
end;

function TPlannerVertTimelineViewEh.GetResourceCaptionArea: TResourceVertCaptionAreaEh;
begin
  Result := TResourceVertCaptionAreaEh(inherited ResourceCaptionArea);
end;

procedure TPlannerVertTimelineViewEh.SetResourceCaptionArea(const Value: TResourceVertCaptionAreaEh);
begin
  inherited SetResourceCaptionArea(Value);
end;

function TPlannerVertTimelineViewEh.CreateHoursBarArea: THoursBarAreaEh;
begin
  Result := THoursVertBarAreaEh.Create(Self);
end;

function TPlannerVertTimelineViewEh.GetHoursColArea: THoursVertBarAreaEh;
begin
  Result := THoursVertBarAreaEh(inherited HoursBarArea);
end;

procedure TPlannerVertTimelineViewEh.SetHoursColArea(
  const Value: THoursVertBarAreaEh);
begin
  inherited HoursBarArea := Value;
end;

function TPlannerVertTimelineViewEh.CreateDatesBarArea: TDatesBarAreaEh;
begin
  Result := TDatesColAreaEh.Create(Self);
end;

{ TPlannerVertDayslineViewEh }

constructor TPlannerVertDayslineViewEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  TimeRange := dlrWeekEh;
end;

destructor TPlannerVertDayslineViewEh.Destroy;
begin
  inherited Destroy;
end;

procedure TPlannerVertDayslineViewEh.DrawDaySplitModeDateCell(ACol,
  ARow: Integer; ARect: TRect; State: TGridDrawState; ADataRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
begin
  WriteTextEh(Canvas, ARect, True, DrawArgs.HorzMargin, DrawArgs.VertMargin,
    DrawArgs.Text, DrawArgs.Alignment, DrawArgs.Layout, False, False, 0, 0,
    False, True);
end;

procedure TPlannerVertDayslineViewEh.GetDateCellDrawParams(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
var
  CellDateTime: TDateTime;
  wn: String;
begin
  CellDateTime := CellToDateTime(ACol, ARow);

  DrawArgs.FontColor := DatesBarArea.Font.Color;
  DrawArgs.FontName := DatesBarArea.Font.Name;
  DrawArgs.FontSize := DatesBarArea.Font.Size;
  DrawArgs.FontStyle := DatesBarArea.Font.Style;
  DrawArgs.BackColor := DatesBarArea.GetActualColor;

  wn := FormatSettings.ShortDayNames[DayOfWeek(CellDateTime)];
  DrawArgs.Text := wn + ', ' + FormatDateTime('D', CellDateTime);
  DrawArgs.Value := CellDateTime;
  DrawArgs.HorzMargin := Canvas.TextWidth('0') div 2;
  DrawArgs.VertMargin := 2;
  DrawArgs.Alignment := taLeftJustify;
  DrawArgs.Layout := tlTop;
end;

function TPlannerVertDayslineViewEh.GetDefaultDatesBarSize: Integer;
begin
  Result := 45;
end;

function TPlannerVertDayslineViewEh.GetDefaultDatesBarVisible: Boolean;
begin
  Result := True;
end;

function TPlannerVertDayslineViewEh.GetTimeRange: TDayslineRangeEh;
begin
  case RangeKind of
    rkWeekByDaysEh: Result := dlrWeekEh;
    rkMonthByDaysEh: Result := dlrMonthEh;
  else
    raise Exception.Create('TPlannerVertDayslineViewEh.RangeKind is inavlide');
  end;
end;

procedure TPlannerVertDayslineViewEh.SetTimeRange(const Value: TDayslineRangeEh);
const
  RangeDays: array [TDayslineRangeEh] of TTimePlanRangeKind = (rkWeekByDaysEh, rkMonthByDaysEh);
begin
  RangeKind := RangeDays[Value];
end;

procedure TPlannerVertDayslineViewEh.BuildDaysGridData;
var
  i: Integer;
  ColGroups: Integer;
  ARowCount: Integer;
  AFixedRowCount, AFixedColCount: Integer;
//  FixedDataWidth: Integer;
  ColWidth: Integer;
  FitGap: Integer;
  ADayGroupColWidth: Integer;
  DataColsWidth: Integer;
//  h: Integer;
begin


  AFixedRowCount := TopGridLineCount;

  if ResourceCaptionArea.Visible then
  begin
    ColGroups := ResourcesCount;
    if FShowUnassignedResource then
      Inc(ColGroups);
    if ColGroups = 0 then
      ColGroups := 1;

    FResourceAxisPos := AFixedRowCount;
    AFixedRowCount := AFixedRowCount + 1;
    ARowCount := FCellsInBand + AFixedRowCount;
  end else
  begin
    ColGroups := ResourcesCount;
    if FShowUnassignedResource then
      Inc(ColGroups);
    if ColGroups = 0 then
      ColGroups := 1;

    ARowCount := FCellsInBand + AFixedRowCount;
//    AFixedRowCount := 0;
    FResourceAxisPos := -1;
  end;

  if DayNameArea.Visible then
  begin
    FDayNameBarPos := AFixedRowCount;
    Inc(ARowCount);
    Inc(AFixedRowCount);
  end else
    FDayNameBarPos := -1;


  if DatesBarArea.Visible then
  begin
    AFixedColCount := 1;
    FDataColsOffset := 1;
    FDaySplitModeCol := 0;
  end else
  begin
    AFixedColCount := 0;
    FDataColsOffset := 0;
    FDaySplitModeCol := -1;
  end;

  FHoursBarIndex := -1;
  FDayGroupCol := -1;

  FixedRowCount := AFixedRowCount;
  FixedColCount := AFixedColCount;
  FDataRowsOffset := AFixedRowCount;

  ColCount := AFixedColCount + ColGroups + (ColGroups-1);
  RowCount := ARowCount;
  SetGridSize(FullColCount, FullRowCount);
  if ColGroups = 1
    then FBarsPerRes := 1
    else FBarsPerRes := 2;


  if FDaySplitModeCol >= 0 then
  begin
    ColWidths[FDaySplitModeCol] := DatesBarArea.GetActualSize;
    ADayGroupColWidth := DatesBarArea.GetActualSize;
  end else
    ADayGroupColWidth := 0;


  DataColsWidth := (GridClientWidth - ADayGroupColWidth - (ColGroups-1)*3);
  ColWidth := DataColsWidth div ColGroups;
  if ColWidth < MinDataColWidth then
  begin
    ColWidth := MinDataColWidth;
    HorzScrollBar.VisibleMode := sbAutoShowEh;
    FitGap := 0;
  end else
  begin
    HorzScrollBar.VisibleMode := sbNeverShowEh;
    FitGap := DataColsWidth mod ColGroups;
  end;

  for i := FixedColCount to ColCount-1 do
  begin
    if IsInterResourceCell(i, 0) then
    begin
      ColWidths[i] := 3;
    end else
    begin
      if FitGap > 0
        then ColWidths[i] := ColWidth + 1
        else ColWidths[i] := ColWidth;
      Dec(FitGap);
    end;
  end;


  if HandleAllocated then
  begin
    DefaultRowHeight := DataBarsArea.GetActualRowHeight;
    for i := 0 to RowCount-1 do
      RowHeights[i] := DefaultRowHeight;

    if TopGridLineCount > 0 then
      RowHeights[FTopGridLineIndex] := 1;

    if FResourceAxisPos >= 0 then
      RowHeights[FResourceAxisPos] := ResourceCaptionArea.GetActualSize;

    if FDayNameBarPos >= 0 then
      RowHeights[FDayNameBarPos] := DayNameArea.GetActualSize;
  end;

  MergeCells(0,0, FixedColCount-1,FixedRowCount-1);

end;

{ TPlannerHorzBandViewEh }

constructor TPlannerHorzTimelineViewEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FResourceColWidth := 100;
  BandOreintation := atboHorizonalEh;
  RangeKind := rkWeekByHoursEh;
  FShowDateRow := True;
end;

destructor TPlannerHorzTimelineViewEh.Destroy;
begin
  inherited Destroy;
end;

function TPlannerHorzTimelineViewEh.CellToDateTime(ACol, ARow: Integer): TDateTime;
begin
  case RangeKind of
    rkDayByHoursEh:
      Result := IncMinute(StartDate, 30 * (ACol - FixedColCount));
    rkWeekByHoursEh:
      Result := IncMinute(StartDate, 30 * (ACol - FixedColCount));
    rkWeekByDaysEh:
      Result := IncDay(StartDate, (ACol - FixedColCount));
    rkMonthByDaysEh:
      Result := IncDay(StartDate, (ACol - FixedColCount));
  else
    Result := 0;
  end;
end;

procedure TPlannerHorzTimelineViewEh.CheckDrawCellBorder(ACol, ARow: Integer;
  BorderType: TGridCellBorderTypeEh; var IsDraw: Boolean;
  var BorderColor: TColor; var IsExtent: Boolean);
var
  DataCell: TGridCoord;
  DataCol: Integer;
  NextCellDateTime, CellDateTime: TDateTime;
  NextCellOffset: Integer;
  Resource: TPlannerResourceEh;
  ResIndex: Integer;
begin
  inherited CheckDrawCellBorder(ACol, ARow, BorderType, IsDraw, BorderColor, IsExtent);

  if ((ARow-FDataRowsOffset >= 0) or ((ARow-FDataRowsOffset = -1 ) and (BorderType = cbtBottomEh))) and
     (ResourcesCount > 0) then
  begin
    Resource := nil;
    if IsInterResourceCell(ACol, ARow) then
    begin
      if BorderType in [cbtLeftEh, cbtRightEh]  then
        IsDraw := False;
      if (ARow-FDataRowsOffset) div FBarsPerRes < PlannerDataSource.Resources.Count then
        Resource := PlannerDataSource.Resources[(ARow-FDataRowsOffset) div FBarsPerRes];
    end else
    begin
      ResIndex := (ARow-FDataRowsOffset) div FBarsPerRes;
      if (ResIndex >= 0) and (ResIndex < PlannerDataSource.Resources.Count) then
        Resource := PlannerDataSource.Resources[ResIndex];
    end;
    if (Resource <> nil) and (Resource.DarkLineColor <> clDefault) then
        BorderColor := Resource.DarkLineColor;
  end;

  if RangeKind in [rkDayByHoursEh, rkWeekByHoursEh] then
  begin
    DataCell := GridCoordToDataCoord(GridCoord(ACol, ARow));

    CellDateTime := CellToDateTime(ACol, ARow);
    if (ACol = 1)
      then NextCellOffset := 2
      else NextCellOffset := 1;

    if ACol < ColCount-NextCellOffset
      then NextCellDateTime := CellToDateTime(ACol+NextCellOffset, ARow)
      else NextCellDateTime := CellDateTime;

    if ACol >= FixedColCount then
    begin
      if (BorderType = cbtRightEh) and
         ( (ARow = FDayGroupRow) or
           ( (DateOf(CellDateTime) <> DateOf(NextCellDateTime)) )
         )  then
    //    BorderColor := cl3DDkShadow;
        IsDraw := False
      else if IsDraw and (DataCell.X >= 0) and (DataCell.Y >= 0) then
      begin
        DataCol := DataCell.X;
        if (BorderType in [cbtLeftEh, cbtRightEh]) and (DataCol mod 2 = 0) then
          BorderColor := ApproximateColor(BorderColor, Color, 255 div 2);
      end;
    end;
  end;
end;

procedure TPlannerHorzTimelineViewEh.BuildHoursGridData;
var
  i, g, ir: Integer;
  RowGroups: Integer;
  AColCount: Integer;
  AFixedRowCount, AFixedColCount: Integer;
  AHoursBarHeight, ADayGroupRowHeight: Integer;
//  h: Integer;
  RowHeight, FitGap: Integer;
  DataRowsHeight: Integer;
  ATimeGroups: Integer;
  InGroupStartCol: Integer;
  BarsInGroup: Integer;
begin

  ClearGridCells;


  if ResourceCaptionArea.Visible then
  begin
    RowGroups := ResourcesCount;
    if FShowUnassignedResource then
      Inc(RowGroups);
    if RowGroups = 0  then
      RowGroups := 1;
    AColCount := FCellsInBand + 1;
    AFixedColCount := 1;
    FResourceAxisPos := 0;
  end else
  begin
    RowGroups := ResourcesCount;
    if FShowUnassignedResource then
      Inc(RowGroups);
    if RowGroups = 0  then
      RowGroups := 1;
    AColCount := FCellsInBand;
    AFixedColCount := 0;
    FResourceAxisPos := -1;
  end;

  if DayNameArea.Visible then
  begin
    FDayNameBarPos := AFixedColCount;
    Inc(AColCount);
    Inc(AFixedColCount);
  end else
    FDayNameBarPos := -1;


  AFixedRowCount := TopGridLineCount;

  if DatesBarArea.Visible then
  begin
    FDayGroupRow := AFixedRowCount;
    AFixedRowCount := AFixedRowCount + 1;
    FDataRowsOffset := AFixedRowCount;
  end else
  begin
//    AFixedRowCount := 0;
    FDataRowsOffset := AFixedRowCount;
    FDayGroupRow := -1;
  end;

  if HoursBarArea.Visible then
  begin
    FHoursBarIndex := AFixedRowCount;
    Inc(AFixedRowCount);
    Inc(FDataRowsOffset);
  end else
  begin
    FHoursBarIndex := -1;
  end;

  FixedRowCount := AFixedRowCount;
  FixedColCount := AFixedColCount;
  FDataColsOffset := AFixedColCount;
  FDaySplitModeRow := -1;

  RowCount := AFixedRowCount + RowGroups + (RowGroups-1);
  ColCount := AColCount;
  SetGridSize(FullColCount, FullRowCount);
  if RowGroups = 1
    then FBarsPerRes := 1
    else FBarsPerRes := 2;

  if TopGridLineCount > 0 then
    RowHeights[FTopGridLineIndex] := 1;


  if FHoursBarIndex >= 0 then
  begin
    RowHeights[FHoursBarIndex] := HoursBarArea.GetActualSize;
    AHoursBarHeight := HoursBarArea.GetActualSize;
  end else
    AHoursBarHeight := 0;


  if FDayGroupRow >= 0 then
  begin
    RowHeights[FDayGroupRow] := DatesBarArea.GetActualSize;
    ADayGroupRowHeight := DatesBarArea.GetActualSize;
  end else
    ADayGroupRowHeight := 0;


  DataRowsHeight := (GridClientHeight - AHoursBarHeight - ADayGroupRowHeight - (RowGroups-1)*3);
  RowHeight := DataRowsHeight div RowGroups;
  if RowHeight < MinDataRowHeight then
  begin
    RowHeight := MinDataRowHeight;
    VertScrollBar.VisibleMode := sbAutoShowEh;
    FitGap := 0;
  end else
  begin
    VertScrollBar.VisibleMode := sbNeverShowEh;
    FitGap := DataRowsHeight mod RowGroups;
  end;

  for i := FixedRowCount to RowCount-1 do
  begin
    if IsInterResourceCell(0, i) then
    begin
      RowHeights[i] := 3;
    end else
    begin
      if FitGap > 0
        then RowHeights[i] := RowHeight + 1
        else RowHeights[i] := RowHeight;
      Dec(FitGap);
    end;
  end;


  if HandleAllocated then
  begin
    Canvas.Font := Font;
    DefaultColWidth := DataBarsArea.GetActualColWidth;
    for i := FixedColCount to ColCount-1 do
      ColWidths[i] := DefaultColWidth;

    if FResourceAxisPos >= 0 then
      ColWidths[FResourceAxisPos] := ResourceCaptionArea.GetActualSize;

    if FDayNameBarPos >= 0 then
      ColWidths[FDayNameBarPos] := DayNameArea.GetActualSize;
  end;


  MergeCells(0,0, FixedColCount-1,FixedRowCount-1);

  ATimeGroups := FCellsInBand div FDayGroupCols;
  for g := 0 to ATimeGroups - 1 do
  begin
    InGroupStartCol := g * FDayGroupCols + FixedColCount;
    BarsInGroup := FDayGroupCols div 2;
    if FDayGroupRow >= 0 then
      MergeCells(InGroupStartCol, FDayGroupRow, FDayGroupCols-1, 0);

    if FHoursBarIndex >= 0 then
      for i := 0 to BarsInGroup - 1 do
      begin
        ir := InGroupStartCol + i*2;
        MergeCells(ir, FHoursBarIndex, 2-1, 0);
        Cell[ir, FHoursBarIndex].Value := FormatFloat('00', i) + ':00';
      end;
  end;

end;

procedure TPlannerHorzTimelineViewEh.CalcLayouts;
begin
  inherited CalcLayouts;
  FDayGroupCols := -1;

  case RangeKind of
    rkDayByHoursEh:
      begin
        FDayGroupCols := 24 * 2;
      end;
    rkWeekByHoursEh:
      begin
        FDayGroupCols := 24 * 2;
      end;
    rkWeekByDaysEh:
      begin
        FDayGroupCols := 1;
      end;
    rkMonthByDaysEh:
      begin
        FDayGroupCols := 1;
      end;
  end;
end;

//procedure TPlannerHorzTimelineViewEh.DrawCell(ACol, ARow: Integer; ARect: TRect;
//  State: TGridDrawState);
//begin
//  if (ACol >= FixedColCount) and (ARow = FDayGroupRow) then
//    DrawTimeGroupCell(ACol, ARow, ARect, State)
//  else if (ACol >= FixedColCount) and (ARow = FHoursBarIndex) then
//    DrawTimeCell(ACol, ARow, ARect, State, 0, ARow-FixedRowCount)
//  else if (ACol = FDayNameBarPos) and (ARow >= FixedRowCount) then
//    DrawDayNamesCell(ACol, ARow, ARect, State, 0, 0)
//  else if (ACol >= FixedColCount) and (ARow = FDaySplitModeRow) then
//    DrawDaySplitModeDateCell(ACol, ARow, ARect, State, ACol-FixedColCount)
//  else if (ACol = 0) and (ARow < FixedRowCount) then
//    DrawTopLeftCell(ACol, ARow, ARect, State, 0, 0)
//  else if (ACol = FResourceAxisPos) and (ARow >= FixedRowCount) then
//  begin
//    if IsInterResourceCell(ACol, ARow)
//      then DrawInterResourceCell(ACol, ARow, ARect, State, 0, 0)
//      else DrawResourceCaptionCell(ACol, ARow, ARect, State, 0, 0);
//  end else
//    DrawDataCell(ACol, ARow, ARect, State, 0, 0);
//end;

procedure TPlannerHorzTimelineViewEh.GetCellType(ACol, ARow: Integer;
  var CellType: TPlannerViewCellTypeEh; var ALocalCol, ALocalRow: Integer);
begin
  if (ACol >= FixedColCount) and (ARow = FDayGroupRow) then
  begin
    CellType := pctDateBarEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end else if (ACol >= FixedColCount) and (ARow = FHoursBarIndex) then
  begin
    CellType := pctTimeCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end else if (ACol >= FixedColCount) and (ARow = FDaySplitModeRow) then
  begin
    CellType := pctDateCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end else if (ACol = FDayNameBarPos) and (ARow >= FixedRowCount) then
  begin
    CellType := pctDayNameCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end else if (ACol = 0) and (ARow < FixedRowCount) then
  begin
    CellType := pctTopLeftCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end else if (ACol = FResourceAxisPos) and (ARow >= FixedRowCount) then
  begin
     if IsInterResourceCell(ACol, ARow)
      then CellType := pctInterResourceSpaceEh
      else CellType := pctResourceCaptionCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end else
  begin
    CellType := pctDataCellEh;
    ALocalCol := 0;
    ALocalRow := 0;
  end;
end;

procedure TPlannerHorzTimelineViewEh.DrawDataCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
var
  NextCellDateTime, CellDateTime: TDateTime;
begin
  if RangeKind in [rkDayByHoursEh, rkWeekByHoursEh] then
  begin
    CellDateTime := CellToDateTime(ACol, ARow);
    if ACol < ColCount-1
      then NextCellDateTime := CellToDateTime(ACol+1, ARow)
      else NextCellDateTime := CellDateTime;

    if DateOf(CellDateTime) <> DateOf(NextCellDateTime) then
    begin

      Canvas.Pen.Color := FMasterGroupLineColor;
      DrawPolyline([Point(ARect.Right-1, ARect.Top), Point(ARect.Right-1, ARect.Bottom)]);
      DrawPolyline([Point(ARect.Right-2, ARect.Top), Point(ARect.Right-2, ARect.Bottom)]);

      ARect.Right := ARect.Right - 2;
    end;
  end;
  inherited DrawDataCell(ACol, ARow, ARect, State, 0, 0, DrawArgs);
end;

procedure TPlannerHorzTimelineViewEh.DrawDaySplitModeDateCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ADataCol: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
begin

end;

procedure TPlannerHorzTimelineViewEh.DrawTimeCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
var
  NextCellDateTime, CellDateTime: TDateTime;
  S: String;
  w: Integer;
  wn: String;
begin
  if RangeKind in [rkDayByHoursEh, rkWeekByHoursEh] then
  begin
    CellDateTime := CellToDateTime(ACol, ARow);
    if ACol < ColCount-2
      then NextCellDateTime := CellToDateTime(ACol+2, ARow)
      else NextCellDateTime := CellDateTime;

    if DateOf(CellDateTime) <> DateOf(NextCellDateTime) then
    begin
      Canvas.Pen.Color := FMasterGroupLineColor;
      DrawPolyline([Point(ARect.Right-1, ARect.Top), Point(ARect.Right-1, ARect.Bottom)]);
      DrawPolyline([Point(ARect.Right-2, ARect.Top), Point(ARect.Right-2, ARect.Bottom)]);

      ARect.Right := ARect.Right - 2;
    end;

    inherited DrawTimeCell(ACol, ARow, ARect, State, ALocalCol, ALocalRow, DrawArgs);
  end else
  begin
    CellDateTime := CellToDateTime(ACol, ARow);
    Canvas.Font := Font;
    wn := FormatSettings.ShortDayNames[DayOfWeek(CellDateTime)];
    s := wn + ', ' + FormatDateTime('D', CellDateTime);
    w := Canvas.TextWidth('0') div 2;
    WriteTextEh(Canvas, ARect, True, w, 2, s,
      taLeftJustify, tlTop, False, False, 0, 0, False, True);
  end;
end;

procedure TPlannerHorzTimelineViewEh.DrawDateBar(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
begin
  DrawTimeGroupCell(ACol, ARow, ARect, State);
end;

procedure TPlannerHorzTimelineViewEh.DrawDateCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
begin
  DrawDaySplitModeDateCell(ACol, ARow, ARect, State, ACol-FixedColCount, DrawArgs);
end;

procedure TPlannerHorzTimelineViewEh.DrawTimeGroupCell(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState);
var
  DrawDate: TDateTime;
  s: String;
begin

  Canvas.Pen.Color := FMasterGroupLineColor;
  DrawPolyline([Point(ARect.Right-1, ARect.Top), Point(ARect.Right-1, ARect.Bottom)]);
  DrawPolyline([Point(ARect.Right-2, ARect.Top), Point(ARect.Right-2, ARect.Bottom)]);

  ARect.Right := ARect.Right - 2;
  DrawDate := CellToDateTime(ACol, ARow);
  s := FormatDateTime(FormatSettings.LongDateFormat, DrawDate) + ' ';

  Canvas.Brush.Color := StyleServices.GetSystemColor(DatesBarArea.GetActualColor);
  Canvas.Font := DatesBarArea.Font;
  Canvas.Font.Color := StyleServices.GetSystemColor(DatesBarArea.Font.Color);

  Canvas.FillRect(ARect);
  if ARect.Left < HorzAxis.FixedBoundary then
    ARect.Left := HorzAxis.FixedBoundary;
  if ARect.Right > HorzAxis.ContraStart then
    ARect.Right := HorzAxis.ContraStart;

  WriteTextEh(Canvas, ARect, True, 0, 0, s, taCenter, tlTop, False, False, 0, 0, False, True);
end;

procedure TPlannerHorzTimelineViewEh.GetDateBarDrawParams(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
var
  DrawDate: TDateTime;
begin
  DrawDate := CellToDateTime(ACol, ARow);
  DrawArgs.Text := FormatDateTime(FormatSettings.LongDateFormat, DrawDate) + ' ';
  DrawArgs.Value := DrawDate;

  DrawArgs.FontName := DatesBarArea.Font.Name;
  DrawArgs.FontColor := DatesBarArea.Font.Color;
  DrawArgs.FontSize := DatesBarArea.Font.Size;
  DrawArgs.FontStyle := DatesBarArea.Font.Style;
  DrawArgs.BackColor := DatesBarArea.GetActualColor;
  DrawArgs.Alignment := taCenter;
  DrawArgs.Layout := tlCenter;
  DrawArgs.Orientation := tohHorizontal;
end;

function TPlannerHorzTimelineViewEh.GetResourceCaptionAreaDefaultSize: Integer;
var
  h: Integer;
begin
  if HandleAllocated then
  begin
    Canvas.Font := ResourceCaptionArea.Font;
    h := Trunc(Canvas.TextWidth('WWWWWWWWWW'));
    if h > DefaultColWidth
      then Result := h
      else Result := DefaultColWidth;
  end else
    Result := DefaultColWidth;
end;

function TPlannerHorzTimelineViewEh.GetResourceViewAtCell(ACol, ARow: Integer): Integer;
begin
  if ARow < FDataRowsOffset then
    Result := -1
  else if IsInterResourceCell(ACol, ARow) then
    Result := -1
  else
    Result := (ARow-FDataRowsOffset) div FBarsPerRes;
end;

function TPlannerHorzTimelineViewEh.IsInterResourceCell(ACol,
  ARow: Integer): Boolean;
begin
  if (ARow-FDataRowsOffset >= 0) and
     (Length(FResourcesView) > 1)
  then
    Result := (ARow-FDataRowsOffset) mod FBarsPerRes = FBarsPerRes - 1
  else
    Result := False;
end;

procedure TPlannerHorzTimelineViewEh.SetDisplayPosesSpanItems;
var
  ASpanRect: TRect;
  SpanItem: TTimeSpanDisplayItemEh;
  i: Integer;
  StartY: Integer;
  ResourceOffset: Integer;
  AStartViewDate, AEndViewDate: TDateTime;
begin
  SetResOffsets;
  StartY := 0; {VertAxis.FixedBoundary;}
  GetViewPeriod(AStartViewDate, AEndViewDate);

  for i := 0 to SpanItemsCount-1 do
  begin
    SpanItem := SpanItems[i];
//    if SpanItem.GridColNum = FDataColsOffset then
    begin
      ASpanRect.Left := {HorzAxis.FixedBoundary + }SpanItem.StartGridRollPos;
      ASpanRect.Right := {HorzAxis.FixedBoundary + }SpanItem.StopGridRolPosl;

      ASpanRect.Top := StartY + 10;
      ASpanRect.Bottom := ASpanRect.Top + RowHeights[FDataRowsOffset] - 20;

      ResourceOffset := GetGridOffsetForResource(SpanItem.PlanItem.Resource);
      OffsetRect(ASpanRect, 0, ResourceOffset);

      CalcRectForInCellRows(SpanItem, ASpanRect);

      SpanItem.FBoundRect := ASpanRect;

      if (SpanItem.PlanItem.StartTime < AStartViewDate) then
        SpanItem.FDrawBackOutInfo := True;
      if (SpanItem.PlanItem.EndTime > AEndViewDate) then
        SpanItem.FDrawForwardOutInfo := True;
    end;
  end;
end;

procedure TPlannerHorzTimelineViewEh.SetMinDataRowHeight(const Value: Integer);
begin
  if FMinDataRowHeight <> Value then
  begin
    FMinDataRowHeight := Value;
    GridLayoutChanged;
  end;
end;

procedure TPlannerHorzTimelineViewEh.CalcRectForInCellRows(
  SpanItem: TTimeSpanDisplayItemEh; var DrawRect: TRect);
var
  RectHeight: Integer;
  SecHeight: Integer;
  OldDrawRect: TRect;
begin
  RectHeight := DrawRect.Bottom - DrawRect.Top;
  if SpanItem.InCellRows > 0 then
  begin
    SecHeight := RectHeight div SpanItem.InCellRows;
    OldDrawRect := DrawRect;
    DrawRect.Top := OldDrawRect.Top + SecHeight * SpanItem.InCellFromRow;
    DrawRect.Bottom := OldDrawRect.Top + SecHeight * (SpanItem.InCellToRow + 1);
  end;
end;

procedure TPlannerHorzTimelineViewEh.SetResOffsets;
var
  i: Integer;
begin
  for i := 0 to Length(FResourcesView)-1 do
    if i*FBarsPerRes < VertAxis.RolCelCount then
    begin
//      if (PlannerDataSource <> nil) and (i < PlannerDataSource.Resources.Count)
//        then FResourcesView[i].Resource := PlannerDataSource.Resources[i]
//        else FResourcesView[i].Resource := nil;
      FResourcesView[i].GridOffset := VertAxis.RolLocCelPosArr[i*FBarsPerRes];
      FResourcesView[i].GridStartAxisBar := i*FBarsPerRes;
    end;
end;

procedure TPlannerHorzTimelineViewEh.SetResourceColWidth(const Value: Integer);
begin
  if FResourceColWidth <> Value then
  begin
    FResourceColWidth := Value;
    GridLayoutChanged;
  end;
end;

function TPlannerHorzTimelineViewEh.CalTimeRowHeight: Integer;
var
  hSize, mSize: TSize;
begin
  Result := 12;
  if not HandleAllocated then Exit;
  if RangeKind in [rkDayByHoursEh, rkWeekByHoursEh] then
  begin
    Canvas.Font := Font;
    mSize := Canvas.TextExtent('00');
    Canvas.Font.Size := Font.Size * 3 div 2;
    hSize := Canvas.TextExtent('00');
    Result := hSize.cx + mSize.cx div 2 + mSize.cx;
    Canvas.Font := Font;
  end else
  begin
    Canvas.Font := Font;
    mSize := Canvas.TextExtent('00');
    Canvas.Font.Size := Font.Size * 3 div 2;
    hSize := Canvas.TextExtent('00');
    Result := hSize.cx + mSize.cx div 2;
    Canvas.Font := Font;
  end;
end;

procedure TPlannerHorzTimelineViewEh.InitSpanItem(
  ASpanItem: TTimeSpanDisplayItemEh);
begin
  inherited InitSpanItem(ASpanItem);
  ASpanItem.FAllowedInteractiveChanges :=
    [sichSpanMovingEh, sichSpanLeftSizingEh, sichSpanRightSizingEh];
  ASpanItem.FTimeOrientation := toHorizontalEh;
end;

procedure TPlannerHorzTimelineViewEh.MouseMove(Shift: TShiftState; X,
  Y: Integer);
begin
  inherited MouseMove(Shift, X, Y);
end;

procedure TPlannerHorzTimelineViewEh.InitSpanItemMoving(
  SpanItem: TTimeSpanDisplayItemEh; MousePos: TPoint);
var
  ACell: TGridCoord;
  ACellTime: TDateTime;
begin
  ACell := MouseCoord(MousePos.X, MousePos.Y);
  ACellTime := CellToDateTime(ACell.X, ACell.Y);
  FMovingDaysShift := DaysBetween(DateOf(SpanItem.PlanItem.StartTime), DateOf(ACellTime));
  FMovingTimeShift := ACellTime - SpanItem.PlanItem.StartTime;
end;

procedure TPlannerHorzTimelineViewEh.UpdateDummySpanItemSize(MousePos: TPoint);
var
  FixedMousePos: TPoint;
  ACell: TGridCoord;
  ANewTime: TDateTime;
  ATimeLen: TDateTime;
  ResViewIdx: Integer;
  AResource: TPlannerResourceEh;
begin
  FixedMousePos := MousePos;
  if FPlannerState in [psSpanMovingEh, psSpanTestMovingEh] then
  begin
    FixedMousePos.X := FixedMousePos.X + FTopLeftSpanShift.cx;
    FixedMousePos.Y := FixedMousePos.Y + FTopLeftSpanShift.cy;
  end;
  if FPlannerState = psSpanRightSizingEh then
  begin
    ACell := CalcCoordFromPoint(FixedMousePos.X + DefaultColWidth div 2, FixedMousePos.Y);
    ANewTime := CellToDateTime(ACell.X, ACell.Y);
    if (FDummyPlanItem.StartTime < ANewTime) and (FDummyPlanItem.EndTime <> ANewTime) then
    begin
      FDummyCheckPlanItem.Assign(FDummyPlanItem);
      FDummyCheckPlanItem.EndTime := ANewTime;
      CheckSetDummyPlanItem(FDummyPlanItemFor, FDummyCheckPlanItem);
      PlannerDataSourceChanged();
    end;
  end else if FPlannerState = psSpanLeftSizingEh then
  begin
    ACell := MouseCoord(FixedMousePos.X + DefaultColWidth  div 2, FixedMousePos.Y);
    ANewTime := CellToDateTime(ACell.X, ACell.Y);
    if (FDummyPlanItem.EndTime > ANewTime) and (FDummyPlanItem.StartTime <> ANewTime) then
    begin
      FDummyCheckPlanItem.Assign(FDummyPlanItem);
      FDummyCheckPlanItem.StartTime := ANewTime;
      CheckSetDummyPlanItem(FDummyPlanItemFor, FDummyCheckPlanItem);
      PlannerDataSourceChanged();
    end;
  end else if FPlannerState in [psSpanMovingEh, psSpanTestMovingEh] then
  begin
    ACell := CalcCoordFromPoint(FixedMousePos.X, FixedMousePos.Y);
    ANewTime := CellToDateTime(ACell.X, ACell.Y);
    ResViewIdx := GetResourceViewAtCell(ACell.X, ACell.Y);
    if ResViewIdx >= 0
      then AResource := FResourcesView[ResViewIdx].Resource
      else AResource := FDummyPlanItem.Resource;

    if (FDummyPlanItem.StartTime <> ANewTime) or
       (AResource <> FDummyPlanItem.Resource) then
    begin
      ATimeLen :=  FDummyPlanItem.EndTime - FDummyPlanItem.StartTime;
      FDummyCheckPlanItem.Assign(FDummyPlanItem);
      FDummyCheckPlanItem.StartTime := ANewTime - FMovingTimeShift;
      FDummyCheckPlanItem.EndTime := FDummyCheckPlanItem.StartTime + ATimeLen;
      FDummyCheckPlanItem.Resource := AResource;
      CheckSetDummyPlanItem(FDummyPlanItemFor, FDummyCheckPlanItem);
      PlannerDataSourceChanged;
    end;
    ShowMoveHintWindow(FDummyPlanItem, MousePos);
  end;
end;

procedure TPlannerHorzTimelineViewEh.SetShowDateRow(const Value: Boolean);
begin
  if FShowDateRow <> Value then
  begin
    FShowDateRow := Value;
    PlannerDataSourceChanged;
  end;
end;

function TPlannerHorzTimelineViewEh.DefaultHoursBarSize: Integer;
begin
  if HandleAllocated then
  begin
    Canvas.Font := HoursBarArea.Font;
    Result := Canvas.TextHeight('Wg') * 2;
  end else
    Result := 45;
end;

function TPlannerHorzTimelineViewEh.CreateDayNameArea: TDayNameAreaEh;
begin
  Result := TDayNameHorzAreaEh.Create(Self);
end;

function TPlannerHorzTimelineViewEh.GetDayNameArea: TDayNameHorzAreaEh;
begin
  Result := TDayNameHorzAreaEh(inherited DayNameArea);
end;

procedure TPlannerHorzTimelineViewEh.SetDayNameArea(const Value: TDayNameHorzAreaEh);
begin
  inherited DayNameArea := Value;
end;

function TPlannerHorzTimelineViewEh.CreateDataBarsArea: TDataBarsAreaEh;
begin
  Result := TDataBarsHorzAreaEh.Create(Self);
end;

function TPlannerHorzTimelineViewEh.GetDataBarsArea: TDataBarsHorzAreaEh;
begin
  Result := TDataBarsHorzAreaEh(inherited DataBarsArea);
end;

procedure TPlannerHorzTimelineViewEh.SetDataBarsArea(
  const Value: TDataBarsHorzAreaEh);
begin
  inherited DataBarsArea := Value;
end;

function TPlannerHorzTimelineViewEh.CreateResourceCaptionArea: TResourceCaptionAreaEh;
begin
  Result := TResourceHorzCaptionAreaEh.Create(Self);
end;

function TPlannerHorzTimelineViewEh.GetResourceCaptionArea: TResourceHorzCaptionAreaEh;
begin
  Result := TResourceHorzCaptionAreaEh(inherited ResourceCaptionArea);
end;

procedure TPlannerHorzTimelineViewEh.SetResourceCaptionArea(const Value: TResourceHorzCaptionAreaEh);
begin
  inherited SetResourceCaptionArea(Value);
end;

function TPlannerHorzTimelineViewEh.CreateHoursBarArea: THoursBarAreaEh;
begin
  Result := THoursHorzBarAreaEh.Create(Self);
end;

function TPlannerHorzTimelineViewEh.GetHoursColArea: THoursHorzBarAreaEh;
begin
  Result := THoursHorzBarAreaEh(inherited HoursBarArea);
end;

procedure TPlannerHorzTimelineViewEh.SetHoursColArea(const Value: THoursHorzBarAreaEh);
begin
  inherited HoursBarArea := Value;
end;

function TPlannerHorzTimelineViewEh.GetDatesRowArea: TDatesRowAreaEh;
begin
  Result := TDatesRowAreaEh(DatesBarArea);
end;

procedure TPlannerHorzTimelineViewEh.SetDatesRowArea(
  const Value: TDatesRowAreaEh);
begin
  DatesBarArea := Value;
end;

function TPlannerHorzTimelineViewEh.CreateDatesBarArea: TDatesBarAreaEh;
begin
  Result := TDatesRowAreaEh.Create(Self);
end;

{ TPlannerHorzDayslineViewEh }

constructor TPlannerHorzDayslineViewEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  TimeRange := dlrWeekEh;
end;

destructor TPlannerHorzDayslineViewEh.Destroy;
begin
  inherited Destroy;
end;

procedure TPlannerHorzDayslineViewEh.BuildDaysGridData;
var
  i: Integer;
  RowGroups: Integer;
  AColCount: Integer;
  AFixedRowCount, AFixedColCount: Integer;
//  FixedDataWidth: Integer;
  RowHeight: Integer;
  FitGap: Integer;
  ADayGroupRowHeight: Integer;
  DataRowsHeight: Integer;
//  h: Integer;
begin


  AFixedRowCount := TopGridLineCount;

  if ResourceCaptionArea.Visible then
  begin
    RowGroups := ResourcesCount;
    if FShowUnassignedResource then
      Inc(RowGroups);
    if RowGroups = 0  then
      RowGroups := 1;

    AColCount := FCellsInBand + 1;
    AFixedColCount := 1;
    FResourceAxisPos := 0;
  end else
  begin
    RowGroups := ResourcesCount;
    if FShowUnassignedResource then
      Inc(RowGroups);
    if RowGroups = 0  then
      RowGroups := 1;

    AColCount := FCellsInBand;
    AFixedColCount := 0;
    FResourceAxisPos := -1;
  end;

  if DayNameArea.Visible then
  begin
    FDayNameBarPos := AFixedColCount;
    Inc(AColCount);
    Inc(AFixedColCount);
  end else
    FDayNameBarPos := -1;


  if DatesBarArea.Visible then
  begin
    FDaySplitModeRow := AFixedRowCount;
    AFixedRowCount := AFixedRowCount + 1;
    FDataRowsOffset := AFixedRowCount;
  end else
  begin
//    AFixedRowCount := 0;
    FDataRowsOffset := AFixedRowCount;
    FDaySplitModeRow := -1;
  end;

  FHoursBarIndex := -1;
  FDayGroupRow := -1;

  FixedRowCount := AFixedRowCount;
  FixedColCount := AFixedColCount;
  FDataColsOffset := AFixedColCount;

  ColCount := AColCount;
  RowCount := AFixedRowCount + RowGroups + (RowGroups-1);
  SetGridSize(FullColCount, FullRowCount);
  if RowGroups = 1
    then FBarsPerRes := 1
    else FBarsPerRes := 2;

  if TopGridLineCount > 0 then
    RowHeights[FTopGridLineIndex] := 1;


  if FDaySplitModeRow >= 0 then
  begin
    RowHeights[FDaySplitModeRow] := DatesBarArea.GetActualSize;
    ADayGroupRowHeight := DatesBarArea.GetActualSize;
  end else
    ADayGroupRowHeight := 0;


  DataRowsHeight := (GridClientHeight - ADayGroupRowHeight - (RowGroups-1)*3);
  RowHeight := DataRowsHeight div RowGroups;
  if RowHeight < MinDataRowHeight then
  begin
    RowHeight := MinDataRowHeight;
    VertScrollBar.VisibleMode := sbAutoShowEh;
    FitGap := 0;
  end else
  begin
    VertScrollBar.VisibleMode := sbNeverShowEh;
    FitGap := DataRowsHeight mod RowGroups;
  end;

  for i := FixedRowCount to RowCount-1 do
  begin
    if IsInterResourceCell(0, i) then
    begin
      RowHeights[i] := 3;
    end else
    begin
      if FitGap > 0
        then RowHeights[i] := RowHeight + 1
        else RowHeights[i] := RowHeight;
      Dec(FitGap);
    end;
  end;


  if HandleAllocated then
  begin
    Canvas.Font := Font;
    DefaultColWidth := DataBarsArea.GetActualColWidth;
    for i := 0 to ColCount-1 do
      ColWidths[i] := DefaultColWidth;

    if FResourceAxisPos >= 0 then
      ColWidths[FResourceAxisPos] := ResourceCaptionArea.GetActualSize;

    if FDayNameBarPos >= 0 then
      ColWidths[FDayNameBarPos] := DayNameArea.GetActualSize;
  end;

  MergeCells(0,0, FixedColCount-1,FixedRowCount-1);

end;

procedure TPlannerHorzDayslineViewEh.DrawDaySplitModeDateCell(ACol,
  ARow: Integer; ARect: TRect; State: TGridDrawState; ADataRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
begin
  WriteTextEh(Canvas, ARect, True, DrawArgs.HorzMargin, DrawArgs.VertMargin,
    DrawArgs.Text, DrawArgs.Alignment, DrawArgs.Layout, False, False, 0, 0,
    False, True);
end;

procedure TPlannerHorzDayslineViewEh.GetDateCellDrawParams(ACol, ARow: Integer;
  ARect: TRect; State: TGridDrawState; ALocalCol, ALocalRow: Integer;
  DrawArgs: TPlannerViewCellDrawArgsEh);
var
  CellDateTime: TDateTime;
  wn: String;
begin
  CellDateTime := CellToDateTime(ACol, ARow);

  DrawArgs.FontColor := DatesBarArea.Font.Color;
  DrawArgs.FontName := DatesBarArea.Font.Name;
  DrawArgs.FontSize := DatesBarArea.Font.Size;
  DrawArgs.FontStyle := DatesBarArea.Font.Style;
  DrawArgs.BackColor := DatesBarArea.GetActualColor;

  wn := FormatSettings.ShortDayNames[DayOfWeek(CellDateTime)];
  DrawArgs.Text := wn + ', ' + FormatDateTime('D', CellDateTime);
  DrawArgs.Value := CellDateTime;
  DrawArgs.HorzMargin := Canvas.TextWidth('0') div 2;
  DrawArgs.VertMargin := 2;
  DrawArgs.Alignment := taLeftJustify;
  DrawArgs.Layout := tlTop;
end;

function TPlannerHorzDayslineViewEh.GetTimeRange: TDayslineRangeEh;
begin
  case RangeKind of
    rkWeekByDaysEh: Result := dlrWeekEh;
    rkMonthByDaysEh: Result := dlrMonthEh;
  else
    raise Exception.Create('TPlannerVertDayslineViewEh.RangeKind is inavlide');
  end;
end;

procedure TPlannerHorzDayslineViewEh.SetTimeRange(
  const Value: TDayslineRangeEh);
const
  RangeDays: array [TDayslineRangeEh] of TTimePlanRangeKind = (rkWeekByDaysEh, rkMonthByDaysEh);
begin
  RangeKind := RangeDays[Value];
end;

{ TPlannerGridLineParamsEh }

constructor TPlannerGridLineParamsEh.Create(AGrid: TCustomGridEh);
begin
  inherited Create(AGrid);
  FPaleColor := clDefault;
end;

function TPlannerGridLineParamsEh.DefaultPaleColor: TColor;
begin
  Result := ApproximateColor(GetBrightColor,
    StyleServices.GetSystemColor(PlannerView.Color), 255 div 2);
end;

function TPlannerGridLineParamsEh.GetBrightColor: TColor;
begin
{$IFDEF EH_LIB_7}
  if BrightColor = clDefault then
  begin
    if {Grid.DrawTitleByThemes and} ThemesEnabled {and (Win32MajorVersion >= 6)} then
      Result := StyleServices.GetSystemColor(clBtnFace)
    else
      Result := inherited GetBrightColor
  end else
    Result := inherited GetBrightColor;
{$ELSE}
  Result := inherited GetBrightColor;
{$ENDIF}
end;

function TPlannerGridLineParamsEh.GetDarkColor: TColor;
{$IFDEF EH_LIB_16}
const
  CFixedStates: array[Boolean] of TThemedGrid =
    (tgGradientFixedCellNormal, tgFixedCellNormal);
   // glcsFlatEh,                glcsThemedEh
{$ENDIF}
{$IFDEF FPC}
{$ELSE}
var
  LColorRef: TColorRef;
{$ENDIF}
{$IFDEF EH_LIB_16}
  LStyle: TCustomStyleServices;
{$ENDIF}
begin
  if DarkColor = clDefault then
  begin
  {$IFDEF EH_LIB_7}
    {$IFDEF EH_LIB_16}
    if CustomStyleActive then
    begin
      LStyle := StyleServices;
      LStyle.GetElementColor(LStyle.GetElementDetails(CFixedStates[Grid.Flat]), ecBorderColor, Result);
    end else
    {$ENDIF}
    {$IFDEF FPC}
    {$ELSE}
    if {Grid.DrawTitleByThemes and} ThemesEnabled and (Win32MajorVersion >= 6) then
    begin
      GetThemeColor(ThemeServices.Theme[teHeader], HP_HEADERITEM, HIS_NORMAL,
        TMT_EDGEFILLCOLOR, LColorRef);
      Result := LColorRef;
    end else
    {$ENDIF}
  {$ENDIF}
      Result := inherited GetDarkColor;
  end else
    Result := inherited GetDarkColor;
end;

function TPlannerGridLineParamsEh.GetPaleColor: TColor;
begin
  if PaleColor <> clDefault
    then Result := PaleColor
    else Result := DefaultPaleColor;
end;

function TPlannerGridLineParamsEh.GetPlannerView: TCustomPlannerViewEh;
begin
  Result := TCustomPlannerViewEh(Grid);
end;

procedure TPlannerGridLineParamsEh.SetPaleColor(const Value: TColor);
begin
  if PaleColor <> Value then
  begin
    FPaleColor := Value;
    Grid.Invalidate;
  end;
end;

{ TPlannerGridControlEh }

constructor TPlannerGridControlEh.Create(AGrid: TCustomPlannerViewEh);
begin
  inherited Create(AGrid);
  FGrid := AGrid;
  DoubleBuffered := True;
end;

destructor TPlannerGridControlEh.Destroy;
begin
  inherited Destroy;
end;

procedure TPlannerGridControlEh.Click;
begin
  if ControlType = pgctNextEventEh then
    FGrid.GotoNextEventInTheFuture
  else
    FGrid.GotoPriorEventInThePast;
end;

procedure TPlannerGridControlEh.Paint;
var
  AText: String;
begin
  Canvas.Brush.Color := FGrid.EventNavBoxColor;
  Canvas.FillRect(Rect(0,0,Width,Height));
  Canvas.Brush.Color := FGrid.EventNavBoxBorderColor;
  Canvas.FrameRect(Rect(0,0,Width,Height));
  if ControlType = pgctNextEventEh
    then AText := SPlannerNextEventEh
    else AText := SPlannerPriorEventEh;

  Canvas.Font := FGrid.EventNavBoxFont;
  Canvas.Font.Color := StyleServices.GetSystemColor(FGrid.EventNavBoxFont.Color);
  if not ClickEnabled then
    Canvas.Font.Color := StyleServices.GetSystemColor(clGrayText);

  WriteTextVerticalEh(Canvas, Rect(0,0,Width,Height), False, 0, 0, AText,
    taCenter, tlCenter, False, False);
end;

procedure TPlannerGridControlEh.Realign;
begin
  FGrid.RealignGridControl(Self);
end;

{ TInGridControlEh }

constructor TInGridControlEh.Create(AGrid: TCustomPlannerViewEh);
begin
  inherited Create;
  FGrid := AGrid;
end;

destructor TInGridControlEh.Destroy;
begin

  inherited Destroy;
end;

procedure TInGridControlEh.Assign(Source: TPersistent);
begin
  if Source is TInGridControlEh then
  begin
    FHorzLocating := TInGridControlEh(Source).HorzLocating;
    FVertLocating := TInGridControlEh(Source).VertLocating;
    FBoundRect := TInGridControlEh(Source).BoundRect;
  end;
end;

procedure TInGridControlEh.DblClick;
begin

end;

procedure TInGridControlEh.GetInGridDrawRect(var ARect: TRect);
begin
  ARect := BoundRect;
  if HorzLocating = brrlGridRolAreaEh then
    OffsetRect(ARect, PlannerView.HorzAxis.FixedBoundary - PlannerView.HorzAxis.RolStartVisPos, 0);
  if VertLocating = brrlGridRolAreaEh then
    OffsetRect(ARect, 0, PlannerView.VertAxis.FixedBoundary - PlannerView.VertAxis.RolStartVisPos);
end;

{ TMoveHintWindow }

procedure TMoveHintWindow.WMEraseBkgnd(var Message: TWmEraseBkgnd);
begin
  Message.Result := 1;
end;

{ TTimeSpanParamsEh }

constructor TDrawElementParamsEh.Create;
begin
  inherited Create;
  FFont := TFont.Create;
//  FFont.Assign(DefaultFont);
  FFont.OnChange := FontChanged;

  FColor := clDefault;
  FFontStored := False;
  FAltColor := clDefault;
  FFillStyle := fsDefaultEh;
  FBorderColor := clDefault;
  FHue := clDefault;
end;

destructor TDrawElementParamsEh.Destroy;
begin
  FreeAndNil(FFont);
  inherited Destroy;
end;

procedure TDrawElementParamsEh.NotifyChanges;
begin
end;

procedure TDrawElementParamsEh.SetAltColor(const Value: TColor);
begin
  if FAltColor <> Value then
  begin
    FAltColor := Value;
    NotifyChanges;
  end;
end;

procedure TDrawElementParamsEh.SetBorderColor(const Value: TColor);
begin
  if FBorderColor <> Value then
  begin
    FBorderColor := Value;
    NotifyChanges;
  end;
end;

procedure TDrawElementParamsEh.SetColor(const Value: TColor);
begin
  if FColor <> Value then
  begin
    FColor := Value;
    NotifyChanges;
  end;
end;

procedure TDrawElementParamsEh.SetHue(const Value: TColor);
begin
  if FHue <> Value then
  begin
    FHue := Value;
    NotifyChanges;
  end;
end;

procedure TDrawElementParamsEh.SetFillStyle(const Value: TPropFillStyleEh);
begin
  if FFillStyle <> Value then
  begin
    FFillStyle := Value;
    NotifyChanges;
  end;
end;

function TDrawElementParamsEh.DefaultFont: TFont;
begin
  Result := nil;
end;

procedure TDrawElementParamsEh.FontChanged(Sender: TObject);
begin
  NotifyChanges;
  FFontStored := True;
//  raise Exception.Create('TTimeSpanParamsEh.FontChanged: ');
end;

function TDrawElementParamsEh.GetActualColor: TColor;
begin
  if Color <> clDefault
    then Result := Color
    else Result := GetDefaultColor;
end;

function TDrawElementParamsEh.GetActualAltColor: TColor;
begin
  if AltColor <> clDefault
    then Result := AltColor
    else Result := GetDefaultAltColor;
end;

function TDrawElementParamsEh.GetActualBorderColor: TColor;
begin
  if BorderColor <> clDefault
    then Result := BorderColor
    else Result := GetDefaultBorderColor;
end;

function TDrawElementParamsEh.GetActualFillStyle: TPropFillStyleEh;
begin
  Result := FillStyle;
end;

function TDrawElementParamsEh.GetActualHue: TColor;
begin
  if Hue <> clDefault
    then Result := Hue
    else Result := StyleServices.GetSystemColor(clSkyBlue);
end;

function TDrawElementParamsEh.GetDefaultAltColor: TColor;
begin
  Result := Color;
end;

function TDrawElementParamsEh.GetDefaultBorderColor: TColor;
begin
  Result := clNone;
end;

function TDrawElementParamsEh.GetDefaultColor: TColor;
begin
  Result := clNone;
end;

function TDrawElementParamsEh.GetDefaultFillStyle: TPropFillStyleEh;
begin
  Result := fsSolidEh;
end;

function TDrawElementParamsEh.GetDefaultHue: TColor;
begin
  Result := clNone;
end;

procedure TDrawElementParamsEh.SetFont(const Value: TFont);
begin
  FFont.Assign(Value);
end;

procedure TDrawElementParamsEh.SetFontStored(const Value: Boolean);
begin
  if FFontStored <> Value then
  begin
    FFontStored := Value;
    RefreshDefaultFont;
  end;
end;

function TDrawElementParamsEh.IsFontStored: Boolean;
begin
  Result := FontStored;
end;

procedure TDrawElementParamsEh.RefreshDefaultFont;
var
  Save: TNotifyEvent;
begin
  if FontStored then Exit;
  Save := FFont.OnChange;
  FFont.OnChange := nil;
  try
    AssignFontDefaultProps;
  finally
    FFont.OnChange := Save;
  end;
end;

procedure TDrawElementParamsEh.AssignFontDefaultProps;
begin
  FFont.Assign(DefaultFont);
end;

{ TPlannerControlTimeSpanParamsEh }

constructor TPlannerControlTimeSpanParamsEh.Create(APlanner: TPlannerControlEh);
begin
  FPlanner := APlanner;
  inherited Create;
  FDefaultColor := clNone;
  FDefaultAltColor := clNone;
  FDefaultBorderColor := clNone;
end;

destructor TPlannerControlTimeSpanParamsEh.Destroy;
begin
  inherited Destroy;
end;

function TPlannerControlTimeSpanParamsEh.GetDefaultAltColor: TColor;
begin
  Result := FDefaultAltColor;
end;

function TPlannerControlTimeSpanParamsEh.GetDefaultBorderColor: TColor;
begin
  Result := FDefaultBorderColor;
end;

function TPlannerControlTimeSpanParamsEh.GetDefaultColor: TColor;
begin
  Result := FDefaultColor;
end;

function TPlannerControlTimeSpanParamsEh.GetDefaultFillStyle: TPropFillStyleEh;
begin
  Result := fsVerticalGradientEh;
end;

function TPlannerControlTimeSpanParamsEh.GetDefaultHue: TColor;
begin
  Result := StyleServices.GetSystemColor(clSkyBlue);
end;

function TPlannerControlTimeSpanParamsEh.DefaultFont: TFont;
begin
  Result := Planner.Font;
end;

procedure TPlannerControlTimeSpanParamsEh.NotifyChanges;
begin
  Planner.LayoutChanged;
  ResetDefaultProps;
end;

procedure TPlannerControlTimeSpanParamsEh.ResetDefaultProps;
var
//  H, S, L: Single;
  DefHue: TColor;
begin
//  RGBtoHSL(ColorToRGB(GetActualHue), H, S, L);
//  DefHue := MakeColor(HSLtoRGB(H, S, 0.81), 0);

  DefHue := ChangeColorLuminance(GetActualHue, 206);

  FDefaultColor := ChangeRelativeColorLuminance(DefHue, 20);
  FDefaultAltColor := DefHue;
//  FDefaultBorderColor := MakeColor(HSLtoRGB(H, S, 0.3), 0);
  FDefaultBorderColor := ChangeColorLuminance(GetActualHue, 76);
end;

{ TEventNavBoxParamsEh }

constructor TEventNavBoxParamsEh.Create(APlanner: TPlannerControlEh);
begin
  inherited Create;
  FPlanner := APlanner;
  FVisible := True;
end;

destructor TEventNavBoxParamsEh.Destroy;
begin
  inherited Destroy;
end;

function TEventNavBoxParamsEh.GetDefaultBorderColor: TColor;
begin
  Result := Planner.GetDefaultEventNavBoxBorderColor;
end;

function TEventNavBoxParamsEh.GetDefaultColor: TColor;
begin
  Result := Planner.GetDefaultEventNavBoxColor;
end;

function TEventNavBoxParamsEh.DefaultFont: TFont;
begin
  Result := Planner.Font;
end;

procedure TEventNavBoxParamsEh.NotifyChanges;
begin
  Planner.NavBoxParamsChanges;
end;

procedure TEventNavBoxParamsEh.SetVisible(const Value: Boolean);
begin
  if FVisible <> Value then
  begin
    FVisible := Value;
    NotifyChanges;
  end;
end;

{ THoursBarAreaEh }


function THoursBarAreaEh.DefaultFont: TFont;
begin
  Result := PlannerView.Font;
end;

function THoursBarAreaEh.DefaultColor: TColor;
begin
  Result := PlannerView.Color;
end;

procedure THoursBarAreaEh.AssignFontDefaultProps;
begin
  FFont.Assign(DefaultFont);
  FFont.Size := FFont.Size * 3 div 2;
end;

function THoursBarAreaEh.DefaultVisible: Boolean;
begin
  Result := True;
end;

function THoursBarAreaEh.DefaultSize: Integer;
begin
  Result := PlannerView.DefaultHoursBarSize;
end;

{ THoursVertBarAreaEh }

function THoursVertBarAreaEh.GetWidth: Integer;
begin
  Result := Size;
end;

procedure THoursVertBarAreaEh.SetWidth(const Value: Integer);
begin
  Size := Value;
end;

{ THoursHorzBarAreaEh }

function THoursHorzBarAreaEh.GetHeight: Integer;
begin
  Result := Size;
end;

procedure THoursHorzBarAreaEh.SetHeight(const Value: Integer);
begin
  Size := Value;
end;

{ TDayNameAreaEh }

function TDayNameAreaEh.DefaultFont: TFont;
begin
  Result := PlannerView.Font;
end;

function TDayNameAreaEh.DefaultSize: Integer;
begin
  Result := PlannerView.GetDayNameAreaDefaultSize;
end;

function TDayNameAreaEh.DefaultColor: TColor;
begin
  Result := PlannerView.GetDayNameAreaDefaultColor;
end;

function TDayNameAreaEh.DefaultVisible: Boolean;
begin
  Result := PlannerView.IsDayNameAreaNeedVisible;
end;

{ TResourceCaptionAreaEh }

function TResourceCaptionAreaEh.DefaultSize: Integer;
begin
  Result := PlannerView.GetResourceCaptionAreaDefaultSize;
end;

function TResourceCaptionAreaEh.DefaultFont: TFont;
begin
  Result := PlannerView.Font;
end;

function TResourceCaptionAreaEh.DefaultVisible: Boolean;
begin
  Result := PlannerView.IsResourceCaptionNeedVisible;
end;

function TResourceCaptionAreaEh.GetVisible: Boolean;
begin
  Result := inherited Visible;
end;

procedure TResourceCaptionAreaEh.SetVisible(const Value: Boolean);
begin
  FEnhancedChanges := True;
  try
    inherited Visible := Value;
  finally
    FEnhancedChanges := False;
  end;
end;

procedure TResourceCaptionAreaEh.NotifyChanges;
begin
  if FEnhancedChanges and not (csLoading in PlannerView.ComponentState) then
    PlannerView.ResetLoadedTimeRange;
  inherited NotifyChanges;
end;

function TResourceCaptionAreaEh.DefaultColor: TColor;
begin
  Result := PlannerView.FixedColor;
end;

{ TResourceVertCaptionAreaEh }

function TResourceVertCaptionAreaEh.GetHeight: Integer;
begin
  Result := Size;
end;

procedure TResourceVertCaptionAreaEh.SetHeight(const Value: Integer);
begin
  Size := Value;
end;

{ TResourceHorzCaptionAreaEh }

function TResourceHorzCaptionAreaEh.GetWidth: Integer;
begin
  Result := Size;
end;

procedure TResourceHorzCaptionAreaEh.SetWidth(const Value: Integer);
begin
  Size := Value;
end;

{ TWeekBarAreaEh }

procedure TWeekBarAreaEh.AssignFontDefaultProps;
begin
  FFont.Assign(DefaultFont);
end;

{ TPlannerViewDrawElementEh }

constructor TPlannerViewDrawElementEh.Create(APlannerView: TCustomPlannerViewEh);
begin
  inherited Create;

  FPlannerView := APlannerView;

  FFont := TFont.Create;
  FFont.Assign(DefaultFont);
  FFont.OnChange := FontChanged;

  FColor := clDefault;
  FFontStored := False;
  FSize := 0;
  FVisible := True;
  FVisibleStored := False;
end;

destructor TPlannerViewDrawElementEh.Destroy;
begin
  FreeAndNil(FFont);
  inherited Destroy;
end;


function TPlannerViewDrawElementEh.DefaultFont: TFont;
begin
  Result := PlannerView.Font;
end;

procedure TPlannerViewDrawElementEh.FontChanged(Sender: TObject);
begin
  NotifyChanges;
  FFontStored := True;
end;

function TPlannerViewDrawElementEh.IsFontStored: Boolean;
begin
  Result := FontStored;
end;

procedure TPlannerViewDrawElementEh.RefreshDefaultFont;
var
  Save: TNotifyEvent;
begin
  if FontStored then Exit;
  Save := FFont.OnChange;
  FFont.OnChange := nil;
  try
    AssignFontDefaultProps;
  finally
    FFont.OnChange := Save;
  end;
end;

procedure TPlannerViewDrawElementEh.SetFont(const Value: TFont);
begin
  FFont.Assign(Value);
end;

procedure TPlannerViewDrawElementEh.SetFontStored(const Value: Boolean);
begin
  if FFontStored <> Value then
  begin
    FFontStored := Value;
    RefreshDefaultFont;
  end;
end;

procedure TPlannerViewDrawElementEh.AssignFontDefaultProps;
begin
  FFont.Assign(DefaultFont);
end;


function TPlannerViewDrawElementEh.GetVisible: Boolean;
begin
  if VisibleStored
    then Result := FVisible
    else Result := DefaultVisible;
end;

function TPlannerViewDrawElementEh.IsVisibleStored: Boolean;
begin
  Result := FVisibleStored;
end;

procedure TPlannerViewDrawElementEh.SetVisible(const Value: Boolean);
begin
  if VisibleStored and (Value = FVisible) then Exit;
  VisibleStored := True;
  FVisible := Value;
  NotifyChanges;
end;

procedure TPlannerViewDrawElementEh.SetVisibleStored(const Value: Boolean);
var
  OldVisible: Boolean;
begin
  OldVisible := Visible;
  if (Value = True) and (IsVisibleStored = False) then
  begin
    FVisibleStored := True;
    FVisible := DefaultVisible;
  end else if (Value = False) and (IsVisibleStored = True) then
  begin
    FVisibleStored := False;
    FVisible := DefaultVisible;
  end;
  if OldVisible <> Visible then
    NotifyChanges;
end;

function TPlannerViewDrawElementEh.DefaultVisible: Boolean;
begin
  {$IFDEF FPC}
  Result := False;
  {$ELSE}
  {$ENDIF}
  raise Exception.Create('TPlannerViewDrawElementEh.DefaultVisible');
end;


procedure TPlannerViewDrawElementEh.SetColor(const Value: TColor);
begin
  if FColor <> Value then
  begin
    FColor := Value;
    NotifyChanges;
  end;
end;

function TPlannerViewDrawElementEh.GetActualColor: TColor;
begin
  if Color <> clDefault
    then Result := Color
    else Result := DefaultColor;
end;

function TPlannerViewDrawElementEh.DefaultColor: TColor;
begin
  Result := PlannerView.Color;
end;


procedure TPlannerViewDrawElementEh.SetSize(const Value: Integer);
begin
  if FSize <> Value then
  begin
    FSize := Value;
    NotifyChanges;
  end;
end;

function TPlannerViewDrawElementEh.GetActualSize: Integer;
begin
  if Size <> 0
    then Result := Size
    else Result := DefaultSize;
end;

function TPlannerViewDrawElementEh.DefaultSize: Integer;
begin
  {$IFDEF FPC}
  Result := 0;
  {$ELSE}
  {$ENDIF}
  raise Exception.Create('TPlannerViewDrawElementEh.DefaultSize');
end;

procedure TPlannerViewDrawElementEh.NotifyChanges;
begin
  PlannerView.GridLayoutChanged;
  PlannerView.Invalidate;
end;

{ TDatesBarAreaEh }

function TDatesBarAreaEh.DefaultSize: Integer;
begin
  Result := PlannerView.GetDefaultDatesBarSize;
end;

function TDatesBarAreaEh.DefaultVisible: Boolean;
begin
  Result := PlannerView.GetDefaultDatesBarVisible;
end;

function TDatesBarAreaEh.GetPlannerView: TPlannerAxisTimelineViewEh;
begin
  Result := TPlannerAxisTimelineViewEh(inherited PlannerView);
end;

{ TDatesColAreaEh }

function TDatesColAreaEh.GetWidth: Integer;
begin
  Result := Size;
end;

procedure TDatesColAreaEh.SetWidth(const Value: Integer);
begin
  Size := Value;
end;

{ TDatesRowAreaEh }

function TDatesRowAreaEh.GetHeight: Integer;
begin
  Result := Size;
end;

procedure TDatesRowAreaEh.SetHeight(const Value: Integer);
begin
  Size := Value;
end;

{ TDayNameVertAreaEh }

function TDayNameVertAreaEh.GetActualHeight: Integer;
begin
  Result := GetActualSize;
end;

function TDayNameVertAreaEh.GetHeight: Integer;
begin
  Result := Size;
end;

procedure TDayNameVertAreaEh.SetHeight(const Value: Integer);
begin
  Size := Value;
end;

{ TDayNameHorzAreaEh }

function TDayNameHorzAreaEh.GetActualWidth: Integer;
begin
  Result := GetActualSize;
end;

function TDayNameHorzAreaEh.GetWidth: Integer;
begin
  Result := Size;
end;

procedure TDayNameHorzAreaEh.SetWidth(const Value: Integer);
begin
  Size := Value;
end;

{ TDataBarsAreaEh }

constructor TDataBarsAreaEh.Create(APlannerView: TCustomPlannerViewEh);
begin
  inherited Create;

  FPlannerView := APlannerView;
  FColor := clDefault;
  FBarSize := 0;
end;

destructor TDataBarsAreaEh.Destroy;
begin
  inherited Destroy;
end;

function TDataBarsAreaEh.DefaultColor: TColor;
begin
  Result := PlannerView.Color;
end;

function TDataBarsAreaEh.DefaultBarSize: Integer;
begin
  Result := PlannerView.GetDataBarsAreaDefaultBarSize;
end;

function TDataBarsAreaEh.GetActualColor: TColor;
begin
  if Color <> clDefault
    then Result := Color
    else Result := DefaultColor;
end;

function TDataBarsAreaEh.GetActualBarSize: Integer;
begin
  if BarSize <> 0
    then Result := BarSize
    else Result := DefaultBarSize;
end;

procedure TDataBarsAreaEh.NotifyChanges;
begin
  PlannerView.GridLayoutChanged;
  PlannerView.Invalidate;
end;

procedure TDataBarsAreaEh.SetColor(const Value: TColor);
begin
  if FColor <> Value then
  begin
    FColor := Value;
    NotifyChanges;
  end;
end;

procedure TDataBarsAreaEh.SetBarSize(const Value: Integer);
begin
  if FBarSize <> Value then
  begin
    FBarSize := Value;
    NotifyChanges;
  end;
end;

{ TDataBarsVertAreaEh }

function TDataBarsVertAreaEh.GetActualRowHeight: Integer;
begin
  Result := GetActualBarSize;
end;

function TDataBarsVertAreaEh.GetRowHeight: Integer;
begin
  Result := BarSize;
end;

procedure TDataBarsVertAreaEh.SetRowHeight(const Value: Integer);
begin
  BarSize := Value;
end;

{ TDataBarsHorzAreaEh }

function TDataBarsHorzAreaEh.GetActualColWidth: Integer;
begin
  Result := GetActualBarSize;
end;

function TDataBarsHorzAreaEh.GetColWidth: Integer;
begin
  Result := BarSize;
end;

procedure TDataBarsHorzAreaEh.SetColWidth(const Value: Integer);
begin
  BarSize := Value;
end;

{ TPlannerVertHourslineViewEh }

constructor TPlannerVertHourslineViewEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  TimeRange := hlrDayEh;
end;

destructor TPlannerVertHourslineViewEh.Destroy;
begin
  inherited Destroy;
end;

function TPlannerVertHourslineViewEh.GetTimeRange: THourslineRangeEh;
begin
  case RangeKind of
    rkDayByHoursEh: Result := hlrDayEh;
    rkWeekByHoursEh: Result := hlrWeekEh;
  else
    raise Exception.Create('TPlannerVertHourslineViewEh.RangeKind is inavlide');
  end
end;

procedure TPlannerVertHourslineViewEh.SetTimeRange(const Value: THourslineRangeEh);
const
  RangeDays: array [THourslineRangeEh] of TTimePlanRangeKind = (rkDayByHoursEh, rkWeekByHoursEh);
begin
  RangeKind := RangeDays[Value];
end;

{ TPlannerHorzHourslineViewEh }

constructor TPlannerHorzHourslineViewEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  TimeRange := hlrDayEh;
end;

destructor TPlannerHorzHourslineViewEh.Destroy;
begin
  inherited Destroy;
end;

function TPlannerHorzHourslineViewEh.GetTimeRange: THourslineRangeEh;
begin
  case RangeKind of
    rkDayByHoursEh: Result := hlrDayEh;
    rkWeekByHoursEh: Result := hlrWeekEh;
  else
    raise Exception.Create('TPlannerVertHourslineViewEh.RangeKind is inavlide');
  end
end;

procedure TPlannerHorzHourslineViewEh.SetTimeRange(
  const Value: THourslineRangeEh);
const
  RangeDays: array [THourslineRangeEh] of TTimePlanRangeKind = (rkDayByHoursEh, rkWeekByHoursEh);
begin
  RangeKind := RangeDays[Value];
end;

{$IFDEF FPC}
{$ELSE}
{ TCustomPlannerControlPrintServiceEh }

constructor TCustomPlannerControlPrintServiceEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TCustomPlannerControlPrintServiceEh.Destroy;
begin
  inherited Destroy;
end;
{$ENDIF}

initialization
end.
