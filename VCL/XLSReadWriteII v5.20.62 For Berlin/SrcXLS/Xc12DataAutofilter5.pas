unit Xc12DataAutofilter5;

interface

uses Classes, SysUtils, Contnrs,
     xpgPUtils,
     Xc12Utils5, Xc12DataStylesheet5,
     XLSUtils5, XLSClassFactory5;

type TXc12SortBy = (x12sbValue,x12sbCellColor,x12sbFontColor,x12sbIcon);
type TXc12IconSetType = (x12ist3Arrows,x12ist3ArrowsGray,x12ist3Flags,x12ist3TrafficLights1,x12ist3TrafficLights2,x12ist3Signs,x12ist3Symbols,x12ist3Symbols2,x12ist4Arrows,x12ist4ArrowsGray,x12ist4RedToBlack,x12ist4Rating,x12ist4TrafficLights,x12ist5Arrows,x12ist5ArrowsGray,x12ist5Rating,x12ist5Quarters);
type TXc12DynamicFilterType = (x12dftNull,x12dftAboveAverage,x12dftBelowAverage,x12dftTomorrow,x12dftToday,x12dftYesterday,x12dftNextWeek,x12dftThisWeek,x12dftLastWeek,x12dftNextMonth,x12dftThisMonth,x12dftLastMonth,x12dftNextQuarter,x12dftThisQuarter,x12dftLastQuarter,x12dftNextYear,x12dftThisYear,x12dftLastYear,x12dftYearToDate,x12dftQ1,x12dftQ2,x12dftQ3,x12dftQ4,x12dftM1,x12dftM2,x12dftM3,x12dftM4,x12dftM5,x12dftM6,x12dftM7,x12dftM8,x12dftM9,x12dftM10,x12dftM11,x12dftM12);
type TXc12FilterOperator = (x12foEqual,x12foLessThan,x12foLessThanOrEqual,x12foNotEqual,x12foGreaterThanOrEqual,x12foGreaterThan);
type TXc12CalendarType = (x12ctNone,x12ctGregorian,x12ctGregorianUs,x12ctJapan,x12ctTaiwan,x12ctKorea,x12ctHijri,x12ctThai,x12ctHebrew,x12ctGregorianMeFrench,x12ctGregorianArabic,x12ctGregorianXlitEnglish);
type TXc12DateTimeGrouping = (x12dtgYear,x12dtgMonth,x12dtgDay,x12dtgHour,x12dtgMinute,x12dtgSecond);
type TXc12SortMethod = (x12smStroke,x12smPinYin,x12smNone);

type TXc12IconFilter = class(TObject)
protected
     FIconSet: TXc12IconSetType;     // Required
     FIconId: integer;
public
     constructor Create;

     procedure Clear;

     property IconSet: TXc12IconSetType read FIconSet write FIconSet;
     property IconId: integer read FIconId write FIconId;
     end;

type TXc12ColorFilter = class(TObject)
protected
     FDXF: TXc12DXF;
     FCellColor: boolean;
public
     constructor Create;

     procedure Clear;

     property DXF: TXc12DXF read FDXF write FDXF;
     property CellColor: boolean read FCellColor write FCellColor;
     end;

type TXc12DynamicFilter = class(TObject)
protected
     FType_: TXc12DynamicFilterType;
     FVal: double;
     FMaxVal: double;
public
     constructor Create;

     procedure Clear;

     property Type_: TXc12DynamicFilterType read FType_ write FType_;
     property Val: double read FVal write FVal;
     property MaxVal: double read FMaxVal write FMaxVal;
     end;

type TXc12CustomFilter = class(TObject)
protected
     FOperator: TXc12FilterOperator;
     FVal: AxUCString;
public
     constructor Create;

     procedure Clear;

     property Operator_: TXc12FilterOperator read FOperator write FOperator;
     // Val = '' when filter is unused.
     property Val: AxUCString read FVal write FVal;
     end;

type TXc12CustomFilters = class(TObject)
protected
     FAssigned: boolean;
     FFilter1: TXc12CustomFilter;
     FFilter2: TXc12CustomFilter;
     FAnd: boolean;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property Assigned: boolean read FAssigned write FAssigned;
     property Filter1: TXc12CustomFilter read FFilter1 write FFilter1;
     property Filter2: TXc12CustomFilter read FFilter2 write FFilter2;
     property And_: boolean read FAnd write FAnd;
     end;

type TXc12DateGroupItem = class(TObject)
protected
     FYear: integer;                            // Required
     FMonth: integer;
     FDay: integer;
     FHour: integer;
     FMinute: integer;
     FSecond: integer;
     FDateTimeGrouping: TXc12DateTimeGrouping;  // Required
public
     constructor Create;

     procedure Clear;

     property Year: integer read FYear write FYear;
     property Month: integer read FMonth write FMonth;
     property Day: integer read FDay write FDay;
     property Hour: integer read FHour write FHour;
     property Minute: integer read FMinute write FMinute;
     property Second: integer read FSecond write FSecond;
     property DateTimeGrouping: TXc12DateTimeGrouping read FDateTimeGrouping write FDateTimeGrouping;
     end;

type TXc12DateGroupItems = class(TObjectList)
private
     function GetItems(Index: integer): TXc12DateGroupItem;
protected
public
     constructor Create;

     function Add: TXc12DateGroupItem;

     property Items[Index: integer]: TXc12DateGroupItem read GetItems; default;
     end;


type TXc12Filters = class(TObject)
protected
     FBlank: boolean;
     FCalendarType: TXc12CalendarType;
     FFilter: TStringList;
     FDateGroupItems: TXc12DateGroupItems;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property Blank: boolean read FBlank write FBlank;
     property CalendarType: TXc12CalendarType read FCalendarType write FCalendarType;
     property Filter: TStringList read FFilter;
     property DateGroupItems: TXc12DateGroupItems read FDateGroupItems write FDateGroupItems;
     end;

type TXc12Top10 = class(TObject)
protected
     FAssigned: boolean;
     FTop: boolean;
     FPercent: boolean;
     FVal: double;         // Required
     FFilterVal: double;
public
     constructor Create;

     procedure Clear;

     property Assigned: boolean read FAssigned write FAssigned;
     property Top: boolean read FTop write FTop;
     property Percent: boolean read FPercent write FPercent;
     property Val: double read FVal write FVal;
     property FilterVal: double read FFilterVal write FFilterVal;
     end;

type TXc12FilterColumn = class(TObject)
protected
     FColId: integer;
     FHiddenButton: boolean;
     FShowButton: boolean;
     FFilters: TXc12Filters;
     FTop10: TXc12Top10;
     FCustomFilters: TXc12CustomFilters;
     FDynamicFilter: TXc12DynamicFilter;
     FColorFilter: TXc12ColorFilter;
     FIconFilter: TXc12IconFilter;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property ColId: integer read FColId write FColId;
     property HiddenButton: boolean read FHiddenButton write FHiddenButton;
     property ShowButton: boolean read FShowButton write FShowButton;
     property Filters: TXc12Filters read FFilters;
     property Top10: TXc12Top10 read FTop10;
     property CustomFilters: TXc12CustomFilters read FCustomFilters;
     property DynamicFilter: TXc12DynamicFilter read FDynamicFilter;
     property ColorFilter: TXc12ColorFilter read FColorFilter;
     property IconFilter: TXc12IconFilter read FIconFilter;
     end;

type TXc12FilterColumns = class(TObjectList)
private
     function GetItems(Index: integer): TXc12FilterColumn;
protected
     FClassFactory: TXLSClassFactory;
public
     constructor Create(AClassFactory: TXLSClassFactory);

     function Add: TXc12FilterColumn;

     property Items[Index: integer]: TXc12FilterColumn read GetItems; default;
     end;

type TXc12SortCondition = class(TObject)
protected
     FDescending: boolean;
     FSortBy: TXc12SortBy;
     FRef: TXLSCellArea;           // Required
     FCustomList: AxUCString;
     FDxfId: integer;
     FIconSet: TXc12IconSetType;
     FIconId: integer;
public
     constructor Create;

     procedure Clear;

     property Descending: boolean read FDescending write FDescending;
     property SortBy: TXc12SortBy read FSortBy write FSortBy;
     property Ref: TXLSCellArea read FRef write FRef;
     property CustomList: AxUCString read FCustomList write FCustomList;
     property DxfId: integer read FDxfId write FDxfId;
     property IconSet: TXc12IconSetType read FIconSet write FIconSet;
     property IconId: integer read FIconId write FIconId;
     end;

type TXc12SortConditions = class(TObjectList)
private
     function GetItems(Index: integer): TXc12SortCondition;
protected
public
     constructor Create;

     function Add: TXc12SortCondition;

     property Items[Index: integer]: TXc12SortCondition read GetItems; default;
     end;

type TXc12SortState = class(TObject)
protected
     FColumnSort: boolean;
     FCaseSensitive: boolean;
     FSortMethod: TXc12SortMethod;
     FRef: TXLSCellArea;       // Required
     FSortConditions: TXc12SortConditions;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property ColumnSort: boolean read FColumnSort write FColumnSort;
     property CaseSensitive: boolean read FCaseSensitive write FCaseSensitive;
     property SortMethod: TXc12SortMethod read FSortMethod write FSortMethod;
     property Ref: TXLSCellArea read FRef write FRef;
     property SortConditions: TXc12SortConditions read FSortConditions;
     end;

type TXc12AutoFilter = class(TObject)
protected
     FClassFactory: TXLSClassFactory;

     FFilterColumns: TXc12FilterColumns;
     FSortState: TXc12SortState;
     FRef: TXLSCellArea;
public
     constructor Create(AClassFactory: TXLSClassFactory);
     destructor Destroy; override;

     procedure Clear;
     function  Active: boolean;

     property FilterColumns: TXc12FilterColumns read FFilterColumns;
     property SortState: TXc12SortState read FSortState;
     property Ref: TXLSCellArea read FRef write FRef;
     end;

implementation

{ TXc12IconFilter }

procedure TXc12IconFilter.Clear;
begin
  FIconSet := x12ist3Arrows;
  FIconId := -1;
end;

constructor TXc12IconFilter.Create;
begin
  Clear;
end;

{ TXc12ColorFilter }

procedure TXc12ColorFilter.Clear;
begin
  FDXF := Nil;
  FCellColor := True;
end;

constructor TXc12ColorFilter.Create;
begin
  Clear;
end;

{ TXc12DynamicFilter }

procedure TXc12DynamicFilter.Clear;
begin
  FType_ := TXc12DynamicFilterType(XPG_UNKNOWN_ENUM);
  FVal := 0;
  FMaxVal := 0;
end;

constructor TXc12DynamicFilter.Create;
begin
  Clear;
end;

{ TXc12CustomFilter }

procedure TXc12CustomFilter.Clear;
begin
  FOperator := x12foEqual;
  FVal := '';
end;

constructor TXc12CustomFilter.Create;
begin
  Clear;
end;

{ TXc12CustomFilters }

procedure TXc12CustomFilters.Clear;
begin
  FAssigned := False;
  FFilter1.Clear;
  FFilter2.Clear;
  FAnd := False;
end;

constructor TXc12CustomFilters.Create;
begin
  FFilter1 := TXc12CustomFilter.Create;
  FFilter2 := TXc12CustomFilter.Create;
end;

destructor TXc12CustomFilters.Destroy;
begin
  FFilter1.Free;
  FFilter2.Free;
  inherited;
end;

{ TXc12Top10 }

procedure TXc12Top10.Clear;
begin
  FAssigned := False;
  FTop := True;
  FPercent := False;
  FVal := 0;
  FFilterVal := 0;
end;

constructor TXc12Top10.Create;
begin
  Clear;
end;

{ TXc12DateGroupItem }

procedure TXc12DateGroupItem.Clear;
begin
  FYear := 0;
  FMonth := 0;
  FDay := 0;
  FHour := 0;
  FMinute := 0;
  FSecond := 0;
  FDateTimeGrouping := x12dtgYear;
end;

constructor TXc12DateGroupItem.Create;
begin
  Clear;
end;

{ TXc12DateGroupItems }

function TXc12DateGroupItems.Add: TXc12DateGroupItem;
begin
  Result := TXc12DateGroupItem.Create;
  inherited Add(Result);
end;

constructor TXc12DateGroupItems.Create;
begin
  inherited Create;
end;

function TXc12DateGroupItems.GetItems(Index: integer): TXc12DateGroupItem;
begin
  Result := TXc12DateGroupItem(inherited Items[Index]);
end;

{ TXc12Filters }

procedure TXc12Filters.Clear;
begin
  FBlank := False;
  FCalendarType := x12ctNone;
  FFilter.Clear;
  FDateGroupItems.Clear;
end;

constructor TXc12Filters.Create;
begin
  FFilter := TStringList.Create;
  FDateGroupItems := TXc12DateGroupItems.Create;
  Clear;
end;

destructor TXc12Filters.Destroy;
begin
  FFilter.Free;
  FDateGroupItems.Free;
  inherited;
end;

{ TXc12FilterColumn }

procedure TXc12FilterColumn.Clear;
begin
  FColId := 0;
  FHiddenButton := False;
  FShowButton := True;
  FFilters.Clear;
  FTop10.Clear;
  FCustomFilters.Clear;
  FDynamicFilter.Clear;
  FColorFilter.Clear;
  FIconFilter.Clear;
end;

constructor TXc12FilterColumn.Create;
begin
  FFilters := TXc12Filters.Create;
  FTop10 := TXc12Top10.Create;
  FCustomFilters := TXc12CustomFilters.Create;
  FDynamicFilter := TXc12DynamicFilter.Create;
  FColorFilter := TXc12ColorFilter.Create;
  FIconFilter := TXc12IconFilter.Create;
end;

destructor TXc12FilterColumn.Destroy;
begin
  FFilters.Free;
  FTop10.Free;
  FCustomFilters.Free;
  FDynamicFilter.Free;
  FColorFilter.Free;
  FIconFilter.Free;
  inherited;
end;

{ TXc12FilterColumns }

function TXc12FilterColumns.Add: TXc12FilterColumn;
begin
  Result := TXc12FilterColumn(FClassFactory.CreateAClass(xcftAutofilterColumn));
  inherited Add(Result);
end;

constructor TXc12FilterColumns.Create(AClassFactory: TXLSClassFactory);
begin
  inherited Create;

  FClassFactory := AClassFactory;
end;

function TXc12FilterColumns.GetItems(Index: integer): TXc12FilterColumn;
begin
  Result := TXc12FilterColumn(inherited Items[Index]);
end;

{ TXc12SortCondition }

procedure TXc12SortCondition.Clear;
begin
  FDescending := False;
  FSortBy := x12sbValue;
  SetCellArea(FRef,0,0);
  FCustomList := '';
  FDxfId := -1;
  FIconSet := x12ist3Arrows;
  FIconId := 0;
end;

constructor TXc12SortCondition.Create;
begin
  Clear;
end;

{ TXc12SortConditions }

function TXc12SortConditions.Add: TXc12SortCondition;
begin
  Result := TXc12SortCondition.Create;
  inherited Add(Result);
end;

constructor TXc12SortConditions.Create;
begin
  inherited Create;
end;

function TXc12SortConditions.GetItems(Index: integer): TXc12SortCondition;
begin
  Result := TXc12SortCondition(inherited Items[Index]);
end;

{ TXc12SortState }

procedure TXc12SortState.Clear;
begin
  FColumnSort := False;
  FCaseSensitive := False;
  FSortMethod := x12smNone;
  ClearCellArea(FRef);
  FSortConditions.Clear;
end;

constructor TXc12SortState.Create;
begin
  FSortConditions := TXc12SortConditions.Create;
end;

destructor TXc12SortState.Destroy;
begin
  FSortConditions.Free;
  inherited;
end;

{ TXc12AutoFilter }

function TXc12AutoFilter.Active: boolean;
begin
  Result := AreaIsAssigned(FRef);
end;

procedure TXc12AutoFilter.Clear;
begin
  FFilterColumns.Clear;
  FSortState.Clear;
  SetCellArea(FRef,-1,-1);
end;

constructor TXc12AutoFilter.Create(AClassFactory: TXLSClassFactory);
begin
  FClassFactory := AClassFactory;

  FFilterColumns := TXc12FilterColumns.Create(FClassFactory);
  FSortState := TXc12SortState.Create;
  Clear;
end;

destructor TXc12AutoFilter.Destroy;
begin
  FFilterColumns.Free;
  FSortState.Free;
  inherited;
end;

end.
