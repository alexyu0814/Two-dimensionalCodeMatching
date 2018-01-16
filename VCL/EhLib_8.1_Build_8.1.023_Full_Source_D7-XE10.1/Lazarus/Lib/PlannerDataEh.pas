{*******************************************************}
{                                                       }
{                        EhLib v8.1                     }
{                                                       }
{                Planner Source Component               }
{                      Build 8.1.004                    }
{                                                       }
{   Copyright (c) 2014-2015 by Dmitry V. Bolshakov      }
{                                                       }
{*******************************************************}

{$I EhLib.Inc}

unit PlannerDataEh;

interface

uses
  Classes, Windows, Messages, Types, Graphics, Dialogs,
{$IFDEF EH_LIB_17} System.Generics.Collections, System.UITypes, {$ENDIF}
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
  ToolCtrlsEh, DateUtils, DB, ImgList, TypInfo,
  SysUtils, Variants, GraphUtil;

type
  TPlannerDataSourceEh = class;
  TPlannerDataItemEh = class;
  TPlannerResourcesEh = class;
  TPlannerItemSourceParamsEh = class;
  TItemSourceFieldsMapEh = class;
  TItemSourceFieldsMapItemEh = class;

  TPlannerDataItemEhClass = class of TPlannerDataItemEh;

  TApplyUpdateToDataStorageEhEvent = procedure (PlannerDataSource: TPlannerDataSourceEh; PlanItem: TPlannerDataItemEh; UpdateStatus: TUpdateStatus) of object;
  TCheckLoadTimePlanRecordEhEvent = procedure (PlannerDataSource: TPlannerDataSourceEh; DataSet: TDataSet; var IsLoadRecord: Boolean) of object;

  TPrepareItemsReaderEhEvent = procedure (PlannerDataSource: TPlannerDataSourceEh;
    RequriedStartDate, RequriedFinishDate, LoadedBorderDate: TDateTime;
    var PreparedReadyStartDate, PreparedFinishDate: TDateTime) of object;
  TReadItemEhEvent = procedure (PlannerDataSource: TPlannerDataSourceEh; PlanItem: TPlannerDataItemEh; var Eof: Boolean) of object;

  TPlannerLoadDataForPeriodEhEvent =  procedure (PlannerDataSource: TPlannerDataSourceEh;
  StartDate, FinishDate, LoadedBorderDate: TDateTime; var LoadedStartDate, LoadedFinishDate: TDateTime) of object;

//  TGetPlannerDataItemClassEhEvent = procedure (PlannerDataSource: TPlannerDataSourceEh;
//    var PlannerDataItemClass: TPlannerDataItemEhClass) of object;


  TEditingStatusEh = (esBrowseEh, esEditEh, esInsertEh);

{ TPlannerResourceEh }

  TPlannerResourceEh = class(TCollectionItem)
  private
    FName: string;
    FImageIndex: TImageIndex;
    FColor: TColor;
    FResourceID: Variant;
    FFaceColor: TColor;
    FDarkLineColor: TColor;
    FBrightLineColor: TColor;

    procedure SetImageIndex(const Value: TImageIndex);
    procedure SetName(const Value: string);
    procedure SetColor(const Value: TColor);
    function GetCollection: TPlannerResourcesEh;
    procedure SetResourceID(const Value: Variant);
    function GetFaceColor: TColor;
    function GetDarkLineColor: TColor;
    function GetBrightLineColor: TColor;

  protected
    function GetDisplayName: string; override;
    procedure AssignTo(Dest: TPersistent); override;
    procedure SetDisplayName(const Value: string); override;

  public
    constructor Create(ACollection: TCollection); override;

    property FaceColor: TColor read GetFaceColor;
    property DarkLineColor: TColor read GetDarkLineColor;
    property BrightLineColor: TColor read GetBrightLineColor;

    property Collection: TPlannerResourcesEh read GetCollection;
  published
    property Color: TColor read FColor write SetColor default clNone;
    property Name: string read FName write SetName;
    property ImageIndex: TImageIndex read FImageIndex write SetImageIndex default -1;
    property ResourceID: Variant read FResourceID write SetResourceID;

  end;

{ TPlannerResourcesEh }

  TPlannerResourcesEh = class(TCollection)
  private
    FPlannerSource: TPlannerDataSourceEh;
    procedure SetItem(Index: Integer; Value: TPlannerResourceEh);
    function GetItem(Index: Integer): TPlannerResourceEh;

  protected
    function GetOwner: TPersistent; override;
    procedure Update(Item: TCollectionItem); override;

  public
    constructor Create(APlannerSource: TPlannerDataSourceEh);

    function Add: TPlannerResourceEh;
    function ResourceByResourceID(ResourceID: Variant): TPlannerResourceEh;

    property Items[Index: Integer]: TPlannerResourceEh read GetItem write SetItem; default;
  end;

{ TPlannerDataItemEh }

  TPlannerDataItemEh = class(TPersistent)
  private
    FTitle: string;
    FEndTime: TDateTime;
    FStartTime: TDateTime;
    FSource: TPlannerDataSourceEh;
    FAllDay: Boolean;
    FResourceID: Variant;
    FUpdateStatus: TUpdateStatus;
    FItemID: Variant;
    FResource: TPlannerResourceEh;
    FFillColor: TColor;
    FBody: String;
    FFrameColor: TColor;
    FChangesApplying: Boolean;
    FEditingStatus: TEditingStatusEh;

    function GetDuration: TDateTime;
    procedure SetEndTime(const Value: TDateTime);
    procedure SetStartTime(const Value: TDateTime);
    procedure SetTitle(const Value: string);
    procedure SetAllDay(const Value: Boolean);
    procedure SetResourceID(const Value: Variant);
    procedure SetItemID(const Value: Variant);
    procedure SetResource(const Value: TPlannerResourceEh);
    procedure SetFillColor(const Value: TColor);
    procedure SetBody(const Value: String);

  protected
    FUpdateCount: Integer;
//    function CreatePlanItemOldValues: TTimePlanItemOldValueEh; virtual;
    procedure Change; virtual;
    procedure UpdateRefResource; virtual;
    procedure StartApplyChanges; virtual;
    procedure FinishApplyChanges; virtual;
    procedure ResolvePlanItemUpdate; virtual;
    procedure BeginInsert; virtual;

  public
    constructor Create(ASource: TPlannerDataSourceEh); virtual;
    destructor Destroy; override;

    procedure BeginEdit; virtual;
    procedure EndEdit(PostChanges: Boolean); virtual;

    function GetFrameColor: TColor; virtual;

    function InsideDay: Boolean; virtual;
    function InsideDayRange: Boolean; virtual;
    procedure Assign(Source: TPersistent); override;
    procedure AssignProperties(Source: TPlannerDataItemEh); virtual;
    procedure MergeUpdates;
    procedure ApplyUpdates;

    property Duration: TDateTime read GetDuration;
    property Source: TPlannerDataSourceEh read FSource;
    property Resource: TPlannerResourceEh read FResource write SetResource;

    property UpdateCount: Integer read FUpdateCount;
    property UpdateStatus: TUpdateStatus read FUpdateStatus;
    property EditingStatus: TEditingStatusEh read FEditingStatus;

  published
    property Title: String read FTitle write SetTitle;
    property Body: String read FBody write SetBody;
    property StartTime: TDateTime read FStartTime write SetStartTime;
    property EndTime: TDateTime read FEndTime write SetEndTime;
    property AllDay: Boolean read FAllDay write SetAllDay;
    property ResourceID: Variant read FResourceID write SetResourceID;
    property ItemID: Variant read FItemID write SetItemID;
    property FillColor: TColor read FFillColor write SetFillColor;
  end;

//  TPlannerDataItemEhClass = class of TPlannerDataItemEh;

{
  ItemSourceParams
    DataSet: TDataSet
    DataDriver: TDataDriverEh
    FieldMap
      TitleFieldName: String;
      BodyFieldName: String;
      StartTimeFieldName: String;
      EndTimeFieldName: String;
      AllDayFieldName: String;
      ResourceIDFieldName: String;
      ItemIDFieldName: String;
      FillColorFieldName: String;

  OnMapItemFields(FieldMapArr = array of PropName, PropNum, FieldName, out FieldIndex)

  LoadTimeItems();

  ResourceSourceParams
    DataSet: TDataSet
    DataDriver: TDataDriverEh
    FieldMap
      ColorFieldName: String;
      NameFieldName: String;
      ImageIndexFieldName: String;
      ResourceIDFieldName: String;

  OnMapResourceFields(FieldMapArr = array of PropName, PropNum, FieldName, out FieldIndex)

  LoadResources();

  LoadData();
}

(*
  TFieldMapArrayItemEh = record
    PropName: String;
    PropNum: Integer;
    FieldName: String;
    FieldIndex: Integer;
  end;

  TFieldMapArrayEh = array of TFieldMapArrayItemEh;

{ TItemSourceFieldMapEh }

  TItemSourceFieldMapEh = class(TPersistent)
  private
    FResourceIDFieldName: String;
    FAllDayFieldName: String;
    FItemIDFieldName: String;
    FFillColorFieldName: String;
    FStartTimeFieldName: String;
    FBodyFieldName: String;
    FTitleFieldName: String;
    FEndTimeFieldName: String;
    FItemSourceParamsEh: TPlannerItemSourceParamsEh;
    procedure SetAllDayFieldName(const Value: String);
    procedure SetBodyFieldName(const Value: String);
    procedure SetEndTimeFieldName(const Value: String);
    procedure SetFillColorFieldName(const Value: String);
    procedure SetItemIDFieldName(const Value: String);
    procedure SetResourceIDFieldName(const Value: String);
    procedure SetStartTimeFieldName(const Value: String);
    procedure SetTitleFieldName(const Value: String);
  public
    constructor Create(AItemSourceParamsEh: TPlannerItemSourceParamsEh);
    destructor Destroy; override;

    property ItemSourceParamsEh: TPlannerItemSourceParamsEh read FItemSourceParamsEh;

  published
    property ItemIDFieldName: String read FItemIDFieldName write SetItemIDFieldName;
    property TitleFieldName: String read FTitleFieldName write SetTitleFieldName;
    property BodyFieldName: String read FBodyFieldName write SetBodyFieldName;
    property StartTimeFieldName: String read FStartTimeFieldName write SetStartTimeFieldName;
    property EndTimeFieldName: String read FEndTimeFieldName write SetEndTimeFieldName;
    property AllDayFieldName: String read FAllDayFieldName write SetAllDayFieldName;
    property ResourceIDFieldName: String read FResourceIDFieldName write SetResourceIDFieldName;
    property FillColorFieldName: String read FFillColorFieldName write SetFillColorFieldName;
  end;

*)

  TItemSourceReadValueEventEh = procedure(DataSource: TPlannerDataSourceEh;
    MapItem: TItemSourceFieldsMapItemEh; const DataItem: TPlannerDataItemEh;
    var Processed: Boolean) of object;

{ TItemSourceFieldsMapItemEh }

  TItemSourceFieldsMapItemEh = class(TCollectionItem)
  private
    FPropertyName: String;
    FField: TField;
    FDataSetFieldName: String;
    FPropInfo: PPropInfo;
    FOnReadValue: TItemSourceReadValueEventEh;
    function GetCollection: TItemSourceFieldsMapEh;

  public
    constructor Create(Collection: TCollection); override;
    procedure Assign(Source: TPersistent); override;
    procedure DefaulReadValue(const DataItem: TPlannerDataItemEh); virtual;
    procedure ReadValue(const DataItem: TPlannerDataItemEh); virtual;

    property Field: TField read FField write FField;
    property PropInfo: PPropInfo read FPropInfo write FPropInfo;
    property Collection: TItemSourceFieldsMapEh read GetCollection;
  published
    property DataSetFieldName: String read FDataSetFieldName write FDataSetFieldName;
    property PropertyName: String read FPropertyName write FPropertyName;

    property OnReadValue: TItemSourceReadValueEventEh read FOnReadValue write FOnReadValue;
  end;

{ TItemSourceFieldsMapEh }

  TItemSourceFieldsMapEh = class(TCollection, IDefaultItemsCollectionEh)
  private
    FItemSourceParamsEh: TPlannerItemSourceParamsEh;
    function GetItem(Index: Integer): TItemSourceFieldsMapItemEh;
    procedure SetItem(Index: Integer; const Value: TItemSourceFieldsMapItemEh);
    function GetSourceParams: TPlannerItemSourceParamsEh;

  protected
    {IInterface}
  {$IFDEF FPC}
    function QueryInterface(constref IID: TGUID; out Obj): HResult; virtual; stdcall;
  {$ELSE}
    function QueryInterface(const IID: TGUID; out Obj): HResult; virtual; stdcall;
  {$ENDIF}
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;

    {IDefaultItemsCollectionEh}
    function CanAddDefaultItems: Boolean;
    procedure AddAllItems(DeleteExisting: Boolean);

  protected
    function GetOwner: TPersistent; override;

  public
    constructor Create(AItemSourceParamsEh: TPlannerItemSourceParamsEh; ItemClass: TCollectionItemClass);
    destructor Destroy; override;

    function Add: TItemSourceFieldsMapItemEh;

    procedure ReadDataRecordValues(DataItem: TPlannerDataItemEh);
    procedure InitItems;

    procedure BuildItems(const DataItemClass: TPlannerDataItemEhClass); virtual;
    property Item[Index: Integer]: TItemSourceFieldsMapItemEh read GetItem write SetItem; default;
    property SourceParams: TPlannerItemSourceParamsEh read GetSourceParams;
  end;

{ TPlannerItemSourceParamsEh }

  TPlannerItemSourceParamsEh = class(TPersistent)
  private
    FDataSet: TDataSet;
    FPlannerDataSource: TPlannerDataSourceEh;
    FFieldsMap: TItemSourceFieldsMapEh;
    procedure SetDataSet(const Value: TDataSet);
    procedure SetFieldMap(const Value: TItemSourceFieldsMapEh);
  protected
    function GetOwner: TPersistent; override;

  public
    constructor Create(APlannerDataSource: TPlannerDataSourceEh);
    destructor Destroy; override;

    property PlannerDataSource: TPlannerDataSourceEh read FPlannerDataSource;
  published
    property DataSet: TDataSet read FDataSet write SetDataSet;
//    property DataDriver: TDataDriverEh
    property FieldsMap: TItemSourceFieldsMapEh read FFieldsMap write SetFieldMap;
  end;

{ TPlannerDataSourceEh }

  TPlannerDataSourceEh = class(TComponent)
  private
    FChanged: Boolean;
    FItems: TList;
    FTimePlanItemClass: TPlannerDataItemEhClass;
    FUpdateCount: Integer;
    FNotificationConsumers: TInterfaceList;
    FResources: TPlannerResourcesEh;
    FOnApplyUpdateToDataStorage: TApplyUpdateToDataStorageEhEvent;
    FOnCheckLoadTimePlanRecord: TCheckLoadTimePlanRecordEhEvent;
    FItemSourceParams: TPlannerItemSourceParamsEh;
    FLoadedFinishDate: TDateTime;
    FLoadedStartDate: TDateTime;
    FOnPrepareItemsReader: TPrepareItemsReaderEhEvent;
    FOnPrepareReadItem: TReadItemEhEvent;
    FAllItemsLoaded: Boolean;
    FOnLoadDataForPeriod: TPlannerLoadDataForPeriodEhEvent;
//    FOnGetPlannerDataItemClass: TGetPlannerDataItemClassEhEvent;

    function GetCount: Integer;
    function GetItems(Index: Integer): TPlannerDataItemEh;
    function GetResources: TPlannerResourcesEh;
    procedure SetResources(const Value: TPlannerResourcesEh);
    procedure SetItemSourceParams(const Value: TPlannerItemSourceParamsEh);
    procedure SetAllItemsLoaded(const Value: Boolean);
    procedure SetLoadedFinishDateC(const Value: TDateTime);
    procedure SetLoadedStartDate(const Value: TDateTime);

  protected
//    FFieldMapArray: TFieldMapArrayEh;

    function IsLoadTimePlanRecord(ADataSet: TDataSet): Boolean; virtual;

    procedure Sort;
    procedure PlanChanged; virtual;
    procedure ResourcesChanged; virtual;
    procedure UpdateRecourdesByRecourdeID; virtual;
    procedure PlanItemChanged(PlanItem: TPlannerDataItemEh); virtual;
    procedure ResolvePlanItemUpdate(PlanItem: TPlannerDataItemEh; UpdateStatus: TUpdateStatus); virtual;
    procedure ResolvePlanItemUpdateToDataStorage(PlanItem: TPlannerDataItemEh; UpdateStatus: TUpdateStatus);
    procedure FreeItem(Item: TPlannerDataItemEh);
//    procedure ExtendLoadedRange(NewStartDate, NewFinishDate: TDateTime); virtual;
    procedure LoadDataForPeriod(StartDate, FinishDate, LoadedBorderDate: TDateTime; var LoadedStartDate, LoadedFinishDate: TDateTime); virtual;
    procedure PrepareItemsReader(StartDate, FinishDate, LoadedBorderDate: TDateTime; var ALoadedStartDate, ALoadedFinishDate: TDateTime); virtual;
    procedure ReadItem(PlanItem: TPlannerDataItemEh; var Eof: Boolean); virtual;

    procedure Notification(AComponent: TComponent; Operation: TOperation); override;

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure RegisterChanges(Value: ISimpleChangeNotificationEh);
    procedure UnRegisterChanges(Value: ISimpleChangeNotificationEh);

    function CreatePlannerItem: TPlannerDataItemEh; virtual;
    function CreateTmpPlannerItem: TPlannerDataItemEh; virtual;
    function NewItem(ATimePlanItemClass: TPlannerDataItemEhClass = nil; IsDummy: Boolean = False): TPlannerDataItemEh;
    function IndexOf(Item: TPlannerDataItemEh): Integer;

    procedure BeginUpdate;
    procedure EndUpdate;
    procedure DeleteItem(Item: TPlannerDataItemEh);
    procedure DeleteItemAt(Index: Integer);
    procedure DeleteItemNoApplyUpdates(Item: TPlannerDataItemEh);
    procedure AddItem(Item: TPlannerDataItemEh);
    procedure FetchTimePlanItem(Item: TPlannerDataItemEh);
    procedure RequestItems(StartTime, EndTime: TDateTime);
    procedure ClearItems;
//    procedure FillFieldMapArr(var AFieldMapArray: TFieldMapArrayEh);
    procedure LoadTimeItems; virtual;
    procedure EnsureDataForPeriod(AStartDate, AEndDate: TDateTime); virtual;
    procedure DefaultPrepareItemsReader(RequriedStartDate, RequriedFinishDate, LoadedBorderDate: TDateTime; var PreparedReadyStartDate, PreparedFinishDate: TDateTime); virtual;
    procedure DefaultReadItem(PlanItem: TPlannerDataItemEh; var Eof: Boolean); virtual;

    procedure StopAutoLoad;
    procedure ResetAutoLoadProcess;

    property Items[Index: Integer]: TPlannerDataItemEh read GetItems; default;
    property Count: Integer read GetCount;
    property Resources: TPlannerResourcesEh read GetResources write SetResources;
    property LoadedStartDate: TDateTime read FLoadedStartDate write SetLoadedStartDate;
    property LoadedFinishDate: TDateTime read FLoadedFinishDate write SetLoadedFinishDateC;
    property AllItemsLoaded: Boolean read FAllItemsLoaded write SetAllItemsLoaded;
    property TimePlanItemClass: TPlannerDataItemEhClass read FTimePlanItemClass write FTimePlanItemClass;

  published
    property ItemSourceParams: TPlannerItemSourceParamsEh read FItemSourceParams write SetItemSourceParams;

    property OnApplyUpdateToDataStorage: TApplyUpdateToDataStorageEhEvent read FOnApplyUpdateToDataStorage write FOnApplyUpdateToDataStorage;
    property OnCheckLoadTimePlanRecord: TCheckLoadTimePlanRecordEhEvent read FOnCheckLoadTimePlanRecord write FOnCheckLoadTimePlanRecord;
    property OnPrepareItemsReader: TPrepareItemsReaderEhEvent read FOnPrepareItemsReader write FOnPrepareItemsReader;
    property OnLoadDataForPeriod: TPlannerLoadDataForPeriodEhEvent read FOnLoadDataForPeriod write FOnLoadDataForPeriod;
    property OnReadItem: TReadItemEhEvent read FOnPrepareReadItem write FOnPrepareReadItem;
//    property OnGetPlannerDataItemClass: TGetPlannerDataItemClassEhEvent read FOnGetPlannerDataItemClass write FOnGetPlannerDataItemClass;
  end;

{ TWorkingTimeCalendarEh }

  TDayTypeEh = (dtWorkdayEh, dtFreedayEh, dtPublicHolidayEh);

  TTimeRangeEh = record
    StartTime: TDateTime;
    FinishTime: TDateTime;
  end;

  TTimeRangesEh = array of TTimeRangeEh;

  TWorkingTimeCalendarEh = class(TComponent)
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    function IsWorkday(ADate: TDateTime): Boolean; virtual;
    procedure GetWorkingTime(ADate: TDateTime; var ATimeRanges: TTimeRangesEh); virtual;
  end;

  function GlobalWorkingTimeCalendar: TWorkingTimeCalendarEh;
  function RegisterGlobalWorkingTimeCalendar(NewWorkingTimeCalendar: TWorkingTimeCalendarEh): TWorkingTimeCalendarEh;

implementation

uses
{$IFDEF EH_LIB_17}
  UIConsts,
{$ENDIF}
  EhLibConsts,
  PlannersEh;

{ TTimePlanItemEh }

constructor TPlannerDataItemEh.Create(ASource: TPlannerDataSourceEh);
begin
  inherited Create;
  FSource := ASource;
  FFillColor := clDefault;
  FFrameColor := ChangeColorLuminance(clSkyBlue, 100);
  FEditingStatus := esInsertEh;
end;

{
function TTimePlanItemEh.CreatePlanItemOldValues: TTimePlanItemOldValueEh;
begin
  Result := TTimePlanItemOldValueEh.Create(Self);
end;
}

destructor TPlannerDataItemEh.Destroy;
begin
  if Assigned(FSource) then
  begin
    FSource.FItems.Remove(Self);
    if FSource.FUpdateCount = 0 
      then FSource.PlanChanged
      else FSource.FChanged := True;
  end;
  inherited Destroy;
end;

procedure TPlannerDataItemEh.BeginEdit;
begin
  if FUpdateCount = 0 then
    FEditingStatus := esEditEh;
  Inc(FUpdateCount);
end;

procedure TPlannerDataItemEh.BeginInsert;
begin
  if FUpdateCount > 0 then
    raise Exception.Create('In the TTimePlanItemEh.BeginInsert method FUpdateCount can''t be > 0');
  if FEditingStatus <> esInsertEh then
    raise Exception.Create('In the TTimePlanItemEh.BeginInsert method FEditingStatus must be in esInsertEh Status');
//  FEditingStatus := esInsertEh;
  Inc(FUpdateCount);
end;

procedure TPlannerDataItemEh.EndEdit(PostChanges: Boolean);
begin
  if FEditingStatus = esBrowseEh then
    raise Exception.Create('The Time Plan Item is not in the Edit or Insert mode');
  Dec(FUpdateCount);
  if (Source <> nil) and (FUpdateCount = 0) then
  begin
    if FEditingStatus = esInsertEh then
    begin
      if PostChanges then
        Source.AddItem(Self)
      else
      begin
        Source.FreeItem(Self);
        Exit;
      end;
//      Change;
    end else
    begin
      Change;
    end;
    FEditingStatus := esBrowseEh;
  end;
end;

procedure TPlannerDataItemEh.Assign(Source: TPersistent);
begin
  if Source is TPlannerDataItemEh then
  begin
    BeginEdit;
    try
      AssignProperties(Source as TPlannerDataItemEh);
    except
      EndEdit(False);
    end;
    EndEdit(True);
  end;
end;

procedure TPlannerDataItemEh.AssignProperties(Source: TPlannerDataItemEh);
begin
  Title := TPlannerDataItemEh(Source).Title;
  FEndTime := TPlannerDataItemEh(Source).EndTime;
  StartTime := TPlannerDataItemEh(Source).StartTime;
  AllDay := TPlannerDataItemEh(Source).AllDay;

  Body := TPlannerDataItemEh(Source).Body;
  ResourceID := TPlannerDataItemEh(Source).ResourceID;
  ItemID := TPlannerDataItemEh(Source).ItemID;
  Resource := TPlannerDataItemEh(Source).Resource;
  FillColor := TPlannerDataItemEh(Source).FillColor;
end;

procedure TPlannerDataItemEh.Change;
begin
  if FChangesApplying then Exit;

  if UpdateStatus = usUnmodified then
    FUpdateStatus := usModified;
  if FUpdateCount = 0 then
  begin
    if UpdateStatus <> usUnmodified then
      ResolvePlanItemUpdate;
    if Source <> nil then
      Source.PlanItemChanged(Self);
  end;
end;

procedure TPlannerDataItemEh.ApplyUpdates;
begin
  if UpdateStatus = usInserted then
    ResolvePlanItemUpdate
  else if UpdateStatus = usDeleted then
    ResolvePlanItemUpdate;
  MergeUpdates;
end;

procedure TPlannerDataItemEh.ResolvePlanItemUpdate;
begin
  StartApplyChanges;
  try
    if Source <> nil then
      Source.ResolvePlanItemUpdate(Self, UpdateStatus);
  finally
    FinishApplyChanges;
  end;
end;

procedure TPlannerDataItemEh.MergeUpdates;
begin
  FUpdateStatus := usUnmodified;
end;

function TPlannerDataItemEh.GetDuration: TDateTime;
begin
  Result := FEndTime - FStartTime;
end;

function TPlannerDataItemEh.GetFrameColor: TColor;
begin
  Result := FFrameColor;
end;

function TPlannerDataItemEh.InsideDay: Boolean;
begin
  if (DateOf(EndTime) = EndTime) and (DateOf(StartTime+1) = EndTime) then
    Result := True
  else if not AllDay and (DateOf(StartTime) = DateOf(EndTime)) then
    Result := True
  else
    Result := False;
end;

function TPlannerDataItemEh.InsideDayRange: Boolean;
begin
  Result := (DateOf(StartTime) = DateOf(EndTime)) or (DateOf(StartTime+1) = EndTime);
end;

procedure TPlannerDataItemEh.SetAllDay(const Value: Boolean);
begin
  if FAllDay <> Value then
  begin
    FAllDay := Value;
    Change;
  end;
end;

procedure TPlannerDataItemEh.SetFillColor(const Value: TColor);
begin
  if FFillColor <> Value then
  begin
    FFillColor := Value;
    if FFillColor = clDefault then
      FFrameColor := ChangeColorLuminance(clSkyBlue, 100)
    else
      FFrameColor := ChangeColorLuminance(FFillColor, 100);
    Change;
  end;
end;

procedure TPlannerDataItemEh.SetItemID(const Value: Variant);
begin
  if FItemID <> Value then
  begin
    FItemID := Value;
    Change;
  end;
end;

procedure TPlannerDataItemEh.SetResourceID(const Value: Variant);
begin
  if not VarSameValue(FResourceID, Value) then
  begin
    FResourceID := Value;
    UpdateRefResource;
    Change;
  end;
end;

procedure TPlannerDataItemEh.SetResource(const Value: TPlannerResourceEh);
begin
  if FResource <> Value then
  begin
    FResource := Value;
    if FResource <> nil then
      FResourceID := FResource.ResourceID
    else
      FResourceID := Null;
    Change;
  end;
end;

procedure TPlannerDataItemEh.SetStartTime(const Value: TDateTime);
begin
  if FStartTime <> Value then
  begin
    FStartTime := NormalizeDateTime(Value);
    if FEndTime < FStartTime then
      FEndTime := FStartTime;
    Change;
  end;
end;

procedure TPlannerDataItemEh.SetEndTime(const Value: TDateTime);
begin
  if FEndTime <> Value then
  begin
    if Value < FStartTime then
      raise Exception.Create(SPlannerEndTimeBeforeStartTimeEh);
    FEndTime := NormalizeDateTime(Value);
    Change;
  end;
end;

procedure TPlannerDataItemEh.SetTitle(const Value: string);
begin
  if FTitle <> Value then
  begin
    FTitle := Value;
    Change;
  end;
end;

procedure TPlannerDataItemEh.SetBody(const Value: String);
begin
  if FBody <> Value then
  begin
    FBody := Value;
    Change;
  end;
end;

procedure TPlannerDataItemEh.UpdateRefResource;
begin
  if Source <> nil then
    FResource := Source.Resources.ResourceByResourceID(ResourceID);
end;

procedure TPlannerDataItemEh.StartApplyChanges;
begin
  if FChangesApplying then
    raise Exception.Create('TTimePlanItemEh.StartApplyChanges: ApplyChanges already started.');
  FChangesApplying := True;
end;

procedure TPlannerDataItemEh.FinishApplyChanges;
begin
  if not FChangesApplying then
    raise Exception.Create('TTimePlanItemEh.FinishApplyChanges: ApplyChanges has not yet started.');
  FChangesApplying := False;
end;

{ TPlannerDataSourceEh }

constructor TPlannerDataSourceEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FItems := TList.Create;
  FTimePlanItemClass := TPlannerDataItemEh;
  FNotificationConsumers := TInterfaceList.Create;
  FResources := TPlannerResourcesEh.Create(Self);
  FItemSourceParams := TPlannerItemSourceParamsEh.Create(Self);
  FLoadedStartDate := 0;
  FLoadedFinishDate := 0;
  FAllItemsLoaded := True;
end;

destructor TPlannerDataSourceEh.Destroy;
begin
  Destroying;
  ClearItems;
  FreeAndNil(FNotificationConsumers);
  FreeAndNil(FItems);
  FreeAndNil(FResources);
  FreeAndNil(FItemSourceParams);
  inherited Destroy;
end;

function TPlannerDataSourceEh.CreatePlannerItem: TPlannerDataItemEh;
begin
  Result := NewItem(TimePlanItemClass);
end;

function TPlannerDataSourceEh.CreateTmpPlannerItem: TPlannerDataItemEh;
begin
  Result := NewItem(TimePlanItemClass, True);
end;

function TPlannerDataSourceEh.NewItem(ATimePlanItemClass: TPlannerDataItemEhClass = nil; IsDummy: Boolean = False): TPlannerDataItemEh;
begin
  if ATimePlanItemClass = nil then
    ATimePlanItemClass := TimePlanItemClass;
  if IsDummy
    then Result := ATimePlanItemClass.Create(nil)
    else Result := ATimePlanItemClass.Create(Self);
//  Result.BeginEdit;
  Result.BeginInsert;
end;

procedure TPlannerDataSourceEh.AddItem(Item: TPlannerDataItemEh);
begin
  // Sort Item Here
  if Item.Source <> Self then
    raise Exception.Create('procedure TPlannerDataSourceEh.Add: TimePlanItem.Source is not belong to adding Source');
  if Item.FEditingStatus <> esInsertEh then
    raise Exception.Create('procedure TPlannerDataSourceEh.Add: TimePlanItem is not in esInsertEh Editing Status');
  Item.FUpdateStatus := usInserted;
  Item.ApplyUpdates;
  Item.FEditingStatus := esBrowseEh;  
  if Item.FUpdateCount > 0 then
    Dec(Item.FUpdateCount);
  FItems.Add(Item);
  PlanChanged;
end;

procedure TPlannerDataSourceEh.FreeItem(Item: TPlannerDataItemEh);
begin
  Item.Free;
end;

procedure TPlannerDataSourceEh.FetchTimePlanItem(Item: TPlannerDataItemEh);
begin
  if Item.FUpdateCount > 0 then
    Dec(Item.FUpdateCount);
  if Item.Source <> Self then
    raise Exception.Create('procedure TPlannerDataSourceEh.FetchTimePlanItem: TimePlanItem.Source is not belong to adding Source');
  if Item.FEditingStatus <> esInsertEh then
    raise Exception.Create('procedure TPlannerDataSourceEh.FetchTimePlanItem: TimePlanItem is not in esInsertEh Editing Status');
  FItems.Add(Item);
  Item.MergeUpdates;
  PlanChanged;
end;

//procedure TPlannerDataSourceEh.AddItemNoApplyToDataStorage(Item: TTimePlanItemEh);
//procedure TPlannerDataSourceEh.AddItemNoApplyUpdates(Item: TTimePlanItemEh);

//procedure TPlannerDataSourceEh.DeleteItemNoApplyToDataStorage();
//procedure TPlannerDataSourceEh.DeleteItemNoApplyUpdates();

//procedure TPlannerDataSourceEh.EndUpdateNoApplyToDataStorage();
//procedure TPlannerDataSourceEh.EndUpdateNoApplyUpdates();

procedure TPlannerDataSourceEh.BeginUpdate;
begin
  Inc(FUpdateCount);
end;

procedure TPlannerDataSourceEh.EndUpdate;
begin
  if FUpdateCount > 0 then
  begin
    Dec(FUpdateCount);
    if FChanged then
      PlanChanged;
  end;
end;

procedure TPlannerDataSourceEh.PlanChanged;
var
  i: Integer;
begin
  if csDestroying in ComponentState then Exit;
  if FUpdateCount = 0 then
  begin
    FChanged := False;
    if FNotificationConsumers <> nil then
      for I := 0 to FNotificationConsumers.Count - 1 do
        (FNotificationConsumers[I] as ISimpleChangeNotificationEh).Change(Self);
  //  if Assigned(FOnChange) then FOnChange(Self);
  end else
    FChanged := True;
end;

procedure TPlannerDataSourceEh.ClearItems;
begin
  BeginUpdate;
  try
    while Count > 0 do
      Items[Count-1].Free;
  finally
    EndUpdate;
  end;
  FLoadedStartDate := 0;
  FLoadedFinishDate := 0;
end;

procedure TPlannerDataSourceEh.DeleteItem(Item: TPlannerDataItemEh);
begin
  Item.FUpdateStatus := usDeleted;
  Item.ApplyUpdates;
  DeleteItemNoApplyUpdates(Item);
end;

procedure TPlannerDataSourceEh.DeleteItemAt(Index: Integer);
begin
  DeleteItem(Items[Index]);
end;

procedure TPlannerDataSourceEh.DeleteItemNoApplyUpdates(Item: TPlannerDataItemEh);
begin
  Item.Free;
end;

function TPlannerDataSourceEh.GetCount: Integer;
begin
  Result := FItems.Count;
end;

function TPlannerDataSourceEh.GetItems(Index: Integer): TPlannerDataItemEh;
begin
  Result := TPlannerDataItemEh(FItems[Index]);
end;

function TPlannerDataSourceEh.IndexOf(Item: TPlannerDataItemEh): Integer;
begin
  Result := FItems.IndexOf(Item);
end;

function TPlannerDataSourceEh.IsLoadTimePlanRecord(ADataSet: TDataSet): Boolean;
begin
  Result := True;
  if Assigned(OnCheckLoadTimePlanRecord) then
    OnCheckLoadTimePlanRecord(Self, ADataSet, Result);
end;

(*
procedure TPlannerDataSourceEh.FillFieldMapArr(var AFieldMapArray: TFieldMapArrayEh);

  function GetFieldIndex(FieldName: String): Integer;
  var
    Field: TField;
  begin
    Result := -1;
    if ItemSourceParams.DataSet <> nil then
    begin
      Field := ItemSourceParams.DataSet.FindField(FieldName);
      if Field <> nil then
        Result := Field.Index;
    end;
  end;

begin
  SetLength(AFieldMapArray, 8);

  AFieldMapArray[0].PropName := 'ItemIDFieldName';
  AFieldMapArray[0].PropNum := 0;
  AFieldMapArray[0].FieldName := ItemSourceParams.FieldMap.ItemIDFieldName;
  AFieldMapArray[0].FieldIndex := GetFieldIndex(ItemSourceParams.FieldMap.ItemIDFieldName);

  AFieldMapArray[1].PropName := 'TitleFieldName';
  AFieldMapArray[1].PropNum := 0;
  AFieldMapArray[1].FieldName := ItemSourceParams.FieldMap.TitleFieldName;
  AFieldMapArray[1].FieldIndex := GetFieldIndex(ItemSourceParams.FieldMap.TitleFieldName);

  AFieldMapArray[2].PropName := 'BodyFieldName';
  AFieldMapArray[2].PropNum := 1;
  AFieldMapArray[2].FieldName := ItemSourceParams.FieldMap.BodyFieldName;
  AFieldMapArray[2].FieldIndex := GetFieldIndex(ItemSourceParams.FieldMap.BodyFieldName);

  AFieldMapArray[3].PropName := 'StartTimeFieldName';
  AFieldMapArray[3].PropNum := 2;
  AFieldMapArray[3].FieldName := ItemSourceParams.FieldMap.StartTimeFieldName;
  AFieldMapArray[3].FieldIndex := GetFieldIndex(ItemSourceParams.FieldMap.StartTimeFieldName);

  AFieldMapArray[4].PropName := 'EndTimeFieldName';
  AFieldMapArray[4].PropNum := 3;
  AFieldMapArray[4].FieldName := ItemSourceParams.FieldMap.EndTimeFieldName;
  AFieldMapArray[4].FieldIndex := GetFieldIndex(ItemSourceParams.FieldMap.EndTimeFieldName);

  AFieldMapArray[5].PropName := 'AllDayFieldName';
  AFieldMapArray[5].PropNum := 4;
  AFieldMapArray[5].FieldName := ItemSourceParams.FieldMap.AllDayFieldName;
  AFieldMapArray[5].FieldIndex := GetFieldIndex(ItemSourceParams.FieldMap.AllDayFieldName);

  AFieldMapArray[6].PropName := 'ResourceIDFieldName';
  AFieldMapArray[6].PropNum := 5;
  AFieldMapArray[6].FieldName := ItemSourceParams.FieldMap.ResourceIDFieldName;
  AFieldMapArray[6].FieldIndex := GetFieldIndex(ItemSourceParams.FieldMap.ResourceIDFieldName);

  AFieldMapArray[7].PropName := 'FillColorFieldName';
  AFieldMapArray[7].PropNum := 7;
  AFieldMapArray[7].FieldName := ItemSourceParams.FieldMap.FillColorFieldName;
  AFieldMapArray[7].FieldIndex := GetFieldIndex(ItemSourceParams.FieldMap.FillColorFieldName);

end;
*)

procedure TPlannerDataSourceEh.EnsureDataForPeriod(AStartDate, AEndDate: TDateTime);
var
  NewLoadedStartDate, NewLoadedFinishDate: TDateTime;
begin
  if AllItemsLoaded then Exit;

  if (LoadedStartDate = 0) and (LoadedFinishDate = 0) then
  begin
    LoadDataForPeriod(AStartDate, AEndDate, 0, NewLoadedStartDate, NewLoadedFinishDate);
    FLoadedStartDate := NewLoadedStartDate;
    FLoadedFinishDate := NewLoadedFinishDate;
    if (LoadedStartDate = 0) and (LoadedFinishDate = 0) then
      FAllItemsLoaded := True;
  end else
  begin
    if AStartDate < LoadedStartDate then
    begin
      BeginUpdate;
      try
        LoadDataForPeriod(AStartDate, LoadedStartDate, LoadedStartDate,
          NewLoadedStartDate, NewLoadedFinishDate);
        FLoadedStartDate := AStartDate;
      finally
        EndUpdate;
      end;
    end;
    if AEndDate > LoadedFinishDate then
    begin
      BeginUpdate;
      try
        LoadDataForPeriod(LoadedFinishDate, AEndDate, LoadedFinishDate,
          NewLoadedStartDate, NewLoadedFinishDate);
        FLoadedFinishDate := AEndDate;
      finally
        EndUpdate;
      end;
    end;
  end;
end;

procedure TPlannerDataSourceEh.LoadDataForPeriod(StartDate, FinishDate,
  LoadedBorderDate: TDateTime; var LoadedStartDate, LoadedFinishDate: TDateTime);
var
  PlanItem: TPlannerDataItemEh;
  Eof: Boolean;
begin
  if Assigned(OnLoadDataForPeriod) then
    OnLoadDataForPeriod(Self, StartDate, FinishDate, LoadedBorderDate, LoadedStartDate, LoadedFinishDate)
  else
  begin
    PrepareItemsReader(StartDate, FinishDate, LoadedBorderDate, LoadedStartDate, LoadedFinishDate);

    BeginUpdate;
    try

      while True do
      begin
        PlanItem := NewItem;
        ReadItem(PlanItem, Eof);
        if Eof then
        begin
          PlanItem.Free;
          Break;
        end else
          FetchTimePlanItem(PlanItem);
      end;

    finally
  // No    FChanged := False; 
      EndUpdate;
    end;
  end;
end;

procedure TPlannerDataSourceEh.PrepareItemsReader(StartDate, FinishDate,
  LoadedBorderDate: TDateTime; var ALoadedStartDate,
  ALoadedFinishDate: TDateTime);
begin
  if Assigned(OnPrepareItemsReader) then
    OnPrepareItemsReader(Self, StartDate, FinishDate, LoadedBorderDate,
      ALoadedStartDate, ALoadedFinishDate)
  else
    DefaultPrepareItemsReader(StartDate, FinishDate, LoadedBorderDate,
      ALoadedStartDate, ALoadedFinishDate);
end;

procedure TPlannerDataSourceEh.DefaultPrepareItemsReader(
  RequriedStartDate, RequriedFinishDate, LoadedBorderDate: TDateTime;
  var PreparedReadyStartDate, PreparedFinishDate: TDateTime);
begin
  if ItemSourceParams.DataSet <> nil then
  begin
    ItemSourceParams.DataSet.First;
    ItemSourceParams.DataSet.DisableControls;
  end;
  PreparedReadyStartDate := 0;
  PreparedFinishDate := 0;
//  FillFieldMapArr(FFieldMapArray);
end;

procedure TPlannerDataSourceEh.ReadItem(PlanItem: TPlannerDataItemEh;
  var Eof: Boolean);
begin
  if Assigned(OnReadItem) then
    OnReadItem(Self, PlanItem, Eof)
  else
    DefaultReadItem(PlanItem, Eof);
end;

procedure TPlannerDataSourceEh.DefaultReadItem(PlanItem: TPlannerDataItemEh;
  var Eof: Boolean);
begin
  if (ItemSourceParams.DataSet = nil) or ItemSourceParams.DataSet.Eof then
    Eof := True;

  if Eof then
  begin
    if ItemSourceParams.DataSet <> nil then
      ItemSourceParams.DataSet.EnableControls;
    Exit;
  end;

  ItemSourceParams.DataSet.Next;
end;

procedure TPlannerDataSourceEh.LoadTimeItems;
var
  PlanItem: TPlannerDataItemEh;
//  AFieldMapArray: TFieldMapArrayEh;
begin
  if ItemSourceParams.DataSet = nil then Exit;

//  FillFieldMapArr(AFieldMapArray);

  BeginUpdate;
  try
  ItemSourceParams.DataSet.DisableControls;

  ClearItems;
  ItemSourceParams.DataSet.First;
  ItemSourceParams.FieldsMap.InitItems;

  while not ItemSourceParams.DataSet.Eof do
  begin

    if IsLoadTimePlanRecord(ItemSourceParams.DataSet) then
    begin

      PlanItem := NewItem();
      ItemSourceParams.FieldsMap.ReadDataRecordValues(PlanItem);

(*
      PlanItem.ItemID := ItemSourceParams.DataSet.Fields[AFieldMapArray[0].FieldIndex].Value;
      if AFieldMapArray[1].FieldIndex >= 0 then
        PlanItem.Title := VarToStr(ItemSourceParams.DataSet.Fields[AFieldMapArray[1].FieldIndex].Value);
      if AFieldMapArray[2].FieldIndex >= 0 then
        PlanItem.Body := VarToStr(ItemSourceParams.DataSet.Fields[AFieldMapArray[2].FieldIndex].Value);
      if AFieldMapArray[3].FieldIndex >= 0 then
        PlanItem.StartTime := ItemSourceParams.DataSet.Fields[AFieldMapArray[3].FieldIndex].AsDateTime;
      if AFieldMapArray[4].FieldIndex >= 0 then
        PlanItem.EndTime := ItemSourceParams.DataSet.Fields[AFieldMapArray[4].FieldIndex].AsDateTime;
      if AFieldMapArray[5].FieldIndex >= 0 then
        PlanItem.AllDay := ItemSourceParams.DataSet.Fields[AFieldMapArray[5].FieldIndex].AsBoolean;
      if AFieldMapArray[6].FieldIndex >= 0 then
        PlanItem.ResourceID := ItemSourceParams.DataSet.Fields[AFieldMapArray[6].FieldIndex].Value;
      if AFieldMapArray[7].FieldIndex >= 0 then
        PlanItem.FillColor := ItemSourceParams.DataSet.Fields[AFieldMapArray[7].FieldIndex].AsInteger;
*)

      FetchTimePlanItem(PlanItem);
    end;
    ItemSourceParams.DataSet.Next;
  end;

  finally
    ItemSourceParams.DataSet.EnableControls;
    EndUpdate;
  end;
end;

procedure TPlannerDataSourceEh.RequestItems(StartTime, EndTime: TDateTime);
begin

end;

procedure TPlannerDataSourceEh.PlanItemChanged(PlanItem: TPlannerDataItemEh);
begin
  PlanChanged;
end;

procedure TPlannerDataSourceEh.ResolvePlanItemUpdate(
  PlanItem: TPlannerDataItemEh; UpdateStatus: TUpdateStatus);
begin
  ResolvePlanItemUpdateToDataStorage(PlanItem, UpdateStatus);
end;

procedure TPlannerDataSourceEh.ResolvePlanItemUpdateToDataStorage(
  PlanItem: TPlannerDataItemEh; UpdateStatus: TUpdateStatus);
begin
  if Assigned(OnApplyUpdateToDataStorage) then
    OnApplyUpdateToDataStorage(Self, PlanItem, UpdateStatus);
end;

procedure TPlannerDataSourceEh.ResourcesChanged;
begin
  UpdateRecourdesByRecourdeID;
  PlanChanged;
end;

procedure TPlannerDataSourceEh.UpdateRecourdesByRecourdeID;
var
  i: Integer;
begin
  for i := 0 to Count-1 do
    Items[i].UpdateRefResource;
end;

procedure TPlannerDataSourceEh.Sort;
begin
  raise Exception.Create('TPlannerDataSourceEh.Sort: TODO');
end;

procedure TPlannerDataSourceEh.RegisterChanges(Value: ISimpleChangeNotificationEh);
begin
  if FNotificationConsumers.IndexOf(Value) < 0 then
    FNotificationConsumers.Add(Value);
end;

procedure TPlannerDataSourceEh.UnRegisterChanges(
  Value: ISimpleChangeNotificationEh);
begin
  FNotificationConsumers.Remove(Value);
end;

procedure TPlannerDataSourceEh.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) then
  begin
    if (AComponent is TDataSet) and (AComponent = ItemSourceParams.DataSet) then
      ItemSourceParams.DataSet := nil;
  end;
end;

procedure TPlannerDataSourceEh.ResetAutoLoadProcess;
begin
  ClearItems;
  FLoadedStartDate := 0;
  FLoadedFinishDate := 0;
  FAllItemsLoaded := False;
  PlanChanged;
end;

procedure TPlannerDataSourceEh.StopAutoLoad;
begin
  FAllItemsLoaded := True;
end;

{ TPlannerResourceEh }

constructor TPlannerResourceEh.Create(ACollection: TCollection);
begin
  inherited Create(ACollection);
  FColor := clDefault;
end;

procedure TPlannerResourceEh.AssignTo(Dest: TPersistent);
var
  DestRes: TPlannerResourceEh;
begin
  if Dest is TPlannerResourceEh then
  begin
    DestRes := TPlannerResourceEh(Dest);
    DestRes.FName := FName;
    DestRes.FImageIndex := FImageIndex;
    DestRes.FColor := FColor;
    DestRes.Changed(False);
  end else
    inherited AssignTo(Dest);
end;

function TPlannerResourceEh.GetCollection: TPlannerResourcesEh;
begin
  Result := TPlannerResourcesEh(inherited Collection);
end;

function TPlannerResourceEh.GetDisplayName: string;
begin
  if FName <> ''
    then Result := FName
    else Result := inherited GetDisplayName;
end;

function TPlannerResourceEh.GetFaceColor: TColor;
begin
  if FFaceColor <> clDefault
    then Result := FFaceColor
    else Result := FColor;
end;

function TPlannerResourceEh.GetBrightLineColor: TColor;
begin
  Result := FBrightLineColor;
end;

function TPlannerResourceEh.GetDarkLineColor: TColor;
begin
  Result := FDarkLineColor;
end;

procedure TPlannerResourceEh.SetDisplayName(const Value: string);
begin
  Name := Value;
end;

procedure TPlannerResourceEh.SetColor(const Value: TColor);
var
{$IFDEF EH_LIB_17}
  H, S, L: Single;
{$ELSE}
  H, S, L: Word;
{$ENDIF}

begin
  if FColor <> Value then
  begin
    FColor := Value;
    Changed(False);

    if FColor <> clDefault then
    begin
{$IFDEF EH_LIB_17}
      RGBtoHSL(ColorToRGB(FColor), H, S, L);
      FFaceColor := MakeColor(HSLtoRGB(H, 0.3, 0.85), 0);
      FDarkLineColor := MakeColor(HSLtoRGB(H, 0.3, 0.75), 0);
      FBrightLineColor := MakeColor(HSLtoRGB(H, 0.3, 0.9), 0);
{$ELSE}
      ColorRGBToHLS(ColorToRGB(FColor), H, L, S);
      FFaceColor := ColorHLSToRGB(H, Trunc(0.85 * 240), Trunc(0.3 * 240));
      FDarkLineColor := ColorHLSToRGB(H, Trunc(0.75 * 240), Trunc(0.3 * 240));
      FBrightLineColor := ColorHLSToRGB(H, Trunc(0.9 * 240), Trunc(0.3 * 240));
{$ENDIF}
    end;
  end;
end;

procedure TPlannerResourceEh.SetImageIndex(const Value: TImageIndex);
begin
  if FImageIndex <> Value then
  begin
    FImageIndex := Value;
    Changed(False);
  end;
end;

procedure TPlannerResourceEh.SetName(const Value: string);
begin
  if FName <> Value then
  begin
    FName := Value;
    Changed(False);
  end;
end;

procedure TPlannerResourceEh.SetResourceID(const Value: Variant);
begin
  if not VarEquals(FResourceID, Value) then
  begin
    FResourceID := Value;
    Changed(False);
  end;
end;

function TPlannerDataSourceEh.GetResources: TPlannerResourcesEh;
begin
  Result := FResources;
end;

procedure TPlannerDataSourceEh.SetAllItemsLoaded(const Value: Boolean);
begin
  if FAllItemsLoaded <> Value then
  begin
    FAllItemsLoaded := Value;
  end;
end;

procedure TPlannerDataSourceEh.SetItemSourceParams(
  const Value: TPlannerItemSourceParamsEh);
begin
  FItemSourceParams.Assign(Value);
end;

procedure TPlannerDataSourceEh.SetLoadedFinishDateC(const Value: TDateTime);
begin
  if FLoadedFinishDate <> Value then
  begin
    FLoadedFinishDate := Value;
    ResourcesChanged;
  end;
end;

procedure TPlannerDataSourceEh.SetLoadedStartDate(const Value: TDateTime);
begin
  if FLoadedStartDate <> Value then
  begin
    FLoadedStartDate := Value;
    ResourcesChanged;
  end;
end;

procedure TPlannerDataSourceEh.SetResources(const Value: TPlannerResourcesEh);
begin
  FResources.Assign(Value);
end;

{ TPlannerResourcesEh }

constructor TPlannerResourcesEh.Create(APlannerSource: TPlannerDataSourceEh);
begin
  inherited Create(TPlannerResourceEh);
  FPlannerSource := APlannerSource;
end;

function TPlannerResourcesEh.Add: TPlannerResourceEh;
begin
  Result := TPlannerResourceEh(inherited Add);
end;

function TPlannerResourcesEh.GetItem(Index: Integer): TPlannerResourceEh;
begin
  Result := TPlannerResourceEh(inherited GetItem(Index));
end;

function TPlannerResourcesEh.GetOwner: TPersistent;
begin
  Result := FPlannerSource;
end;

function TPlannerResourcesEh.ResourceByResourceID(ResourceID: Variant): TPlannerResourceEh;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Count-1 do
    if VarEquals(Items[i].ResourceID, ResourceID) then
    begin
      Result := Items[i];
      Break;
    end;
end;

procedure TPlannerResourcesEh.SetItem(Index: Integer;
  Value: TPlannerResourceEh);
begin
  inherited SetItem(Index, Value);
end;

procedure TPlannerResourcesEh.Update(Item: TCollectionItem);
begin
  if Assigned(FPlannerSource) then
    FPlannerSource.ResourcesChanged;
end;

{ TWorkingTimeCalendarEh }

var
  FGlobalWorkingTimeCalendar: TWorkingTimeCalendarEh;

function GlobalWorkingTimeCalendar: TWorkingTimeCalendarEh;
begin
  Result := FGlobalWorkingTimeCalendar;
end;

function RegisterGlobalWorkingTimeCalendar(NewWorkingTimeCalendar: TWorkingTimeCalendarEh): TWorkingTimeCalendarEh;
begin
  Result := FGlobalWorkingTimeCalendar;
  FGlobalWorkingTimeCalendar := NewWorkingTimeCalendar;
end;

constructor TWorkingTimeCalendarEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TWorkingTimeCalendarEh.Destroy;
begin
  inherited Destroy;
end;

procedure TWorkingTimeCalendarEh.GetWorkingTime(ADate: TDateTime;
  var ATimeRanges: TTimeRangesEh);
begin
  SetLength(ATimeRanges, 1);
  ATimeRanges[0].StartTime := EncodeTime(8,0,0,0);
  ATimeRanges[0].FinishTime := EncodeTime(17,0,0,0);
end;

function TWorkingTimeCalendarEh.IsWorkday(ADate: TDateTime): Boolean;
begin
  Result := DayOfWeek(ADate) in [2,3,4,5,6];
end;

procedure InitUnit;
begin
  RegisterGlobalWorkingTimeCalendar(TWorkingTimeCalendarEh.Create(nil));
end;

procedure FinalizeUnit;
begin
//  ShowMessage('FinalizeUnit PlannerDataEh');
  FreeAndNil(FGlobalWorkingTimeCalendar);
//  ShowMessage('FinalizeUnit PlannerDataEh end');
end;

(*
{ TItemSourceFieldMapEh }

constructor TItemSourceFieldMapEh.Create(
  AItemSourceParamsEh: TPlannerItemSourceParamsEh);
begin
  inherited Create;
  FItemSourceParamsEh := AItemSourceParamsEh;
end;

destructor TItemSourceFieldMapEh.Destroy;
begin
  inherited Destroy;
end;

procedure TItemSourceFieldMapEh.SetAllDayFieldName(const Value: String);
begin
  FAllDayFieldName := Value;
end;

procedure TItemSourceFieldMapEh.SetBodyFieldName(const Value: String);
begin
  FBodyFieldName := Value;
end;

procedure TItemSourceFieldMapEh.SetEndTimeFieldName(const Value: String);
begin
  FEndTimeFieldName := Value;
end;

procedure TItemSourceFieldMapEh.SetFillColorFieldName(const Value: String);
begin
  FFillColorFieldName := Value;
end;

procedure TItemSourceFieldMapEh.SetItemIDFieldName(const Value: String);
begin
  FItemIDFieldName := Value;
end;

procedure TItemSourceFieldMapEh.SetResourceIDFieldName(const Value: String);
begin
  FResourceIDFieldName := Value;
end;

procedure TItemSourceFieldMapEh.SetStartTimeFieldName(const Value: String);
begin
  FStartTimeFieldName := Value;
end;

procedure TItemSourceFieldMapEh.SetTitleFieldName(const Value: String);
begin
  FTitleFieldName := Value;
end;

*)

{ TPlannerItemSourceParamsEh }

constructor TPlannerItemSourceParamsEh.Create(APlannerDataSource: TPlannerDataSourceEh);
begin
  inherited Create;
  FPlannerDataSource := APlannerDataSource;
  FFieldsMap := TItemSourceFieldsMapEh.Create(Self, TItemSourceFieldsMapItemEh);
end;

destructor TPlannerItemSourceParamsEh.Destroy;
begin
  FreeAndNil(FFieldsMap);
  inherited Destroy;
end;

function TPlannerItemSourceParamsEh.GetOwner: TPersistent;
begin
  Result := PlannerDataSource;
end;

procedure TPlannerItemSourceParamsEh.SetDataSet(const Value: TDataSet);
begin
  if FDataSet <> Value then
  begin
    FDataSet := Value;
    if FDataSet <> nil then
      FDataSet.FreeNotification(PlannerDataSource);
  end;
end;

procedure TPlannerItemSourceParamsEh.SetFieldMap(const Value: TItemSourceFieldsMapEh);
begin
  FFieldsMap.Assign(Value);
end;

{ TItemSourceFieldsMapEh }

function TItemSourceFieldsMapEh.Add: TItemSourceFieldsMapItemEh;
begin
  Result := TItemSourceFieldsMapItemEh(inherited Add);
end;

procedure TItemSourceFieldsMapEh.BuildItems(
  const DataItemClass: TPlannerDataItemEhClass);
var
  PropList: TPropListArray;
  i: Integer;
  AItem: TItemSourceFieldsMapItemEh;
begin
  PropList := GetPropListAsArray(DataItemClass.ClassInfo, tkProperties);
  for i := 0 to Length(PropList)-1 do
  begin
    AItem := Add;
    AItem.PropertyName := String(PropList[i].Name);
  end;
end;

constructor TItemSourceFieldsMapEh.Create(
  AItemSourceParamsEh: TPlannerItemSourceParamsEh;
  ItemClass: TCollectionItemClass);
begin
  inherited Create(ItemClass);
  FItemSourceParamsEh := AItemSourceParamsEh;
end;

destructor TItemSourceFieldsMapEh.Destroy;
begin
  inherited Destroy;
end;

procedure TItemSourceFieldsMapEh.InitItems;
var
  i: Integer;
begin
  for i := 0 to Count-1 do
  begin
    Item[i].FField := nil;
    if SourceParams.DataSet <> nil then
      Item[i].FField := SourceParams.DataSet.FindField(Item[i].DataSetFieldName);

    Item[i].FPropInfo := nil;
    if Item[i].PropertyName <> '' then
      Item[i].FPropInfo := GetPropInfo(
        SourceParams.PlannerDataSource.TimePlanItemClass.ClassInfo,
        Item[i].PropertyName);
  end;
end;

procedure TItemSourceFieldsMapEh.ReadDataRecordValues(DataItem: TPlannerDataItemEh);
var
  i: Integer;
begin
  for i := 0 to Count-1 do
    Item[i].ReadValue(DataItem);
end;

function TItemSourceFieldsMapEh.GetSourceParams: TPlannerItemSourceParamsEh;
begin
  Result := FItemSourceParamsEh;
end;

function TItemSourceFieldsMapEh.GetItem(
  Index: Integer): TItemSourceFieldsMapItemEh;
begin
  Result := TItemSourceFieldsMapItemEh(inherited Items[Index]);
end;

function TItemSourceFieldsMapEh.GetOwner: TPersistent;
begin
  Result := SourceParams;
end;

procedure TItemSourceFieldsMapEh.SetItem(Index: Integer;
  const Value: TItemSourceFieldsMapItemEh);
begin
  inherited Items[Index] := Value;
end;

function TItemSourceFieldsMapEh.CanAddDefaultItems: Boolean;
begin
  Result := True;
end;

procedure TItemSourceFieldsMapEh.AddAllItems(DeleteExisting: Boolean);
begin
  with Add do
  begin
    PropertyName := 'Title';
    if (SourceParams.DataSet <> nil) and (SourceParams.DataSet.FindField('Title') <> nil) then
      DataSetFieldName := 'Title';
  end;
  with Add do
  begin
    PropertyName := 'Body';
    if (SourceParams.DataSet <> nil) and (SourceParams.DataSet.FindField('Body') <> nil) then
      DataSetFieldName := 'Body';
  end;
  with Add do
  begin
    PropertyName := 'StartTime';
    if (SourceParams.DataSet <> nil) and (SourceParams.DataSet.FindField('StartTime') <> nil) then
      DataSetFieldName := 'StartTime';
  end;
  with Add do
  begin
    PropertyName := 'EndTime';
    if (SourceParams.DataSet <> nil) and (SourceParams.DataSet.FindField('EndTime') <> nil) then
      DataSetFieldName := 'EndTime';
  end;
  with Add do
  begin
    PropertyName := 'AllDay';
    if (SourceParams.DataSet <> nil) and (SourceParams.DataSet.FindField('AllDay') <> nil) then
      DataSetFieldName := 'AllDay';
  end;
  with Add do
  begin
    PropertyName := 'ResourceID';
    if (SourceParams.DataSet <> nil) and (SourceParams.DataSet.FindField('ResourceID') <> nil) then
      DataSetFieldName := 'ResourceID';
  end;
  with Add do
  begin
    PropertyName := 'ItemID';
    if (SourceParams.DataSet <> nil) and (SourceParams.DataSet.FindField('ItemID') <> nil) then
      DataSetFieldName := 'ItemID';
  end;
  with Add do
  begin
    PropertyName := 'FillColor';
    if (SourceParams.DataSet <> nil) and (SourceParams.DataSet.FindField('FillColor') <> nil) then
      DataSetFieldName := 'FillColor';
  end;
end;

{$IFDEF FPC}
function TItemSourceFieldsMapEh.QueryInterface(constref IID: TGUID; 
  out Obj): HResult;
{$ELSE}
function TItemSourceFieldsMapEh.QueryInterface(const IID: TGUID; 
  out Obj): HResult;
{$ENDIF}
begin
  if GetInterface(IID, Obj)
    then Result := 0
    else Result := E_NOINTERFACE;
end;

function TItemSourceFieldsMapEh._AddRef: Integer;
begin
  Result := -1;
end;

function TItemSourceFieldsMapEh._Release: Integer;
begin
  Result := -1;
end;

{ TItemSourceFieldsMapItemEh }

constructor TItemSourceFieldsMapItemEh.Create(Collection: TCollection);
begin
  inherited Create(Collection);
end;

function TItemSourceFieldsMapItemEh.GetCollection: TItemSourceFieldsMapEh;
begin
  Result := TItemSourceFieldsMapEh(inherited Collection);
end;

procedure TItemSourceFieldsMapItemEh.DefaulReadValue(
  const DataItem: TPlannerDataItemEh);
var
  TypeKind: TTypeKind;
begin
  if (PropInfo <> nil) and (Field <> nil) and not Field.IsNull then
  begin
    TypeKind := PropType_getKind(PropInfo_getPropType(PropInfo));
    case TypeKind of
      tkInteger:
        SetOrdProp(DataItem, PropInfo, Field.AsInteger);
      tkChar:
        if Length(Field.AsString) > 0 then
          SetOrdProp(DataItem, PropInfo, Ord(Field.AsString[1]));
      tkEnumeration:
        SetOrdProp(DataItem, PropInfo, Field.AsInteger);
      tkFloat:
        SetFloatProp(DataItem, PropInfo, Field.AsFloat);
      tkString, tkLString:
        SetStrProp(DataItem, PropInfo, Field.AsString);
      tkWString:
        SetWideStrProp(DataItem, PropInfo, WideString(Field.AsString));
      tkSet:
        SetOrdProp(DataItem, PropInfo, Field.AsInteger);
      tkVariant:
        SetVariantProp(DataItem, PropInfo, Field.AsVariant);
{$IFDEF EH_LIB_13}
      tkInt64:
        SetInt64Prop(DataItem, PropInfo, Field.AsLargeInt);
{$ENDIF}

      tkClass:
        raise Exception.Create('TItemSourceFieldsMapItemEh.ReadValue: tkClass is not supported');
      tkMethod:
        raise Exception.Create('TItemSourceFieldsMapItemEh.ReadValue: tkMethod is not supported');
{$IFDEF EH_LIB_6}
      tkInterface:
        raise Exception.Create('TItemSourceFieldsMapItemEh.ReadValue: tkInterface is not supported');
{$ENDIF}
{$IFDEF EH_LIB_12}
      tkUString:
        SetWideStrProp(DataItem, PropInfo, Field.AsString);
{$ENDIF}
    end;
  end;
end;

procedure TItemSourceFieldsMapItemEh.ReadValue(const DataItem: TPlannerDataItemEh);
var
  Processed: Boolean;
begin
  Processed := False;
  if Assigned(OnReadValue) then
    OnReadValue(Collection.SourceParams.PlannerDataSource, Self, DataItem, Processed);
  if not Processed then
    DefaulReadValue(DataItem);
end;

procedure TItemSourceFieldsMapItemEh.Assign(Source: TPersistent);
begin
  if Source is TItemSourceFieldsMapItemEh then
  begin
    DataSetFieldName := TItemSourceFieldsMapItemEh(Source).DataSetFieldName;
    PropertyName := TItemSourceFieldsMapItemEh(Source).PropertyName;
  end;
end;

initialization
  InitUnit;
finalization
  FinalizeUnit;
end.
