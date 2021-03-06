{*******************************************************}
{                                                       }
{                       EhLib v8.1                      }
{             Planner Design-Time Editors               }
{                      Build 8.1.03                     }
{                                                       }
{    Copyright (c) 2014-2015 by Dmitry V. Bolshakov     }
{                                                       }
{*******************************************************}

{$I EhLib.Inc}

unit PlannerDesignEditorsEh;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, Dialogs, ComCtrls, ToolWin,
{$IFDEF EH_LIB_17} System.UITypes, System.Types, {$ENDIF}
  DBGridEhGrouping, ToolCtrlsEh, DBGridEhToolCtrls, DynVarsEh, MemTableDataEh,
  ObjectInspectorEh,
{$IFDEF DESIGNTIME}
  {$IFDEF FPC}
  PropEdits, ComponentEditors,
  {$ELSE}
  DesignIntf, DesignEditors, VCLEditors, ToolsAPI, DesignWindows,
  ComponentChildrenDesignEditorsEh, ComponentChildrenDesignSelectClassDialogEh,
  DBConsts,
  {$ENDIF}
{$ENDIF}
  DB, MemTableEh,
  {$IFDEF FPC}
  EhLibLCL, DBGridsEh,
  {$ELSE}
  EhLibVCL, DBGridEh, ColnEdit, DefaultItemsCollectionEditorsEh,
  {$ENDIF}
  GridsEh, DBAxisGridsEh,
  EhLibDesignAbout,
  PlannersEh, PlannerDataEh, PlannerCalendarPickerEh;

{$IFDEF FPC}
{$ELSE}
type

{ TPlannersDesignServiceEh }

  TPlannersDesignServiceEh = class(TCompChildrenDesignServiceEh)
    class procedure GetChildClasses(ClassList: TList); override;
    class function CreateChild(MasterComponent: TComponent; ChildClass: TComponentClass): TComponent; override;
    class function GetFormCaption(MasterComponent: TComponent): String; override;
    class procedure MoveChildUp(MasterComponent: TComponent; ChildComponent: TComponent); override;
    class procedure MoveChildDown(MasterComponent: TComponent; ChildComponent: TComponent); override;
  end;
{$ENDIF}

procedure Register;

implementation

{$R PlannersEh.dcr}

var
  ImpExpListEditorsList: TList = nil;

{$IFDEF DESIGNTIME}

type

  TPlannerEhEditor = class(TComponentEditor)
    function  GetVerbCount: Integer; override;
    function  GetVerb(Index: Integer): string; override;
    procedure ExecuteVerb(Index: Integer); override;
    procedure NextPlannerGrid;
    procedure PriorPlannerGrid;
  end;

function TPlannerEhEditor.GetVerbCount: Integer;
var
  PlView: TPlannerControlEh;
begin
  PlView := TPlannerControlEh(Component);
  if PlView.PlannerViewCount > 1
    then Result := 3
    else Result := 1;
end;

function TPlannerEhEditor.GetVerb(Index: Integer): string;
begin
  case Index of
    0: Result := 'PlannerView list editor...';
    1: Result := 'Next PlannerView';
    2: Result := 'Prior PlannerView';
  end;
end;

procedure TPlannerEhEditor.ExecuteVerb(Index: Integer);
begin
  case Index of
    {$IFDEF FPC}
    {$ELSE}
    0: ShowComponentChildrenEditor(Designer, Component);
    {$ENDIF}
    1: NextPlannerGrid;
    2: PriorPlannerGrid;
  end;
end;

procedure TPlannerEhEditor.NextPlannerGrid;
var
  PlView: TPlannerControlEh;
begin
  PlView := TPlannerControlEh(Component);
  if PlView.ActivePlannerView.ViewIndex < PlView.PlannerViewCount-1 then
    PlView.ActivePlannerView := PlView.PlannerView[PlView.ActivePlannerView.ViewIndex+1]
  else
    PlView.ActivePlannerView := PlView.PlannerView[0];
end;

procedure TPlannerEhEditor.PriorPlannerGrid;
var
  PlView: TPlannerControlEh;
begin
  PlView := TPlannerControlEh(Component);
  if PlView.ActivePlannerView.ViewIndex > 0 then
    PlView.ActivePlannerView := PlView.PlannerView[PlView.ActivePlannerView.ViewIndex-1]
  else
    PlView.ActivePlannerView := PlView.PlannerView[PlView.PlannerViewCount-1];
end;

{ TItemSourceFieldMapProperty }

type

  TItemSourceFieldMapProperty = class(TStringProperty)
  public
    function GetAttributes: TPropertyAttributes; override;
    procedure GetValueList(List: TStrings); virtual;
    procedure GetValues(Proc: TGetStrProc); override;
  end;

function TItemSourceFieldMapProperty.GetAttributes: TPropertyAttributes;
begin
  Result := [paValueList, paSortList, paMultiSelect];
end;

procedure TItemSourceFieldMapProperty.GetValueList(List: TStrings);
var
  PlSrc: TItemSourceFieldsMapEh;
begin
  if (GetComponent(0) = nil) then Exit;
  PlSrc := nil;
  if (GetComponent(0) is TItemSourceFieldsMapEh) then
    PlSrc := (GetComponent(0) as TItemSourceFieldsMapEh);
  if (PlSrc <> nil) and (PlSrc.SourceParams.DataSet <> nil) then
  begin
    PlSrc.SourceParams.DataSet.GetFieldNames(List);
  end;
end;

procedure TItemSourceFieldMapProperty.GetValues(Proc: TGetStrProc);
var
  i: Integer;
  Values: TStringList;
begin
  Values := TStringList.Create;
  try
    GetValueList(Values);
    for i := 0 to Values.Count - 1 do Proc(Values[i]);
  finally
    Values.Free;
  end;
end;

type

{ TPlannerDataSourceEditorEh  }

  TPlannerDataSourceEditorEh = class(TComponentEditor)
    function  GetVerbCount: Integer; override;
    function  GetVerb(Index: Integer): string; override;
    procedure ExecuteVerb(Index: Integer); override;
  end;

procedure TPlannerDataSourceEditorEh.ExecuteVerb(Index: Integer);
{$IFDEF FPC}
{$ELSE}
var
  Collection: TCollection;
{$ENDIF}
begin
{$IFDEF FPC}
{$ELSE}
  if Component is TPlannerDataSourceEh then
    Collection := (Component as TPlannerDataSourceEh).ItemSourceParams.FieldsMap
  else
    Collection := nil;
{$ENDIF}
  case Index of
    0 :
    {$IFDEF FPC}
      ;
    {$ELSE}
      if Collection <> nil then
        ShowCollectionEditorClass(Designer, TDefaultItemsCollectionEditorEh, Component,
          Collection, 'FieldsMap', [coAdd, coDelete, coMove]);
    {$ENDIF}
    2:
      ShowAboutForm;
  end;
end;

function TPlannerDataSourceEditorEh.GetVerb(Index: Integer): string;
begin
  case Index of
    0: Result := 'Edit ItemSourceParams.FieldsMap collection';
    1: Result := '-';
    2: Result := EhLibVerInfo + ' ' + EhLibBuildInfo + ' ' + EhLibEditionInfo;
  end;
end;

function TPlannerDataSourceEditorEh.GetVerbCount: Integer;
begin
  Result := 3;
end;

{$ENDIF}

procedure Register;
begin
{$IFDEF DESIGNTIME}
  RegisterComponents('EhLib Controls', [TPlannerControlEh]);
  RegisterComponents('EhLib Controls', [TPlannerCalendarPickerEh]);
  RegisterComponents('EhLib Components', [TPlannerDataSourceEh]);

  RegisterComponentEditor(TPlannerControlEh, TPlannerEhEditor);
  RegisterComponentEditor(TPlannerDataSourceEh, TPlannerDataSourceEditorEh);

{$IFDEF FPC}
{$ELSE}
  RegisterPropertyEditor(TypeInfo(TCollection), TPlannerItemSourceParamsEh, 'FieldsMap', TDefaultItemsCollectionProperty);
{$ENDIF}

//  RegisterPropertyEditor(TypeInfo(string), TItemSourceFieldMapEh, 'TitleFieldName', TItemSourceFieldMapProperty);
//  RegisterPropertyEditor(TypeInfo(string), TItemSourceFieldMapEh, 'BodyFieldName', TItemSourceFieldMapProperty);
//  RegisterPropertyEditor(TypeInfo(string), TItemSourceFieldMapEh, 'StartTimeFieldName', TItemSourceFieldMapProperty);
//  RegisterPropertyEditor(TypeInfo(string), TItemSourceFieldMapEh, 'EndTimeFieldName', TItemSourceFieldMapProperty);
//  RegisterPropertyEditor(TypeInfo(string), TItemSourceFieldMapEh, 'AllDayFieldName', TItemSourceFieldMapProperty);
//  RegisterPropertyEditor(TypeInfo(string), TItemSourceFieldMapEh, 'ResourceIDFieldName', TItemSourceFieldMapProperty);
//  RegisterPropertyEditor(TypeInfo(string), TItemSourceFieldMapEh, 'ItemIDFieldName', TItemSourceFieldMapProperty);
//  RegisterPropertyEditor(TypeInfo(string), TItemSourceFieldMapEh, 'FillColorFieldName', TItemSourceFieldMapProperty);

//  RegisterPropertyEditor(TypeInfo(TItemSourceFieldsMapEh), TCustomADODataSet, 'Parameters', TCollectionProperty);


  RegisterClasses([TPlannerDayViewEh, TPlannerWeekViewEh,
    TPlannerMonthViewEh,
    TPlannerVertDayslineViewEh, TPlannerHorzDayslineViewEh,
    TPlannerVertHourslineViewEh, TPlannerHorzHourslineViewEh
  ]);

  {$IFDEF FPC}
  {$ELSE}
  UnlistPublishedProperty(TCustomPlannerViewEh, 'AlignWithMargins');
  UnlistPublishedProperty(TCustomPlannerViewEh, 'Left');
  UnlistPublishedProperty(TCustomPlannerViewEh, 'Top');
  UnlistPublishedProperty(TCustomPlannerViewEh, 'Width');
  UnlistPublishedProperty(TCustomPlannerViewEh, 'Height');
  UnlistPublishedProperty(TCustomPlannerViewEh, 'Margins');
  {$ENDIF}

{$ENDIF}
end;

{$IFDEF FPC}
{$ELSE}
{ TPlannersDesignServiceEh }

class function TPlannersDesignServiceEh.CreateChild(MasterComponent: TComponent;
  ChildClass: TComponentClass): TComponent;
var
  NewName: String;
  n: Integer;
begin
  Result := TPlannerControlEh(MasterComponent).
    CreatePlannerGrid(TCustomPlannerGridClassEh(ChildClass), MasterComponent.Owner);

  n := 0;
  while True do
  begin
    Inc(n);
    NewName := Copy(Result.ClassName, 2, Length(Result.ClassName)) + IntToStr(n);
    if Result.Owner.FindComponent(NewName) <> nil then Continue;
    Result.Name := NewName;
    Break;
  end;
end;

class procedure TPlannersDesignServiceEh.GetChildClasses(ClassList: TList);
begin
  ClassList.Add(TPlannerDayViewEh);
  ClassList.Add(TPlannerWeekViewEh);
  ClassList.Add(TPlannerMonthViewEh);
  ClassList.Add(TPlannerVertDayslineViewEh);
  ClassList.Add(TPlannerVertHourslineViewEh);
  ClassList.Add(TPlannerHorzDayslineViewEh);
  ClassList.Add(TPlannerHorzHourslineViewEh);
end;

class function TPlannersDesignServiceEh.GetFormCaption(
  MasterComponent: TComponent): String;
begin
  Result := MasterComponent.Name + '.PlannerViews Editor';
end;

class procedure TPlannersDesignServiceEh.MoveChildDown(MasterComponent,
  ChildComponent: TComponent);
var
  PlannerControl: TPlannerControlEh;
  PlannerView: TCustomPlannerViewEh;
  Designer: IDesignerNotify;
begin
  PlannerControl := TPlannerControlEh(MasterComponent);
  PlannerView := TCustomPlannerViewEh(ChildComponent);
  if PlannerView.ViewIndex < PlannerControl.PlannerViewCount-1  then
  begin
    PlannerView.ViewIndex := PlannerView.ViewIndex + 1;
    Designer := FindRootDesigner(MasterComponent);
    if Designer <> nil then
      Designer.Modified;
  end;
end;

class procedure TPlannersDesignServiceEh.MoveChildUp(MasterComponent,
  ChildComponent: TComponent);
var
//  PlannerControl: TPlannerControlEh;
  PlannerView: TCustomPlannerViewEh;
  Designer: IDesignerNotify;
begin
  PlannerView := TCustomPlannerViewEh(ChildComponent);
  if PlannerView.ViewIndex > 0 then
  begin
    PlannerView.ViewIndex := PlannerView.ViewIndex - 1;
    Designer := FindRootDesigner(MasterComponent);
    if Designer <> nil then
      Designer.Modified;
  end;
end;

{$ENDIF}

initialization
{$IFDEF FPC}
{$ELSE}
  RegisterCompChildrenDesignService(TPlannerControlEh, TPlannersDesignServiceEh)
{$ENDIF}
finalization
  FreeAndNil(ImpExpListEditorsList);
{$IFDEF FPC}
{$ELSE}
  UnregisterCompChildrenDesignService(TPlannersDesignServiceEh);
{$ENDIF}
end.
