{*******************************************************}
{                                                       }
{                       EhLib v8.1                      }
{               Design Time TCompoListEditor            }
{                      Build 8.1.002                    }
{                                                       }
{   Copyright (c) 2014-2015 by Dmitry V. Bolshakov      }
{                                                       }
{*******************************************************}

{$I EhLib.Inc}

unit CompoMansDesignEh;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, Dialogs, ComCtrls, ToolWin,
{$IFDEF EH_LIB_9} Types, {$ENDIF}
{$IFDEF EH_LIB_17} System.UITypes, {$ENDIF}
  DBGridEhGrouping, ToolCtrlsEh, DBGridEhToolCtrls, DynVarsEh, MemTableDataEh,
  ObjectInspectorEh,
{$IFDEF DESIGNTIME}
  {$IFDEF FPC}
  LCLType, PropEdits, ComponentEditors,
  {$ELSE}
  DesignIntf, DesignEditors, VCLEditors, ToolsAPI, DesignWindows, TypInfo,
  {$ENDIF}
  EhLibReg,
{$ENDIF}
  {$IFDEF FPC}
  EhLibLCL, MaskEdit, DBGridsEh,
  {$ELSE}
  EhLibVCL, Mask, DBGridEh,
  {$ENDIF}
  CompoMansEh, EhLibMTE, StrUtils,
  DB, MemTableEh, GridsEh, DBAxisGridsEh, StdCtrls,
  DBCtrlsEh;

type
//{$IFDEF DESIGNTIME}
  TCompoListEditor = class(TDesignWindow)
//{$ELSE}
//  TCompoListEditor = class(TForm)
//{$ENDIF}
    ToolBar1: TToolBar;
    ToolButton3: TToolButton;
    tbShowAll: TToolButton;
    tbHideAll: TToolButton;
    DBGridEh1: TDBGridEh;
    DataSource1: TDataSource;
    MemTableEh1: TMemTableEh;
    MemTableEh1CompName: TStringField;
    MemTableEh1RefComp: TRefObjectField;
    tbHide: TToolButton;
    tbShow: TToolButton;
    MemTableEh1PositionStr: TStringField;
    MemTableEh1OldPositionsX: TIntegerField;
    MemTableEh1OldPositionsY: TIntegerField;
    MemTableEh1CompTypeName: TStringField;
    procedure tbAddClick(Sender: TObject);
    procedure tlDelClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure DataSource1DataChange(Sender: TObject; Field: TField);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure tbHideClick(Sender: TObject);
    procedure tbShowClick(Sender: TObject);
    procedure ToolButton6Click(Sender: TObject);
    procedure tbHideAllClick(Sender: TObject);
    procedure DBGridEh1ApplyFilter(Sender: TObject);
    procedure DBGridEh1ContextPopup(Sender: TObject; MousePos: TPoint;
      var Handled: Boolean);
    procedure DBGridEh1Columns0GetCellParams(Sender: TObject; EditMode: Boolean;
      Params: TColCellParamsEh);
    procedure DBGridEh1SelectionChanged(Sender: TObject);
  private
    FCompoList: TCompoManEh;
    FChilderList: TList;
    FInSelection: Boolean;
    FOldPosList: TStringList;
    FClosing: Boolean;
    FCollectionClassName: string;
    FInertnalSetSelection: Boolean;
    procedure SetCompoList(const Value: TCompoManEh);
    procedure SaveOldPos(const Comp: TComponent);
    procedure SaveCompoList;
    procedure LoadCompoList;
    procedure HideComponent(const Comp: TComponent);
    function GetDelOldPos(const Comp: TComponent; var DesignInfo: LongInt): Boolean;
    function IsComponentHiden(const Comp: TComponent): Boolean;
    function GetRegKey: string;
    procedure ReadRegState;
    procedure WriteRegState;

//    procedure ProcessGetStrProc(const Value: string);
//    procedure AddOneChildToList(Child: TComponent);
  public
    procedure InitForm;
    procedure UpdateList;
    procedure UpdateChildrenList;

    function SelectNewChildClass: TComponentClass; virtual;
    function CreateChild(ChildClass: TComponentClass): TComponent;
    function GetSelectedList: IDesignerSelections;
    function CheckItemInList(Item: TPersistent): Boolean;

    procedure XModuleServicesSaveAll;
    procedure CloseEditor;
    procedure UpdateActionsState;

{$IFDEF DESIGNTIME}
    procedure ItemDeleted(const ADesigner: IDesigner; Item: TPersistent); override;
    procedure ItemsModified(const Designer: IDesigner); override;
    procedure SelectionChanged(const ADesigner: IDesigner; const ASelection: IDesignerSelections); override;
    procedure DesignerClosed(const ADesigner: IDesigner; AGoingDormant: Boolean); override;
{$ENDIF}

    property CompoList: TCompoManEh read FCompoList write SetCompoList;
  end;

//var
//  ImpExpListEditor2: TCompChildrenEditor;

{ TCompChildrenDesignServiceEh }

  TCompChildrenDesignServiceEh = class
    class procedure GetChildClasses(ClassList: TList); virtual;
    class function CreateChild(MasterComponent: TComponent; ChildClass: TComponentClass): TComponent; virtual;
  end;

  TCompChildrenDesignServiceClassEh = class of TCompChildrenDesignServiceEh;

function ShowComponentListEditor(
{$IFDEF DESIGNTIME}
  ADesigner: IDesigner;
{$ENDIF}
  ACompoList: TCompoManEh): TCompoListEditor;

procedure Register;

implementation

uses ComponentChildrenDesignSelectClassDialogEh, DBConsts, DesignConst,
  Registry, DesignMenus;

{$R *.dfm}

type
  TComponentCrack = class(TComponent);
  TCompoManEhCrack = class(TCompoManEh);
  TCustomDBGridEhCrack = class(TCustomDBGridEh);


{$IFNDEF EH_LIB_14}
function FindDelimiter(const Delimiters, S: string; StartIdx: Integer = 1): Integer;
var
  Stop: Boolean;
  Len: Integer;
begin
  Result := 0;

  Len := Length(S);
  Stop := False;
  while (not Stop) and (StartIdx <= Len) do
    if IsDelimiter(Delimiters, S, StartIdx) then
    begin
      Result := StartIdx;
      Stop := True;
    end
    else
      Inc(StartIdx);
end;

function SplitString(const S, Delimiters: string): TStringDynArray;
var
  StartIdx: Integer;
  FoundIdx: Integer;
  SplitPoints: Integer;
  CurrentSplit: Integer;
  i: Integer;
begin
  Result := nil;

  if S <> '' then
  begin
    { Determine the length of the resulting array }
    SplitPoints := 0;
    for i := 1 to Length(S) do
      if IsDelimiter(Delimiters, S, i) then
        Inc(SplitPoints);

    SetLength(Result, SplitPoints + 1);

    { Split the string and fill the resulting array }
    StartIdx := 1;
    CurrentSplit := 0;
    repeat
      FoundIdx := FindDelimiter(Delimiters, S, StartIdx);
      if FoundIdx <> 0 then
      begin
        Result[CurrentSplit] := Copy(S, StartIdx, FoundIdx - StartIdx);
        Inc(CurrentSplit);
        StartIdx := FoundIdx + 1;
      end;
    until CurrentSplit = SplitPoints;

    // copy the remaining part in case the string does not end in a delimiter
    Result[SplitPoints] := Copy(S, StartIdx, Length(S) - StartIdx + 1);
  end;
end;
{$ENDIF}

type

{ TDesignNotificationHandler }

  TDesignNotificationHandler = class(TInterfacedObject, IDesignNotification)
  protected
    procedure DesignerClosed(const ADesigner: IDesigner; AGoingDormant: Boolean);
    procedure DesignerOpened(const ADesigner: IDesigner; AResurrecting: Boolean);
    procedure ItemDeleted(const ADesigner: IDesigner; AItem: TPersistent);
    procedure ItemInserted(const ADesigner: IDesigner; AItem: TPersistent);
    procedure ItemsModified(const ADesigner: IDesigner);
    procedure SelectionChanged(const ADesigner: IDesigner; const ASelection: IDesignerSelections);
  end;

{ TChildrenDesignServiceKey }

  TChildrenDesignServiceKey = class(TPersistent)
  public
    ComponentClass: TComponentClass;
    DesignServiceClass: TCompChildrenDesignServiceClassEh;
  end;

var
  ChildrenEditorsList: TList = nil;
  ChildrenDesignServiceList: TList = nil;
  DesignNotification: IDesignNotification;

{$IFDEF DESIGNTIME}

type

  TCompoManEhEditor = class(TComponentEditor)
    function  GetVerbCount: Integer; override;
    function  GetVerb(Index: Integer): string; override;
    procedure ExecuteVerb(Index: Integer); override;
  end;

function TCompoManEhEditor.GetVerbCount: Integer;
begin
  Result := 1;
end;

function TCompoManEhEditor.GetVerb(Index: Integer): string;
begin
  case Index of
    0: Result := 'Component List editor...';
  end;
end;

procedure TCompoManEhEditor.ExecuteVerb(Index: Integer);
begin
  case Index of
    0: ShowComponentListEditor(Designer, TCompoManEh(Component));
  end;
end;

{$ENDIF}

function ShowComponentListEditor(
{$IFDEF DESIGNTIME}
  ADesigner: IDesigner;
{$ENDIF}
  ACompoList: TCompoManEh): TCompoListEditor;
var
  I: Integer;
begin
  if ChildrenEditorsList = nil then
    ChildrenEditorsList := TList.Create;

  for I := 0 to ChildrenEditorsList.Count-1 do
  begin
    Result := TCompoListEditor(ChildrenEditorsList[I]);
    with Result do
{$IFDEF DESIGNTIME}
      if (Designer = ADesigner) and (CompoList = ACompoList) then
{$ENDIF}
      begin
        Show;
        BringToFront;
        Exit;
      end;
  end;

  Result := TCompoListEditor.Create(Application);
  with Result do
  try
{$IFDEF DESIGNTIME}
    Designer := ADesigner;
{$ENDIF}
    CompoList := ACompoList;
    LoadCompoList;
    Show;
  except
    Free;
  end;
end;

{ TCompChildrenEditor }

procedure TCompoListEditor.DataSource1DataChange(Sender: TObject;
  Field: TField);
{$IFDEF DESIGNTIME}
var
  Child: TComponent;
{$ENDIF}
begin
  if FInSelection then Exit;

  FInSelection := True;
  try
  if not MemTableEh1.IsEmpty and (DBGridEh1.SelectedRows.Count = 0) then
  begin
{$IFDEF DESIGNTIME}
    Child := TComponent(MemTableEh1RefComp.Value);
    Designer.SelectComponent(Child);
{$ENDIF}
  end;
  finally
    FInSelection := False;
  end;
end;

procedure TCompoListEditor.DBGridEh1ApplyFilter(Sender: TObject);
begin
  FInSelection := True;
  try
    DBGridEh1.DefaultApplyFilter;
  finally
    FInSelection := False;
  end;
end;

procedure TCompoListEditor.DBGridEh1Columns0GetCellParams(Sender: TObject;
  EditMode: Boolean; Params: TColCellParamsEh);
begin
  if DBGridEh1.SelectedRows.Count > 0 then
  begin
    if DBGridEh1.SelectedRows.CurrentRowSelected then
      Params.State := [gdSelected]
    else
      Params.State := [];
  end;
end;

procedure TCompoListEditor.DBGridEh1ContextPopup(Sender: TObject;
  MousePos: TPoint; var Handled: Boolean);
var
  LDesignLocalMenu: IDesignLocalMenu;
  Sel: IDesignerSelections;
  LDesigner: IDesigner;
  LPopupMenu: IPopupMenu;
  P: TPoint;
begin
  Supports(FindRootDesigner(CompoList), IDesigner, LDesigner);
  if Supports(FindRootDesigner(CompoList), IDesignLocalMenu, LDesignLocalMenu) then
  begin
    TCustomDBGridEhCrack(DBGridEh1).CancelMode;
    Sel := CreateSelectionList;
    LDesigner.GetSelections(Sel);
{$IFDEF EH_LIB_13}
    LPopupMenu := LDesignLocalMenu.BuildLocalMenu(Sel, [lmComponent]);
{$ELSE}
    LPopupMenu := LDesignLocalMenu.BuildLocalMenu([lmComponent]);
{$ENDIF}
    P := DBGridEh1.ClientToScreen(Point(MousePos.X, MousePos.Y));

//        if (ssLeft in Shift) and (ssDouble in Shift) and (DropdownMenu.Items.Count > 0)
//          then (LPopupMenu.Items.GetItem(0) as IMenuItem2).Click
//          else
    LPopupMenu.Popup(P.X, P.Y);

    Exit;
  end;
end;

procedure TCompoListEditor.FormCreate(Sender: TObject);
begin
  FChilderList := TList.Create;
  FOldPosList := TStringList.Create;
  ChildrenEditorsList.Add(Self);
end;

procedure TCompoListEditor.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  WriteRegState;
  Action := caFree;
end;

procedure TCompoListEditor.FormDestroy(Sender: TObject);
begin
  if ChildrenEditorsList <> nil then
    ChildrenEditorsList.Remove(Self);
  FreeAndNil(FChilderList);
  FreeAndNil(FOldPosList);
end;

procedure TCompoListEditor.DBGridEh1SelectionChanged(Sender: TObject);
var
  i: Integer;
  Rec: TMemoryRecordEh;
  bm: TUniBookmarkEh;
  SelList: IDesignerSelections;
  Comp: TComponent;
begin
  if FInertnalSetSelection then Exit;
  if FInSelection then Exit;
  SelList := CreateSelectionList;
  for i := 0 to DBGridEh1.SelectedRows.Count-1 do
  begin
    bm := DBGridEh1.SelectedRows[i];
    Rec := MemTableEh1.BookmarkToRec(bm);
    if Rec <> nil then
    begin
      Comp := TComponent(Integer(Rec.DataValues['RefComp', dvvValueEh]));
      SelList.Add(Comp);
    end;
  end;
  FInertnalSetSelection := True;
  try
    Designer.SetSelections(SelList)
  finally
    FInertnalSetSelection := False;
  end;
  UpdateActionsState;
end;

procedure TCompoListEditor.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_F11 then
  begin
    ShowObjectInspectorForm(MemTableEh1RefComp.Value, Rect(Left+Width+10, Top, Left+Width+10+300, Top+Height));
  end;
end;

procedure TCompoListEditor.InitForm;
begin
  ReadRegState;
  MemTableEh1.Open;
  UpdateList;
  Caption := 'Form Components Manager - ' + CompoList.Name;
end;

function TCompoListEditor.IsComponentHiden(const Comp: TComponent): Boolean;
begin
  Result := Comp.DesignInfo = MakeLong(Word(0), Word(-50));
end;

procedure TCompoListEditor.HideComponent(const Comp: TComponent);
begin
  if Comp = nil then Exit;
  SaveOldPos(Comp);
  Comp.DesignInfo := MakeLong(Word(0), Word(-50));
end;

{$IFDEF DESIGNTIME}

procedure TCompoListEditor.DesignerClosed(const ADesigner: IDesigner;
  AGoingDormant: Boolean);
begin
  inherited DesignerClosed(ADesigner, AGoingDormant);
  if Designer = ADesigner then
    Close;
end;

procedure TCompoListEditor.ItemDeleted(const ADesigner: IDesigner; Item: TPersistent);
begin
  inherited ItemDeleted(ADesigner, Item);
  if FClosing then Exit;
  if Item = CompoList then
    CloseEditor
  else
  begin
    if CheckItemInList(Item) then
      UpdateList;
  end;
end;

procedure TCompoListEditor.ItemsModified(const Designer: IDesigner);
begin
  inherited ItemsModified(Designer);
  if FClosing then Exit;
  UpdateList;
end;

procedure TCompoListEditor.SelectionChanged(const ADesigner: IDesigner;
  const ASelection: IDesignerSelections);
var
//  i: Integer;
  SecComp: TPersistent;
begin
  inherited SelectionChanged(ADesigner, ASelection);
  if FInertnalSetSelection then Exit;
  if FClosing then Exit;
  if ASelection.Count = 0 then Exit;
  SecComp := ASelection.Items[0];

  if not FInSelection then
  begin
    FInSelection := True;
    MemTableEh1.Locate('RefComp', Integer(SecComp), []);
    FInSelection := False;
  end else
    MemTableEh1.Locate('RefComp', Integer(SecComp), []);
  UpdateActionsState;
end;

{$ENDIF}

//procedure TCompoListEditor.AddOneChildToList(Child: TComponent);
//begin
//  FChilderList.Add(Child);
//end;

procedure TCompoListEditor.UpdateActionsState;
var
  SelList: IDesignerSelections;
  i: Integer;
  ShowCanActive, HideCanActive: Boolean;
begin
  ShowCanActive := False;
  HideCanActive := False;
  SelList := GetSelectedList;

  for i := 0 to SelList.Count-1 do
  begin
    if IsComponentHiden(TComponent(SelList[i])) then
      ShowCanActive := True
    else
      HideCanActive := True;
  end;
  tbHide.Enabled := HideCanActive;
  tbShow.Enabled := ShowCanActive;
end;

function ListSortCompare(Item1, Item2: Pointer): Integer;
begin
  Result := CompareText(TComponent(Item1).Name, TComponent(Item2).Name);
end;

procedure TCompoListEditor.UpdateChildrenList;
var
  i: Integer;
begin
  FChilderList.Clear;
  for i := 0 to CompoList.Owner.ComponentCount-1 do
    if not (CompoList.Owner.Components[i] is TControl) and
       not CompoList.Owner.Components[i].HasParent and
          (CompoList.Owner.Components[i] <> CompoList)
    then
      FChilderList.Add(CompoList.Owner.Components[i]);
  FChilderList.Sort(ListSortCompare);
end;

function TCompoListEditor.CheckItemInList(Item: TPersistent): Boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to FChilderList.Count-1 do
    if FChilderList[i] = Item then
    begin
      Result := True;
      Break;
    end;
end;

procedure TCompoListEditor.UpdateList;
var
  i: Integer;
  ChildItem: TComponent;
  CompPos: TPoint;
begin
  UpdateChildrenList;
  ChildItem := TComponent(MemTableEh1RefComp.Value);
  MemTableEh1.DisableControls;
  FInSelection := True;
  try
    MemTableEh1.EmptyTable;
    for i := 0 to FChilderList.Count-1 do
    begin
      MemTableEh1.Append;
      MemTableEh1CompName.AsString := TComponent(FChilderList[i]).Name;
      MemTableEh1CompTypeName.AsString := TComponent(FChilderList[i]).ClassName;
      MemTableEh1RefComp.Value := FChilderList[i];
      CompPos.X := Smallint(LongRec(TComponent(FChilderList[i]).DesignInfo).Lo);
      CompPos.Y := Smallint(LongRec(TComponent(FChilderList[i]).DesignInfo).Hi);
      if (CompPos.X = 0) and (CompPos.Y = -50)  then
        MemTableEh1PositionStr.AsString := '(Hiden)'
      else
        MemTableEh1PositionStr.AsString :=
          '(' + IntToStr(CompPos.X) + ',' + IntToStr(CompPos.Y) + ')';
      MemTableEh1.Post;
    end;
    MemTableEh1.Locate('RefComp', Integer(ChildItem), []);
  finally
    MemTableEh1.EnableControls;
    FInSelection := False;
  end;
end;

procedure TCompoListEditor.XModuleServicesSaveAll;
var
  XModuleServices: IOTAModuleServices;
  XResult: HResult;
begin
  Designer.Modified;
  XResult := BorlandIDEServices.QueryInterface(IOTAModuleServices, XModuleServices);
  if (XResult=S_Ok) and Assigned(XModuleServices) then XModuleServices.SaveAll;
end;

procedure TCompoListEditor.SetCompoList(const Value: TCompoManEh);
begin
  FCompoList := Value;
  FCollectionClassName := Value.ClassName;
  InitForm;
end;

function TCompoListEditor.SelectNewChildClass: TComponentClass;
begin
  Result := nil;
end;

procedure TCompoListEditor.CloseEditor;
begin
  FClosing := True;
  Close;
end;

function TCompoListEditor.CreateChild(ChildClass: TComponentClass): TComponent;
//var
//  DesignService: TCompChildrenDesignServiceClassEh;
begin
//  DesignService := GetDesignServiceByClass(TComponentClass(CompoList.ClassType));
//  Result := DesignService.CreateChild(CompoList, ChildClass);
  Result := nil;
end;

procedure TCompoListEditor.tbAddClick(Sender: TObject);
var
  ChildClass: TComponentClass;
//  Child: TComponent;
//  NewName: String;
//  n: Integer;
begin
  ChildClass := SelectComponentChildEhClass(TComponentClass(CompoList.ClassType));
  if ChildClass <> nil then
  begin
    CreateChild(ChildClass);
    UpdateList;
  end;
end;

procedure TCompoListEditor.SaveCompoList;
var
  i: Integer;
  Comp: TComponent;
  ACompoList: TCompoManEhCrack;
  Pos: TPoint;
begin
  ACompoList := TCompoManEhCrack(CompoList);
  ACompoList.FVisibleComponentListPos.Clear;
  for i := 0 to FOldPosList.Count-1 do
  begin
    Comp := CompoList.Owner.FindComponent(FOldPosList[i]);
    if Comp <> nil then
    begin
      Pos.X := Smallint(LongRec(FOldPosList.Objects[i]).Lo);
      Pos.Y := Smallint(LongRec(FOldPosList.Objects[i]).Hi);
      ACompoList.FVisibleComponentListPos.Add(
        FOldPosList[i] + ',' + IntToStr(Pos.X) + ',' + IntToStr(Pos.Y));
    end;
  end;
end;

procedure TCompoListEditor.LoadCompoList;
var
  i: Integer;
  ACompoList: TCompoManEhCrack;
  Pos: TPoint;
  SplitData: TStringDynArray;
begin
  SplitData := nil;
  ACompoList := TCompoManEhCrack(CompoList);
  FOldPosList.Clear;
  for i := 0 to ACompoList.FVisibleComponentListPos.Count-1 do
  begin
    SplitData := SplitString(ACompoList.FVisibleComponentListPos[i], ',');
    Pos := Point(StrToInt(SplitData[1]), StrToInt(SplitData[2]));
    FOldPosList.AddObject(SplitData[0], TObject(MakeLong(Pos.X, Pos.Y)));
  end;
end;

procedure TCompoListEditor.SaveOldPos(const Comp: TComponent);
var
  Idx: Integer;
begin
  if Comp.DesignInfo <> MakeLong(Word(0), Word(-50)) then
  begin
    Idx := FOldPosList.IndexOf(Comp.Name);
    if Idx >= 0 then
    begin
      FOldPosList.Objects[Idx] := TObject(Comp.DesignInfo);
    end else
      FOldPosList.AddObject(Comp.Name, TObject(Comp.DesignInfo));
  end;
end;

function TCompoListEditor.GetDelOldPos(const Comp: TComponent; var DesignInfo: LongInt): Boolean;
var
  Idx: Integer;
begin
  Result := False;
  Idx := FOldPosList.IndexOf(Comp.Name);
  if Idx >= 0 then
  begin
    DesignInfo := LongInt(FOldPosList.Objects[Idx]);
    Result := True;
  end;
end;

function TCompoListEditor.GetRegKey: string;
begin
  Result := ComponentDesigner.Environment.GetBaseRegKey + '\' +
    sIniEditorsName + '\Collection Editor';
end;

procedure TCompoListEditor.ReadRegState;
var
//  I: Integer;
  ALeft, ATop, AWidth, AHeight: Integer;
  CentrPoint: TPoint;
begin
  with TRegIniFile.Create(GetRegKey) do
  try
    AWidth := ReadInteger(FCollectionClassName, 'Width', Width);
    AHeight := ReadInteger(FCollectionClassName, 'Height', Height);
    ALeft := ReadInteger(FCollectionClassName, 'Left', Left);
    ATop := ReadInteger(FCollectionClassName, 'Top', Top);
    SetBounds(ALeft, ATop, AWidth, AHeight);
    CentrPoint := Point(Left + Width div 2, Top + Height div 2);

    if CentrPoint.X < Screen.WorkAreaRect.Left then
      Left := Screen.WorkAreaRect.Left - Width div 2;
    if CentrPoint.Y < Screen.WorkAreaRect.Top then
      Top := Screen.WorkAreaRect.Top - Height div 2;

    if CentrPoint.X > Screen.WorkAreaRect.Right then
      Left := Screen.WorkAreaRect.Right - Width div 2;
    if CentrPoint.Y > Screen.WorkAreaRect.Bottom then
      Top := Screen.WorkAreaRect.Bottom - Height div 2;

//    Splitter1.Top := Toolbar1.Top + Toolbar1.Height;
//    ToolBar1.Visible := ReadBool(FCollectionClassName, 'Toolbar', True);
//    Splitter1.Visible := Toolbar1.Visible;
//    LargeButtons := ReadBool(FCollectionClassName, 'LargeButtons', False);
//    ListView1.HandleNeeded;
//    if ListView1.Columns.Count > 1 then
//      for I := 0 to ListView1.Columns.Count - 1 do
//        ListView1.Column[I].Width := ReadInteger(FCollectionClassName,
//          Format('Column%d', [I]), ListView1.Column[I].WidthType);
  finally
    Free;
  end;
end;

procedure TCompoListEditor.WriteRegState;
begin
  with TRegIniFile.Create(GetRegKey) do
  try
    EraseSection(FCollectionClassName);
    WriteInteger(FCollectionClassName, 'Width', Width);
    WriteInteger(FCollectionClassName, 'Height', Height);
    WriteInteger(FCollectionClassName, 'Left', Left);
    WriteInteger(FCollectionClassName, 'Top', Top);
//    WriteBool(FCollectionClassName, 'LargeButtons', LargeButtons);
//    WriteBool(FCollectionClassName, 'Toolbar', ToolBar1.Visible);
//    for I := 0 to ListView1.Columns.Count - 1 do
//      WriteInteger(FCollectionClassName, Format('Column%d', [I]),
//        ListView1.Column[I].WidthType);
  finally
    Free;
  end;
end;

procedure TCompoListEditor.tbHideAllClick(Sender: TObject);
var
  i: Integer;
  Rec: TMemoryRecordEh;
  Comp: TComponent;
begin
  for i := 0 to MemTableEh1.RecordsView.MemTableData.RecordsList.Count-1 do
  begin
    Rec := MemTableEh1.RecordsView.MemTableData.RecordsList.Rec[i];
    Comp := TComponent(Integer(Rec.DataValues['RefComp', dvvValueEh]));
    if not IsComponentHiden(Comp) then
      HideComponent(Comp);
  end;
  XModuleServicesSaveAll;
  SaveCompoList;
end;

function TCompoListEditor.GetSelectedList: IDesignerSelections;
var
//  SelList: IDesignerSelections;
  Comp: TComponent;
  i: Integer;
  bm: TUniBookmarkEh;

  procedure AddToSel(Rec: TMemoryRecordEh);
  begin
    if Rec <> nil then
    begin
      Comp := TComponent(Integer(Rec.DataValues['RefComp', dvvValueEh]));
      Result.Add(Comp);
    end;
  end;

begin
  Result := CreateSelectionList;
  if DBGridEh1.SelectedRows.Count = 0 then
    AddToSel(MemTableEh1.Rec)
  else
    for i := 0 to DBGridEh1.SelectedRows.Count-1 do
    begin
      bm := DBGridEh1.SelectedRows[i];
      AddToSel(MemTableEh1.BookmarkToRec(bm));
    end;
end;

procedure TCompoListEditor.tbHideClick(Sender: TObject);
{$IFDEF DESIGNTIME}
var
  Child: TComponent;
  SelList: IDesignerSelections;
  i: Integer;
{$ENDIF}
begin
  if not MemTableEh1.IsEmpty then
  begin
{$IFDEF DESIGNTIME}
    SelList := GetSelectedList;

    for i := 0 to SelList.Count-1 do
    begin
      Child := TComponent(SelList[i]);
      if Child = nil then Exit;
      SaveOldPos(Child);
      Child.DesignInfo := MakeLong(Word(0), Word(-50));
    end;

    XModuleServicesSaveAll;
    SaveCompoList;
    UpdateActionsState;
{$ENDIF}
  end;
end;

procedure TCompoListEditor.tbShowClick(Sender: TObject);
{$IFDEF DESIGNTIME}
var
  Child: TComponent;
  NewPos: TPoint;
  DesignInfo: LongInt;
  SelList: IDesignerSelections;
  i: Integer;
{$ENDIF}
begin
  if not MemTableEh1.IsEmpty then
  begin
{$IFDEF DESIGNTIME}

    SelList := GetSelectedList;

    for i := 0 to SelList.Count-1 do
    begin
      Child := TComponent(SelList[i]);
      if (Child = nil) or not IsComponentHiden(Child) then Continue;
      if GetDelOldPos(Child, DesignInfo) then
        Child.DesignInfo := DesignInfo
      else
      begin
        NewPos := Point(TControl(CompoList.Owner).ClientWidth,
                        TControl(CompoList.Owner).ClientHeight);
        NewPos.X := NewPos.X div 2 - 12;
        NewPos.Y := NewPos.Y div 2 - 12;
        Child.DesignInfo := MakeLong(Word(NewPos.X), Word(NewPos.Y));
      end;
    end;

    XModuleServicesSaveAll;
    SaveCompoList;
    UpdateActionsState;
{$ENDIF}
  end;
end;

procedure TCompoListEditor.tlDelClick(Sender: TObject);
var
  Msg: String;
  Child: TComponent;
begin
  Msg := SDeleteRecordQuestion;
  if not MemTableEh1.IsEmpty then
  begin
    Child := TComponent(MemTableEh1RefComp.Value);
    if MessageDlg('Do you want to Delete ' + Child.Name, mtConfirmation, mbOKCancel, 0) <> mrCancel then
    begin
      MemTableEh1.Prior;
      Child.Free;
    end;
  end;
end;

//procedure TCompoListEditor.ProcessGetStrProc(const Value: string);
//begin
//  DBMemoEh1.Lines.Add(Value);
//end;

procedure TCompoListEditor.ToolButton6Click(Sender: TObject);
begin
//  Designer.GetComponentNames(GetTypeData(TComponent.ClassInfo), ProcessGetStrProc);
//  GetComponentNamesEh(DBMemoEh1.Lines, TComponent, True);
end;

{ TCompChildrenDesignServiceEh }

class function TCompChildrenDesignServiceEh.CreateChild(MasterComponent: TComponent;
  ChildClass: TComponentClass): TComponent;
begin
  Result := nil;
end;

class procedure TCompChildrenDesignServiceEh.GetChildClasses(ClassList: TList);
begin

end;

procedure RestoreComponentsPositions(CompoList: TCompoManEh);
var
//  Owner: TComponent;
  i: Integer;
  ACompoList: TCompoManEhCrack;
  Pos: TPoint;
  SplitData: TStringDynArray;
  ACompoName: String;
  Comp: TComponent;
begin

  SplitData := nil;
  ACompoList := TCompoManEhCrack(CompoList);
//  ShowMessage('ACompoList.Name ' + ACompoList.Name);
//  if ACompoList.FVisibleComponentListPos = nil then
//    ShowMessage('ACompoList.FVisibleComponentListPos = nil ')
//  else
//    ShowMessage('ACompoList.FVisibleComponentListPos.Count = ' + ACompoList.FVisibleComponentListPos.Count.ToString );

  for i := 0 to ACompoList.FVisibleComponentListPos.Count-1 do
  begin
    SplitData := SplitString(ACompoList.FVisibleComponentListPos[i], ',');
    Pos := Point(StrToInt(SplitData[1]), StrToInt(SplitData[2]));
    ACompoName := SplitData[0];
    Comp := CompoList.Owner.FindComponent(ACompoName);
    if Comp <> nil then
    begin
      Comp.DesignInfo := MakeLong(Word(Pos.X), Word(Pos.Y));
    end;
  end;

//  XModuleServicesSaveAll;
end;

procedure InitDesignNotification;
begin
  if DesignNotification = nil then
  begin
    DesignNotification := TDesignNotificationHandler.Create as IDesignNotification;
    RegisterDesignNotification(DesignNotification);
  end;
end;

procedure DoneDesignNotification;
begin
  if DesignNotification <> nil then
    UnregisterDesignNotification(DesignNotification);
  DesignNotification := nil;
end;

procedure InitUnit;
begin
  InitDesignNotification;
//  DesignTimeCompoListDeleteNotification := DesignTimeCompoListDeleteNotificationProc;
end;

procedure FinalizeUnit;
begin
  DoneDesignNotification;
end;

procedure Register;
begin
{$IFDEF DESIGNTIME}
  RegisterComponents('EhLib Components', [TCompoManEh]);
  RegisterComponentEditor(TCompoManEh, TCompoManEhEditor);

{$ENDIF}
end;

{ TDesignNotificationHandler }

procedure TDesignNotificationHandler.DesignerClosed(const ADesigner: IDesigner;
  AGoingDormant: Boolean);
begin

end;

procedure TDesignNotificationHandler.DesignerOpened(const ADesigner: IDesigner;
  AResurrecting: Boolean);
begin

end;

procedure TDesignNotificationHandler.ItemDeleted(const ADesigner: IDesigner;
  AItem: TPersistent);
var
  XModuleServices: IOTAModuleServices;
  XResult: HResult;
begin
  if AItem is TCompoManEh then
  begin
//    ShowMessage('TCompoManEh TDesignNotificationHandler.ItemDeleted');
    RestoreComponentsPositions(AItem as TCompoManEh);
    ADesigner.Modified;
    XResult := BorlandIDEServices.QueryInterface(IOTAModuleServices, XModuleServices);
    if (XResult=S_Ok) and Assigned(XModuleServices) then XModuleServices.SaveAll;
  end;
end;

procedure TDesignNotificationHandler.ItemInserted(const ADesigner: IDesigner;
  AItem: TPersistent);
begin

end;

procedure TDesignNotificationHandler.ItemsModified(const ADesigner: IDesigner);
begin

end;

procedure TDesignNotificationHandler.SelectionChanged(
  const ADesigner: IDesigner; const ASelection: IDesignerSelections);
begin

end;

initialization
  InitUnit;
finalization
  FinalizeUnit;
  FreeAndNil(ChildrenEditorsList);
  FreeAndNil(ChildrenDesignServiceList);
end.
