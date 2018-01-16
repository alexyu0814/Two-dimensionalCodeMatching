unit Xc12FileData5;

{-
********************************************************************************
******* XLSReadWriteII V5.00                                             *******
*******                                                                  *******
******* Copyright(C) 1999,2012 Lars Arvidsson, Axolot Data               *******
*******                                                                  *******
******* email: components@axolot.com                                     *******
******* URL:   http://www.axolot.com                                     *******
********************************************************************************
** Users of the XLSReadWriteII component must accept the following            **
** disclaimer of warranty:                                                    **
**                                                                            **
** XLSReadWriteII is supplied as is. The author disclaims all warranties,     **
** expressedor implied, including, without limitation, the warranties of      **
** merchantability and of fitness for any purpose. The author assumes no      **
** liability for damages, direct or consequential, which may result from the  **
** use of XLSReadWriteII.                                                     **
********************************************************************************
}

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}

interface

uses Classes, SysUtils, Contnrs, IniFiles,
     XLSUtils5,
     xpgPSimpleDOM, XLSReadWriteOPC5, Xc12DefaultData5,
     xpgParseContentType,
     xpgParseDocPropsApp;

type TXLSPivotCacheRecords = class(TObject)
protected
     FStream: TMemoryStream;
public
     constructor Create;
     destructor Destroy; override;

     property Data: TMemoryStream read FStream;
     end;

type TXLSPivotCacheDefinition = class(TIndexObject)
protected
     FUsageCount: integer;
     FDOM       : TXpgSimpleDOM;
     FCacheId   : integer;
     FRecords   : TXLSPivotCacheRecords;
public
     constructor Create;
     destructor Destroy; override;

     procedure Use;
     procedure Release;

     procedure SetSource(const AArea,ASheet: AxUCString);

     property DOM: TXpgSimpleDOM read FDOM write FDOM;
     property CacheId: integer read FCacheId write FCacheId;
     property Records: TXLSPivotCacheRecords read FRecords;
     end;

type TXLSPivotCacheDefinitions = class(TIndexObjectList)
private
     function GetItems(const Index: integer): TXLSPivotCacheDefinition;
protected
public
     constructor Create;
     destructor Destroy; override;

     procedure Release(AItem: TXLSPivotCacheDefinition);

     function  Add: TXLSPivotCacheDefinition;
     function  Find(const ACacheId: integer): TXLSPivotCacheDefinition; overload;
     function  Find(const AItem: TXLSPivotCacheDefinition): integer; overload;

     property Items[const Index: integer]: TXLSPivotCacheDefinition read GetItems; default;
     end;

type TXLSPivotTable = class(TObject)
protected
     FCache: TXLSPivotCacheDefinition;
     FDOM  : TXpgSimpleDOM;
public
     constructor Create;
     destructor Destroy; override;

     property Cache: TXLSPivotCacheDefinition read FCache write FCache;
     property DOM  : TXpgSimpleDOM read FDOM;
     end;

type TXLSPivotTables = class(TObjectList)
private
     function GetItems(const Index: integer): TXLSPivotTable;
protected
public
     constructor Create;

     function  Add: TXLSPivotTable;

     property Items[const Index: integer]: TXLSPivotTable read GetItems; default;
     end;

type TXLSSavedFileData = class(TObject)
protected
     FType   : AxUCString;
     FTarget : AxUCString;
     FContent: AxUCString;
     FData   : TStream;
public
     destructor Destroy; override;

     property Type_  : AxUCString read FType write FType;
     property Target : AxUCString read FTarget write FTarget;
     property Content: AxUCString read FContent write FContent;
     property Data   : TStream read FData write FData;
     end;

type TXLSSavedFileDataList = class(TObjectList)
private
     function GetItems(Index: integer): TXLSSavedFileData;
protected
public
     constructor Create;

     function  FindByType(AType: AxUCString): TXLSSavedFileData;

     procedure Add(AType,ATarget,AContent: AxUCString; AData: TStream);

     property Items[Index: integer]: TXLSSavedFileData read GetItems; default;
     end;

type TXLSSavedFileDataSheet = class(TObject)
protected
     FSheetData: TXLSSavedFileDataList;
     FTableData: TXLSSavedFileDataList;
public
     constructor Create;
     destructor Destroy; override;

     property SheetData: TXLSSavedFileDataList read FSheetData;
     property TableData: TXLSSavedFileDataList read FTableData;
     end;

type TXc12FileData = class(TObject)
private
     procedure SetUseAlternateZip(const Value: boolean);
protected
     FContentType  : TXPGDocContentType;
     FDocPropsApp  : TXPGDocDocPropsApp;
     FDocPropsCore : TXpgSimpleDOM;
     FOPC          : TOPC_XLSX;
     FSavedRoot    : TXLSSavedFileDataList;
     FSavedWorkbook: TXLSSavedFileDataList;
{$ifdef DELPHI_5}
     FSavedSheets  : TStringList;
{$else}
     FSavedSheets  : THashedStringList;
{$endif}
     FUseAlternateZip: boolean;

     procedure SetDefaultData;
     procedure ClearSavedSheets;

     procedure ReadContentTypes;
     procedure ReadDocPropsApp;
     procedure ReadDocPropsCore;

     procedure WriteDocPropsCore;
     procedure WriteTheme;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     procedure LoadFromStream(AZIPStream: TStream);
     procedure ReadUnusedData;
     procedure WriteUnusedData;
     procedure WriteUnusedDataSheet(AIndex: integer; AOPCSheet: TOPCItem);
     procedure BeginSaveToStream(AZIPStream: TStream);
     procedure CommitSaveToStream;

     procedure AddSaveSheet(AId: AxUCString);

     function  StreamByName(const AName: AxUCString): TStream;

     property OPC : TOPC_XLSX read FOPC;
     property DocPropsApp : TXPGDocDocPropsApp read FDocPropsApp;
     property DocPropsCore : TXpgSimpleDOM read FDocPropsCore;
     property UseAlternateZip: boolean read FUseAlternateZip write SetUseAlternateZip;
     end;

implementation

{ TXLSFileDataXLSX }

procedure TXc12FileData.AddSaveSheet(AId: AxUCString);
begin
  FSavedSheets.Add(AId);
end;

procedure TXc12FileData.BeginSaveToStream(AZIPStream: TStream);
begin
  FOPC.Clear;

  FOPC.OpenWrite(AZIPStream,FUseAlternateZip);

  WriteDocPropsCore;
end;

procedure TXc12FileData.Clear;
begin
  FDocPropsApp.Root.Clear;
  FDocPropsCore.Clear;

  FSavedRoot.Clear;
  FSavedWorkbook.Clear;
  ClearSavedSheets;
  FSavedSheets.Clear;

  FOPC.Clear;

  SetDefaultData;
end;

procedure TXc12FileData.ClearSavedSheets;
var
  i: integer;
begin
  for i := 0 to FSavedSheets.Count - 1 do
    TXLSSavedFileDataSheet(FSavedSheets.Objects[i]).Free;
end;

procedure TXc12FileData.CommitSaveToStream;
begin
  if FSavedWorkbook.FindByType(OPC_XLSX_THEME) = Nil then
    WriteTheme;
  WriteUnusedData;
  FOPC.Close;
end;

constructor TXc12FileData.Create;
begin
{$ifdef MSWINDOWS}
  FUseAlternateZip := False;
{$else}
  FUseAlternateZip := True;
{$endif}

  FOPC           := TOPC_XLSX.Create;
  FContentType   := TXPGDocContentType.Create;
  FDocPropsApp   := TXPGDocDocPropsApp.Create;
  FDocPropsCore  := TXpgSimpleDOM.Create;

  FSavedRoot     := TXLSSavedFileDataList.Create;
  FSavedWorkbook := TXLSSavedFileDataList.Create;
{$ifdef DELPHI_5}
  FSavedSheets   := TStringList.Create;
{$else}
  FSavedSheets   := THashedStringList.Create;
{$endif}

  SetDefaultData;
end;

destructor TXc12FileData.Destroy;
begin
  FOPC.Free;
  FContentType.Free;
  FDocPropsApp.Free;
  FDocPropsCore.Free;
  FSavedRoot.Free;
  FSavedWorkbook.Free;
  ClearSavedSheets;
  FSavedSheets.Free;

  inherited;
end;

procedure TXc12FileData.LoadFromStream(AZIPStream: TStream);
begin
  FOPC.OpenRead(AZIPStream,FUseAlternateZip);

  ReadContentTypes;
  ReadDocPropsCore;
end;

procedure TXc12FileData.ReadContentTypes;
var
  Stream: TStream;
begin
  Stream := FOPC.ReadContentType;
  if Stream <> Nil then begin
    FContentType.LoadFromStream(Stream);
    Stream.Free;
  end;
end;

procedure TXc12FileData.ReadDocPropsApp;
var
  Stream: TStream;
begin
  Stream := FOPC.ReadDocPropsApp;
  try
    FDocPropsApp.LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TXc12FileData.ReadDocPropsCore;
var
  Stream: TStream;
begin
  Stream := FOPC.ReadDocPropsCore;
  if Stream <> Nil then begin
    FDocPropsCore.LoadFromStream(Stream);
    Stream.Free;
  end;
end;

procedure TXc12FileData.ReadUnusedData;
var
  i: integer;
  List: TOPCItemList;
  SaveList: TXLSSavedFileDataList;
  SavedSheets: TXLSSavedFileDataSheet;

procedure AddSheet(AId: AxUCString);
var
  i: integer;
begin
  i := FSavedSheets.IndexOf(AId);
  if i < 0 then
    raise Exception.CreateFmt('Unknown OPC sheet id "%s"',[AId]);
  SavedSheets := TXLSSavedFileDataSheet(FSavedSheets.Objects[i]);
  if SavedSheets = Nil then begin
    SavedSheets := TXLSSavedFileDataSheet.Create;
    FSavedSheets.Objects[i] := SavedSheets;
  end;
end;

begin
  List := TOPCItemList.Create;
  try
    FOPC.SaveUnchecked(List);

    for i := 0 to List.Count - 1 do begin
      if (List[i].TargetMode <> otmInternal) or (List[i].Name = 'calcChain.xml') or (List[i].Name = 'app.xml')  or (Copy(List[i].Name,1,15) = 'printerSettings') then
        Continue;
      if List[i].Parent.Type_ = OPC_XLSX_WORKSHEET then begin
        AddSheet(List[i].Parent.Id);
        if List[i].Type_ = OPC_XLSX_TABLE then
          SaveList := SavedSheets.TableData
        else
          SaveList := SavedSheets.SheetData;
      end
      else if List[i].Parent.Type_ = OPC_XLSX_WORKBOOK then
        SaveList := FSavedWorkbook
      else if List[i].Parent.Type_ = OPC_ROOT then
        SaveList := FSavedRoot
      else
        SaveList := FSavedRoot;
//        raise Exception.CreateFmt('Unknown OPC type "%s" in save data',[List[i].Type_]);
     SaveList.Add(List[i].Type_,List[i].Target,List[i].Content,List[i].Data);
     List[i].Data := Nil;
    end;
  finally
    List.Free;
  end;
end;

procedure TXc12FileData.SetDefaultData;
begin
  FDocPropsApp.Root.RootAttributes.AddNameValue('xmlns',OOXML_URI_OFFICEDOC_EXTENDED_PROPERIES);
  FDocPropsApp.Root.RootAttributes.AddNameValue('xmlns:vt',OOXML_URI_OFFICEDOC_EXTENDED_DOCPROPSVTYPES);
  FDocPropsCore.LoadFromString(XLS_DEFAULT_DOCPROPS_CORE);
end;

procedure TXc12FileData.SetUseAlternateZip(const Value: boolean);
begin
// Must use built in zip when not windows.
{$ifdef MSWINDOWS}
  FUseAlternateZip := Value;
{$endif}
end;

function TXc12FileData.StreamByName(const AName: AxUCString): TStream;
var
  i: integer;
begin
  for i := 0 to FSavedWorkbook.Count - 1 do begin
    if FSavedWorkbook[i].Target = AName then begin
      Result := FSavedWorkbook[i].Data;
      Exit;
    end;
  end;
  Result := Nil;
end;

procedure TXc12FileData.WriteDocPropsCore;
var
  Stream: TMemoryStream;
begin
  Stream := TMemoryStream.Create;
  try
    FDocPropsCore.SaveToStream(Stream);
    FOPC.AddDocPropsCore(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TXc12FileData.WriteTheme;
var
  Stream: TStringStream;
begin
  Stream := TStringStream.Create('');
  try
    Stream.WriteString(XLS_DEFAULT_THEME);
    FOPC.AddTheme(Stream,1);
  finally
    Stream.Free;
  end;
end;

procedure TXc12FileData.WriteUnusedData;
var
  i: integer;

procedure SaveData(AParent: TOPCItem; AData: TXLSSavedFileData);
var
  OPC: TOPCItem;
begin
  OPC := FOPC.ItemCreate(AParent,AData.Type_,AData.Target,AData.Content);
  FOPC.ItemWrite(OPC,AData.Data);
  FOPC.ItemClose(OPC);
end;

begin
  for i := 0 to FSavedRoot.Count - 1 do
    SaveData(FOPC.Root,FSavedRoot[i]);
  for i := 0 to FSavedWorkbook.Count - 1 do
    SaveData(FOPC.Workbook,FSavedWorkbook[i]);

// Saved below
//  for i := 0 to FSavedSheets.Count - 1 do begin
//    if FSavedSheets.Objects[i] <> Nil then begin
//      List := TXLSSavedFileDataSheet(FSavedSheets.Objects[i]).SheetData;
//      OPCSheet := FOPC.FindSheet(i + 1);
//      if OPCSheet <> Nil then begin
//        for j := 0 to List.Count - 1 do
//          SaveData(OPCSheet,List[j]);
//      end;
//    end;
//  end;
end;

procedure TXc12FileData.WriteUnusedDataSheet(AIndex: integer; AOPCSheet: TOPCItem);
var
  i: integer;
  List: TXLSSavedFileDataList;
  OPC: TOPCItem;
begin
  if (AIndex >= 0) and (AIndex < FSavedSheets.Count) then begin
    if FSavedSheets.Objects[AIndex] <> Nil then begin

      List := TXLSSavedFileDataSheet(FSavedSheets.Objects[AIndex]).SheetData;
      for i := 0 to List.Count - 1 do begin
        OPC := FOPC.ItemCreate(AOPCSheet,List[i].Type_,List[i].Target,List[i].Content);
        FOPC.ItemWrite(OPC,List[i].Data);
        FOPC.ItemClose(OPC);
      end;

      List := TXLSSavedFileDataSheet(FSavedSheets.Objects[AIndex]).TableData;
      for i := 0 to List.Count - 1 do begin
        OPC := FOPC.ItemCreate(AOPCSheet,List[i].Type_,List[i].Target,List[i].Content);
        FOPC.ItemWrite(OPC,List[i].Data);
        FOPC.ItemClose(OPC);
      end;
    end;
  end;
end;

{ TXLSSavedFileData }

destructor TXLSSavedFileData.Destroy;
begin
  if FData <> Nil then
    FData.Free;
  inherited;
end;

{ TXLSSavedFileDataList }

procedure TXLSSavedFileDataList.Add(AType, ATarget,AContent: AxUCString; AData: TStream);
var
  Item: TXLSSavedFileData;
begin
  Item := TXLSSavedFileData.Create;
  Item.Type_ := AType;
  Item.Target := ATarget;
  Item.Content := AContent;
  Item.Data := AData;
  inherited Add(Item);
end;

constructor TXLSSavedFileDataList.Create;
begin
  inherited Create;
end;

function TXLSSavedFileDataList.FindByType(AType: AxUCString): TXLSSavedFileData;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].Type_ = AType then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXLSSavedFileDataList.GetItems(Index: integer): TXLSSavedFileData;
begin
  Result := TXLSSavedFileData(inherited Items[Index]);
end;

{ TXLSSavedFileDataSheet }

constructor TXLSSavedFileDataSheet.Create;
begin
  FSheetData := TXLSSavedFileDataList.Create;
  FTableData := TXLSSavedFileDataList.Create;
end;

destructor TXLSSavedFileDataSheet.Destroy;
begin
  FSheetData.Free;
  FTableData.Free;

  inherited;
end;

{ TXLSPivotTables }

function TXLSPivotTables.Add: TXLSPivotTable;
begin
  Result := TXLSPivotTable.Create;

  inherited Add(Result);
end;

constructor TXLSPivotTables.Create;
begin
  inherited Create;
end;

function TXLSPivotTables.GetItems(const Index: integer): TXLSPivotTable;
begin
  Result := TXLSPivotTable(inherited Items[Index]);
end;

{ TXLSPivotCacheDefinitions }

function TXLSPivotCacheDefinitions.Add: TXLSPivotCacheDefinition;
begin
  Result := TXLSPivotCacheDefinition.Create;
  inherited Add(Result);
end;

constructor TXLSPivotCacheDefinitions.Create;
begin
  inherited Create;
end;

destructor TXLSPivotCacheDefinitions.Destroy;
begin
  inherited;
end;

function TXLSPivotCacheDefinitions.Find(const AItem: TXLSPivotCacheDefinition): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result] = AItem then
      Exit;
  end;
  Result := -1;
end;

function TXLSPivotCacheDefinitions.Find(const ACacheId: integer): TXLSPivotCacheDefinition;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].FCacheId = ACacheId then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXLSPivotCacheDefinitions.GetItems(const Index: integer): TXLSPivotCacheDefinition;
begin
  Result := TXLSPivotCacheDefinition(inherited Items[Index]);
end;

procedure TXLSPivotCacheDefinitions.Release(AItem: TXLSPivotCacheDefinition);
var
  i: integer;
begin
  i := Find(AItem);
  if i >= 0 then begin
    AItem.Release;
    if AItem.FUsageCount <= 0 then
      Delete(i);
  end;
end;

{ TXLSPivotCacheDefinition }

constructor TXLSPivotCacheDefinition.Create;
begin
  FDOM := TXpgSimpleDOM.Create;
  FRecords := TXLSPivotCacheRecords.Create;
end;

destructor TXLSPivotCacheDefinition.Destroy;
begin
  FDOM.Free;
  FRecords.Free;

  inherited;
end;

procedure TXLSPivotCacheDefinition.Release;
begin
  Dec(FUsageCount);
end;

procedure TXLSPivotCacheDefinition.SetSource(const AArea, ASheet: AxUCString);
var
  Node: TXpgDOMNode;
begin
  Node := FDOM.Root.Find('pivotCacheDefinition/cacheSource/worksheetSource',False);
  if Node <> Nil then begin
    Node.Attributes.Update('ref',AArea);
    Node.Attributes.Update('sheet',ASheet);
  end;
end;

procedure TXLSPivotCacheDefinition.Use;
begin
  Inc(FUsageCount);
end;

{ TXLSPivotCacheRecords }

constructor TXLSPivotCacheRecords.Create;
begin
  FStream := TMemoryStream.Create;
end;

destructor TXLSPivotCacheRecords.Destroy;
begin
  FStream.Free;

  inherited;
end;

{ TXLSPivotTable }

constructor TXLSPivotTable.Create;
begin
  FDOM := TXpgSimpleDOM.Create;
end;

destructor TXLSPivotTable.Destroy;
begin
  FDOM.Free;

  inherited;
end;

end.
