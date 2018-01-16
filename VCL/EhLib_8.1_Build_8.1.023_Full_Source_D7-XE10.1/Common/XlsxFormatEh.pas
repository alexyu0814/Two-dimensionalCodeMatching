{*******************************************************}
{                                                       }
{                      EhLib v8.1                       }
{                                                       }
{           Classes to work with Xlsx Format            }
{                     Build 8.1.05                      }
{                                                       }
{      Copyright (c) 2015-15 by Dmitry V. Bolshakov     }
{                                                       }
{*******************************************************}

unit XlsxFormatEh;

{$I EhLib.Inc}

interface

uses
  Windows, SysUtils, Classes, Graphics, Dialogs, GridsEh, Controls,
{$IFDEF EH_LIB_6} Variants, {$ENDIF}
{$IFDEF EH_LIB_16} System.Zip, {$ENDIF}
{$IFDEF FPC}
  XMLRead, DOM, EhLibXmlConsts,
{$ELSE}
  xmldom, XMLIntf, XMLDoc, EhLibXmlConsts,
{$ENDIF}

{$IFDEF EH_LIB_17} System.UITypes, System.Contnrs, {$ENDIF}
{$IFDEF CIL}
  EhLibVCLNET,
  System.Runtime.InteropServices, System.Reflection,
{$ELSE}
  {$IFDEF FPC}
    EhLibLCL, LCLType, Win32Extra,
  {$ELSE}
    EhLibVCL, Messages, SqlTimSt,
  {$ENDIF}
{$ENDIF}
  Db, Clipbrd, ComObj, StdCtrls;

{$IFDEF FPC}
{$ELSE}

type

  TCustomFileZipingProviderEh = class;

  TXlsxCellDataEh = class(TPersistent)
  private
    FColor: TColor;
    FFont: TFont;
    FValue: Variant;
    FHorzLine: Boolean;
    FVertLine: Boolean;
    FDisplayText: String;
    FDisplayFormat: String;
    FAlignment: TAlignment;
  public
    constructor Create;
    destructor Destroy; override;

    procedure Clear;

    property Value: Variant read FValue write FValue;
    property DisplayText: String read FDisplayText write FDisplayText;
    property DisplayFormat: String read FDisplayFormat write FDisplayFormat;
    property Color: TColor read FColor write FColor;
    property Font: TFont read FFont;
    property HorzLine: Boolean read FHorzLine write FHorzLine;
    property VertLine: Boolean read FVertLine write FVertLine;
    property Alignment: TAlignment read FAlignment write FAlignment;
  end;

 { TXlsxFileWriterEh }

  RStyle = record
    NumFmt: Integer;
    NumFont: Integer;
    NumAlignment: Integer;
    NumBackColor: Integer;
    NumBorder: Integer;
    Vertical: Boolean;
    Wrap: Boolean;
  end;

  TXlsxFileWriterEh = class(TPersistent)
  private
    FBackColor: array of TColor;
    FBorder: array of Integer;
    FCurCol: Integer;
    FCurRow: Integer;
    FDataRowCount: Integer;
    FFmtNode: IXMLNode;
    FFonts: array of TFont;
    FNumFmt: Integer;
    FRowNode: IXMLNode;
    FStyles: array of RStyle;
    FXMLSheet: IXMLDocument;
    FXMLStyles: IXMLDocument;
    FZipFileProvider: TCustomFileZipingProviderEh;
    FColCount: Integer;
    FCellData: TXlsxCellDataEh;
    FFileName: String;
    function GetCurColNum: Integer;
    function GetCurRowNum: Integer;
  protected
    function GetColWidth(ACol: Integer): Integer; virtual;
    function InitTableCheckEof: Boolean; virtual;
    function InitRowCheckEof(ARow: Integer): Boolean; virtual;
    function GetColCount: Integer; virtual;
    function SysVarToStr(var Val: Variant): String;

    function SetBackColor(Color: TColor): Integer;
    function SetBorder(Border: Integer): Integer;
    function SetFont(Font: TFont): Integer;
    function SetStyle(NumFmt: Integer; Font: TFont; Alignment: Integer; BackColor: TColor; Border: Integer; Vert, Wrap: Boolean): Integer;

    procedure InitFileData; virtual;
    procedure CreateXMLs; virtual;
    procedure CreateStaticXMLs; virtual;
    procedure CreateDynamicXMLs; virtual;

    procedure WriteXMLs; virtual;
    procedure WriteSheetXML; virtual;
    procedure WriteColumnsSec(Root: IXMLNode); virtual;
    procedure WriteStylesXML; virtual;
    procedure SaveDataToFile; virtual;

    procedure WriteRow(ARow: Integer; var Eof: Boolean); virtual;
    procedure WriteValue(ACol, ARow: Integer); virtual;
    procedure GetCellData(ACol, ARow: Integer; CellData: TXlsxCellDataEh); virtual;

  public
    constructor Create;
    destructor Destroy; override;

    procedure WritetToFile(const AFileName: String); virtual;
    property CurRowNum: Integer read GetCurRowNum;
    property CurColNum: Integer read GetCurColNum;
  end;

{ TCustomFileZipingProviderEh }

  TCustomFileZipingProviderEh = class(TObject)
  public
    class function CreateInstance: TCustomFileZipingProviderEh; virtual;
    function InitZipFile(FileName: String): TStream; virtual; abstract;
    procedure AddFile(Data: TStream; const FilePathAndName: string); virtual; abstract;
    procedure FinalizeZipFile; overload; virtual; abstract;
    procedure FinalizeZipFile(AStream: TStream); overload; virtual; abstract;
  end;

  TCustomFileZipingProviderEhClass = class of TCustomFileZipingProviderEh;

{$IFDEF EH_LIB_16} 

{ TSystemZipFileProvider }

  TSystemZipFileProviderEh = class(TCustomFileZipingProviderEh)
  private
    FStream: TStream;
    FZipFile: TZipFile;
  public
    constructor Create;
    destructor Destroy; override;

    class function CreateInstance: TCustomFileZipingProviderEh; override;
    function InitZipFile(FileName: String): TStream; override;
    procedure AddFile(Data: TStream; const FilePathAndName: string); override;
    procedure FinalizeZipFile; override;
    procedure FinalizeZipFile(AStream: TStream); override;

    property Stream: TStream read FStream write FStream;
  end;

{$ENDIF} 

function ZipFileProviderClass: TCustomFileZipingProviderEhClass;
function RegisterZipFileProviderClass(AZipFileProviderClass: TCustomFileZipingProviderEhClass): TCustomFileZipingProviderEhClass;

function ZEGetA1ByCol(ColNum: Integer; StartZero: Boolean = True): string;
function ZEGetColByA1(const A1Str: String; StartZero: Boolean = True): Integer;
function ZEA1RectToRect(const A1Rect: String): TRect;
function ZEA1CellToPoint(const A1CellRef: String): TPoint;
function XlsxNumFormatToVCLFormat(const XlsxNumFormatId, XlsxNumFormat: String): String;
function XlsxCellTypeNameToVaType(const tn: String): TVarType;

{$ENDIF} // not Lazarus

implementation

{$IFDEF FPC}
{$ELSE}

var
  FZipFileProviderClass: TCustomFileZipingProviderEhClass;

const
  ZE_STR_ARRAY: array [0..25] of char = (
  'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N',
  'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');

var
  NumberFormatArr: array[0..49] of String =
  (
    '', //0 General
    '0', //1 0
    '0.00', //2 0.00
    '#,##0', //3 #,##0
    '#,##0.00', //4 #,##0.00
    '', //5
    '', //6
    '', //7
    '', //8
    '0%', //9 0%
    '0.00%', //10 0.00%
    '0.00E+00', //11 0.00E+00
    '# ?/?', //12 # ?/?
    '# ??/??', //13 # ??/??
    'mm-dd-yy', //14 mm-dd-yy
    'd-mmm-yy', //15 d-mmm-yy
    'd-mmm', //16 d-mmm
    'mmm-yy', //17 mmm-yy
    'h:nn AM/PM', //18 h:mm AM/PM
    'h:nn:ss AM/PM', //19 h:mm:ss AM/PM
    'h:nn', //20 h:mm
    'h:nn:ss', //21 h:mm:ss
    'm/d/yy h:nn', //22 m/d/yy h:mm
    '', //23
    '', //24
    '', //25
    '', //26
    '', //27
    '', //28
    '', //29
    '', //30
    '', //31
    '', //32
    '', //33
    '', //34
    '', //35
    '', //36
    '#,##0; (#,##0)', //37 #,##0 ;(#,##0)
    '#,##0;[Red](#,##0)', //38 #,##0 ;[Red](#,##0)
    '#,##0.00;(#,##0.00)', //39 #,##0.00;(#,##0.00)
    '#,##0.00;[Red](#,##0.00)', //40 #,##0.00;[Red](#,##0.00)
    '', //41
    '', //42
    '', //43
    '', //44
    'nn:ss', //45 mm:ss
    '[h]:nn:ss', //46 [h]:mm:ss
    'nnss.0', //47 mmss.0
    '##0.0E+0', //48 ##0.0E+0
    '@' //49 @
  );

const
  SecretSystLDateFormatCode = 'F800';

function ConvertXlsxNumFormatToVCLFormat(const XlsxNumFormat: String): String;
var
  i: Integer;
  InBrakets: Boolean;
  EndOfS, ElapsedSign: String;
  LocaleCode: String;
//  ElapsedHours, ElapsedMins, ElapsedSecs: Boolean;
begin
  Result := '';
  i := 1;
  InBrakets := False;
  ElapsedSign := '';
  LocaleCode := '';
//  ElapsedHours := False;
//  ElapsedMins := False;
//  ElapsedSecs := False;
  while i <= Length(XlsxNumFormat) do
  begin
    if InBrakets then
    begin
      if XlsxNumFormat[i] = ']' then
      begin
        InBrakets := False;
        Result := Result + ElapsedSign;
        ElapsedSign := '';
      end else if (XlsxNumFormat[i] = 'h') and ( (Length(ElapsedSign) = 0) or (ElapsedSign[1] = 'h') ) then
        ElapsedSign := ElapsedSign + 'h'
      else if (XlsxNumFormat[i] = 'm') and ((Length(ElapsedSign) = 0) or (ElapsedSign[1] = 'm')) then
        ElapsedSign := ElapsedSign + 'n'
      else if (XlsxNumFormat[i] = 's') and ((Length(ElapsedSign) = 0) or (ElapsedSign[1] = 's')) then
        ElapsedSign := ElapsedSign + 's'
      else
        ElapsedSign := '';

      if InBrakets and CharInSetEh(XlsxNumFormat[i], ['A','B','C','D','E','F','0','1','2','3','4','5','6','7','8','9']) then
        LocaleCode := LocaleCode + XlsxNumFormat[i];

      i := i + 1;
      Continue;
    end;
    if XlsxNumFormat[i] = '[' then
    begin
      i := i + 1;
      InBrakets := True;
      Continue;
    end;
    if CharInSetEh(XlsxNumFormat[i], ['_', '*']) then
    begin
      i := i + 2;
      Continue;
    end;
    if XlsxNumFormat[i] = '\' then
    begin
      i := i + 1;
      if i <= Length(XlsxNumFormat) then
      begin
        Result := Result + '"'+XlsxNumFormat[i]+'"';
        i := i + 1;
      end;
      Continue;
    end;
    Result := Result + XlsxNumFormat[i];
    i := i + 1;
  end;

  EndOfS := Copy(Result, Length(Result)-1, 2);
  if EndOfS = ';@' then
    Result := Copy(Result, 1, Length(Result)-2);
  if LocaleCode = SecretSystLDateFormatCode then
    Result := FormatSettings.LongDateFormat;
end;

function XlsxNumFormatToVCLFormat(const XlsxNumFormatId, XlsxNumFormat: String): String;
begin
//  if StrToInt(XlsxNumFormatId) = 164 then
//    Result := FormatSettings.LongDateFormat
//  else
  if XlsxNumFormat <> '' then
    Result := ConvertXlsxNumFormatToVCLFormat(XlsxNumFormat)
  else if StrToInt(XlsxNumFormatId) = 14 then
    Result := FormatSettings.ShortDateFormat
  else if StrToInt(XlsxNumFormatId) <= 49 then
    Result := NumberFormatArr[StrToInt(XlsxNumFormatId)]
  else
    Result := '';
end;

function XlsxCellTypeNameToVaType(const tn: String): TVarType;
var
  StrType: TVarType;
begin
{$IFDEF EH_LIB_12}
  StrType := varUString;
{$ELSE}
  StrType := varOleStr;
{$ENDIF}
  if tn = 'b' then
    Result := varBoolean
  else if tn = 'd' then
    Result := varDate
  else if tn = 'e' then
    Result := StrType
  else if tn = 'inlineStr' then
    Result := StrType
  else if tn = 'n' then
    Result := varDouble
  else if tn = 's' then
    Result := StrType
  else if tn = 'str' then
    Result := StrType
  else
    Result := varDouble
end;

function ZEA1RectToRect(const A1Rect: String ): TRect;
var
  i: Integer;
  s1, s2: String;
  Done: Boolean;
begin
  Result := EmptyRect;
  Done := False;
  for i := 1 to Length(A1Rect) do
  begin
    if A1Rect[i] = ':' then
    begin
      s1 := Copy(A1Rect, 1, i-1);
      Result.TopLeft := ZEA1CellToPoint(s1);
      s2 := Copy(A1Rect, i+1, Length(A1Rect));
      Result.BottomRight := ZEA1CellToPoint(s2);
      Done := True;
      Break;
    end;
  end;
  if not Done then
  begin
    Result.TopLeft := ZEA1CellToPoint(A1Rect);
    Result.BottomRight := ZEA1CellToPoint(A1Rect);
  end;
end;

function ZEA1CellToPoint(const A1CellRef: String): TPoint;
var
  i: Integer;
  sx, sy: String;
begin
  Result.X := 0;
  Result.Y := 0;
  for i := 1 to Length(A1CellRef) do
  begin
    if CharInSetEh(A1CellRef[i], ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']) then
    begin
      sx := Copy(A1CellRef, 1, i-1);
      sy := Copy(A1CellRef, i, Length(A1CellRef));
      Result.X := ZEGetColByA1(sx);
      Result.Y := StrToInt(sy)-1;
      Break;
    end;
  end;
end;

function ZEGetColByA1(const A1Str: String; StartZero: Boolean = True): Integer;
var
  i: Integer;
  c: Char;
  Mul: Integer;
begin
  Result := -1;
  Mul := -1;
  for i := Length(A1Str) downto 1 do
  begin
    if CharInSetEh(A1Str[i], ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']) then
      Continue;
    c := UpperCase(A1Str[i])[1];
    case c of
      'A': Mul := 1;
      'B': Mul := 2;
      'C': Mul := 3;
      'D': Mul := 4;
      'E': Mul := 5;
      'F': Mul := 6;
      'G': Mul := 7;
      'H': Mul := 8;
      'I': Mul := 9;
      'J': Mul := 10;
      'K': Mul := 11;
      'L': Mul := 12;
      'M': Mul := 13;
      'N': Mul := 14;
      'O': Mul := 15;
      'P': Mul := 16;
      'Q': Mul := 17;
      'R': Mul := 18;
      'S': Mul := 19;
      'T': Mul := 20;
      'U': Mul := 21;
      'V': Mul := 22;
      'W': Mul := 23;
      'X': Mul := 24;
      'Y': Mul := 25;
      'Z': Mul := 26;
    end;
    if Result = -1 then
      Result := Mul
    else
      Result := Mul * 26 + Result;
  end;
  Result := Result - 1;
end;

function ZEGetA1byCol(ColNum: Integer; StartZero: Boolean = True): string;
var
  t, n: Integer;
  S: string;
begin
  t := ColNum;
  if (not StartZero) then
    Dec(t);
  Result := '';
  S := '';
  while t >= 0 do
  begin
    n := t mod 26;
    t := (t div 26) - 1;
    S := S + ZE_STR_ARRAY[n];
  end;
  for t := Length(S) downto 1 do
    Result := Result + S[t];
end;

function IsNumber(const st: String): Boolean;
var
  b: Boolean;
  i, n: Integer;
begin
  b := True;
  n := Length(st);
  for i := 1 to n do
    if not CharInSetEh(st[i], ['#', '0', ',', '.']) then
      b := False;
  Result := b;
end;

function ColorToHex(Acolor: TColor): String;
begin
  Result :=
    IntToHex(GetRValue(ColorToRGB(AColor)), 2) +
    IntToHex(GetGValue(ColorToRGB(AColor)), 2) +
    IntToHex(GetBValue(ColorToRGB(AColor)), 2);
end;

{ TDBGridEhExportAsXlsx }

procedure InitXMLDocument(var XML: IXMLDocument);
begin
  XML.Options := XML.Options-[doNamespaceDecl];
  XML.Encoding := 'UTF-8';
  XML.StandAlone := 'yes';
end;

procedure ToTheme(var Stream: TStream);
var
  XML: IXMLDocument;
begin
  XML := NewXMLDocument;
  XML.LoadFromXML(SThemeEh);
  XML.SaveToStream(Stream);
end;

procedure ToRels(var Stream: TStream);
const
  st = 'http://schemas.openxmlformats.org/package/2006/relationships';
  ArrId: array [1 .. 3] of string = ('rId3', 'rId2', 'rId1');
  ArrType: array [1 .. 3] of string =
    ('http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
     'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
     'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument');
  ArrTarget: array [1 .. 3] of string = (
    'docProps/app.xml',
    'docProps/core.xml',
    'xl/workbook.xml');
var
  i: Integer;
  Xml: IXMLDocument;
  Root: IXMLNode;
begin
  Xml := NewXMLDocument;

  InitXMLDocument(Xml);
  Xml.AddChild('Relationships').Attributes['xmlns'] := st;
  for i := 1 to 3 do
  begin
    Root := Xml.CreateElement('Relationship', st);
    Root.Attributes['Id'] := ArrId[i];
    Root.Attributes['Type'] := ArrType[i];
    Root.Attributes['Target'] := ArrTarget[i];
    Xml.DocumentElement.ChildNodes.Add(Root);
  end;
  Xml.SaveToStream(Stream);
end;

procedure ToRelsWorkbook(var Stream: TStream);
const
  st = 'http://schemas.openxmlformats.org/package/2006/relationships';
  ArrId: array [1 .. 3] of string = ('rId3', 'rId2', 'rId1');
  ArrType: array [1 .. 3] of string =
    ('http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
  ArrTarget: array [1 .. 3] of string = ('styles.xml', 'theme/theme1.xml',
    'worksheets/sheet1.xml');
var
  i: Integer;
  Xml: IXMLDocument;
  Root: IXMLNode;
begin
  Xml := NewXMLDocument;

  InitXMLDocument(Xml);
  Xml.AddChild('Relationships').Attributes['xmlns'] := st;
  for i := 1 to 3 do
  begin
    Root := Xml.CreateElement('Relationship', st);
    Root.Attributes['Id'] := ArrId[i];
    Root.Attributes['Type'] := ArrType[i];
    Root.Attributes['Target'] := ArrTarget[i];
    Xml.DocumentElement.ChildNodes.Add(Root);
  end;
  Xml.SaveToStream(Stream);
end;

procedure ToApp(var Stream: TStream);
const
  st = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
  ArrElement: array [1 .. 8] of string = ('Application', 'DocSecurity',
    'ScaleCrop', 'Company', 'LinksUpToDate', 'SharedDoc', 'HyperlinksChanged',
    'AppVersion');
  ArrText: array [1 .. 8] of string = ('Microsoft Excel', '0', 'false',
    'SPecialiST RePack, SanBuild', 'false', 'false', 'false', '14.0300');
var
  i: Integer;
  Xml: IXMLDocument;
  Root, Node: IXMLNode;
begin
  Xml := NewXMLDocument;
  InitXMLDocument(Xml);
  Root := Xml.AddChild('Properties');
  Root.Attributes['xmlns'] := st;
  Root.Attributes['xmlns:vt'] := 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';

  for i := 1 to 8 do
  begin
    Root := Xml.CreateElement(ArrElement[i], st);
    Root.Text := ArrText[i];
    Xml.DocumentElement.ChildNodes.Add(Root);
  end;

  Root := Xml.CreateElement('HeadingPairs', st);
  Node := Root.AddChild('vt:vector');
  Node.Attributes['size'] := 2;
  Node.Attributes['baseType'] := 'variant';
  Node.AddChild('vt:variant').AddChild('vt:lpstr').Text := 'Sheets';
  Node.AddChild('vt:variant').AddChild('vt:i4').Text := '1';
  Xml.DocumentElement.ChildNodes.Add(Root);

  Root := Xml.CreateElement('TitlesOfParts', st);
  Node := Root.AddChild('vt:vector');
  Node.Attributes['size'] := 1;
  Node.Attributes['baseType'] := 'lpstr';
  Node.AddChild('vt:lpstr').Text := 'Sheet1';
  Xml.DocumentElement.ChildNodes.Add(Root);

  Xml.SaveToStream(Stream);
end;

procedure ToCore(var Stream: TStream);
const
  st = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
  ArrAttribut: array [1 .. 5] of string = ('xmlns:cp', 'xmlns:dc',
    'xmlns:dcterms', 'xmlns:dcmitype', 'xmlns:xsi');
  ArrTextAttribut: array [1 .. 5] of string = (st,
    'http://purl.org/dc/elements/1.1/', 'http://purl.org/dc/terms/',
    'http://purl.org/dc/dcmitype/',
    'http://www.w3.org/2001/XMLSchema-instance');
  ArrCreator: array [1 .. 3] of string = ('dc:creator', 'cp:lastModifiedBy',
    'dcterms:created');
var
  i: Integer;
  Xml: IXMLDocument;
  Root: IXMLNode;
begin
  Xml := NewXMLDocument;
  InitXMLDocument(Xml);
  Root := Xml.AddChild('cp:coreProperties');
  for i := 1 to 5 do
    Root.Attributes[ArrAttribut[i]] := ArrTextAttribut[i];

  for i := 1 to 3 do
  begin
    Root := Xml.CreateElement(ArrCreator[i], '');
    if i < 3 then
      Root.Text := 'EhLib'
    else
      Root.Text := FormatDateTime('yyyy-mm-dd', Date) + 'T' +
        FormatDateTime('hh:mm:ss', Time) + 'Z';
    Xml.DocumentElement.ChildNodes.Add(Root);
  end;
  Root.Attributes['xsi:type'] := 'dcterms:W3CDTF';

  Xml.SaveToStream(Stream);
end;

procedure ToWorkbook(var Stream: TStream);
const
  st = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  ArrAttributes: array [1 .. 11] of string = ('xl', '5', '5', '9303', '480',
    '120', '27795', '12585', 'Sheet1', '1', 'rId1');
  ArrNameAttributes: array [1 .. 11] of string = ('appName', 'lastEdited',
    'lowestEdited', 'rupBuild', 'xWindow', 'yWindow', 'windowWidth',
    'windowHeight', 'name', 'sheetId', 'r:id');
var
  i: Integer;
  Xml: IXMLDocument;
  Root, Node: IXMLNode;
begin
  Xml := NewXMLDocument;
  InitXMLDocument(Xml);
  Root := Xml.AddChild('workbook');
  Root.Attributes['xmlns'] := st;
  Root.Attributes['xmlns:r'] :=
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  Root := Xml.CreateElement('fileVersion', st);
  for i := 1 to 4 do
    Root.Attributes[ArrNameAttributes[i]] := ArrAttributes[i];
  Xml.DocumentElement.ChildNodes.Add(Root);
  Root := Xml.CreateElement('workbookPr', st);
  Root.Attributes['defaultThemeVersion'] := '124226';
  Xml.DocumentElement.ChildNodes.Add(Root);
  Root := Xml.CreateElement('bookViews', st);
  Node := Root.AddChild('workbookView');
  for i := 5 to 8 do
    Node.Attributes[ArrNameAttributes[i]] := ArrAttributes[i];
  Xml.DocumentElement.ChildNodes.Add(Root);
  Root := Xml.CreateElement('sheets', st);
  Node := Root.AddChild('sheet');
  for i := 9 to 11 do
    Node.Attributes[ArrNameAttributes[i]] := ArrAttributes[i];
  Xml.DocumentElement.ChildNodes.Add(Root);
  Root := Xml.CreateElement('calcPr', st);
  Root.Attributes['calcId'] := '145621';
  Xml.DocumentElement.ChildNodes.Add(Root);

  Xml.SaveToStream(Stream);
end;

procedure ToContentTypes(var Stream: TStream);
const
  st = 'http://schemas.openxmlformats.org/package/2006/content-types';
  ArrExtension: array [1 .. 8] of string = ('rels', 'xml', '/xl/workbook.xml',
    '/xl/worksheets/sheet1.xml', '/xl/theme/theme1.xml', '/xl/styles.xml',
    '/docProps/core.xml', '/docProps/app.xml');
  ArrContent: array [1 .. 8] of string =
    ('application/vnd.openxmlformats-package.relationships+xml',
    'application/xml',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
    'application/vnd.openxmlformats-officedocument.theme+xml',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
    'application/vnd.openxmlformats-package.core-properties+xml',
    'application/vnd.openxmlformats-officedocument.extended-properties+xml');
var
  i: Integer;
  Xml: IXMLDocument;
  Root: IXMLNode;
begin
  Xml := NewXMLDocument;
  InitXMLDocument(Xml);
  Root := Xml.AddChild('Types');
  Root.Attributes['xmlns'] := st;

  for i := 1 to 2 do
  begin
    Root := Xml.CreateElement('Default', st);
    Root.Attributes['Extension'] := ArrExtension[i];
    Root.Attributes['ContentType'] := ArrContent[i];
    Xml.DocumentElement.ChildNodes.Add(Root);
  end;

  for i := 3 to 8 do
  begin
    Root := Xml.CreateElement('Override', st);
    Root.Attributes['PartName'] := ArrExtension[i];
    Root.Attributes['ContentType'] := ArrContent[i];
    Xml.DocumentElement.ChildNodes.Add(Root);
  end;

  Xml.SaveToStream(Stream);
end;

procedure CreateSheetXml(var Xml: IXMLDocument);
const
  st = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  ArrAttributes: array [1 .. 5] of string = (st,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'http://schemas.openxmlformats.org/markup-compatibility/2006', 'x14ac',
    'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');
  ArrNameAttributes: array [1 .. 5] of string = ('xmlns', 'xmlns:r', 'xmlns:mc',
    'mc:Ignorable', 'xmlns:x14ac');
var
  i: Integer;
  Root, Node: IXMLNode;
begin
  Xml := NewXMLDocument;
  InitXMLDocument(Xml);
  Root := Xml.AddChild('worksheet');
  for i := 1 to 5 do
    Root.Attributes[ArrNameAttributes[i]] := ArrAttributes[i];

  Root := Xml.CreateElement('dimension', st);
  Root.Attributes['ref'] := 'A1';
  Xml.DocumentElement.ChildNodes.Add(Root);
  Root := Xml.CreateElement('sheetViews', st);
  Node := Root.AddChild('sheetView');
  Node.Attributes['tabSelected'] := '1';
  Node.Attributes['workbookViewId'] := '0';
  Xml.DocumentElement.ChildNodes.Add(Root);
  Root := Xml.CreateElement('sheetFormatPr', st);
  Root.Attributes['defaultRowHeight'] := '15';
  Root.Attributes['x14ac:dyDescent'] := '0.25';
  Xml.DocumentElement.ChildNodes.Add(Root);
  Root := Xml.CreateElement('cols', st);
  Xml.DocumentElement.ChildNodes.Add(Root);
  Root := Xml.CreateElement('sheetData', st);
  Xml.DocumentElement.ChildNodes.Add(Root);
end;

procedure CreateStylesXml(var Xml: IXMLDocument);
const
  st = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  ArrAttributes: array [1 .. 25] of string = (st,
    'http://schemas.openxmlformats.org/markup-compatibility/2006', 'x14ac',
    'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac', '11', '1',
    'Calibri', '2', '204', 'minor', '0', '0', '0', '0', '0', '0', '0', '0', '0',
    'Normal', '0', '0', '0', 'TableStyleMedium2', 'PivotStyleLight16');
  ArrNameAttributes: array [1 .. 25] of string = ('xmlns', 'xmlns:mc',
    'mc:Ignorable', 'xmlns:x14ac', 'val', 'theme', 'val', 'val', 'val', 'val',
    'numFmtId', 'fontId', 'fillId', 'borderId', 'numFmtId', 'fontId', 'fillId',
    'borderId', 'xfId', 'name', 'xfId', 'builtinId', 'count',
    'defaultTableStyle', 'defaultPivotStyle');
  ArrChild: array [5 .. 15] of string = ('sz', 'color', 'name', 'family',
    'charset', 'scheme', 'left', 'right', 'top', 'bottom', 'diagonal');
var
  i: Integer;
  Root, Node: IXMLNode;
begin
  Xml := NewXMLDocument;
  InitXMLDocument(Xml);
  Root := Xml.AddChild('styleSheet');
  for i := 1 to 1 do
    Root.Attributes[ArrNameAttributes[i]] := ArrAttributes[i];

  Root := Xml.CreateElement('fonts', st);
  Root.Attributes['count'] := '1';
  Node := Root.AddChild('font');
  for i := 5 to 10 do
    Node.AddChild(ArrChild[i]).Attributes[ArrNameAttributes[i]] := ArrAttributes[i];
  Xml.DocumentElement.ChildNodes.Add(Root);
  Root := Xml.CreateElement('fills', st);
  Root.Attributes['count'] := '2';
  Root.AddChild('fill').AddChild('patternFill').Attributes['patternType'] := 'none';
  Root.AddChild('fill').AddChild('patternFill').Attributes['patternType'] := 'gray125';
  Xml.DocumentElement.ChildNodes.Add(Root);
  Root := Xml.CreateElement('borders', st);
  Root.Attributes['count'] := '1';
  Node := Root.AddChild('border');
  for i := 11 to 15 do
    Node.AddChild(ArrChild[i]);
  Xml.DocumentElement.ChildNodes.Add(Root);
  Root := Xml.CreateElement('cellStyleXfs', st);
  Root.Attributes['count'] := '1';
  Node := Root.AddChild('xf');
  for i := 11 to 14 do
    Node.Attributes[ArrNameAttributes[i]] := ArrAttributes[i];
  Xml.DocumentElement.ChildNodes.Add(Root);
  Root := Xml.CreateElement('cellXfs', st);
  Root.Attributes['count'] := '1';
  Node := Root.AddChild('xf');
  for i := 15 to 19 do
    Node.Attributes[ArrNameAttributes[i]] := ArrAttributes[i];
  Xml.DocumentElement.ChildNodes.Add(Root);
  Root := Xml.CreateElement('cellStyles', st);
  Root.Attributes['count'] := '1';
  Node := Root.AddChild('cellStyle');
  for i := 20 to 22 do
    Node.Attributes[ArrNameAttributes[i]] := ArrAttributes[i];
  Xml.DocumentElement.ChildNodes.Add(Root);
  Root := Xml.CreateElement('dxfs', st);
  Root.Attributes['count'] := '0';
  Xml.DocumentElement.ChildNodes.Add(Root);
  Root := Xml.CreateElement('tableStyles', st);
  for i := 23 to 25 do
    Root.Attributes[ArrNameAttributes[i]] := ArrAttributes[i];
  Xml.DocumentElement.ChildNodes.Add(Root);
  Root := Xml.CreateElement('extLst', st);
  Node := Root.AddChild('ext');
  Node.Attributes['uri'] := '{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}';
  Node.Attributes['xmlns:x14'] := 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
  Node.AddChild('x14:slicerStyles').Attributes['defaultSlicerStyle'] := 'SlicerStyleLight1';
  Xml.DocumentElement.ChildNodes.Add(Root);
end;

constructor TXlsxFileWriterEh.Create;
begin
  inherited Create;
  FCellData := TXlsxCellDataEh.Create;
end;

destructor TXlsxFileWriterEh.Destroy;
begin
  FreeAndNil(FCellData);
  inherited Destroy;
end;

procedure TXlsxFileWriterEh.InitFileData;
begin
  if ZipFileProviderClass = nil then
    raise Exception.Create('ZipFileProviderClass is not assigned');
  FZipFileProvider := ZipFileProviderClass.CreateInstance;
  FZipFileProvider.InitZipFile(FFileName);
end;

procedure TXlsxFileWriterEh.CreateXMLs;
begin
  CreateStaticXMLs;
  CreateDynamicXMLs;
end;

procedure TXlsxFileWriterEh.CreateStaticXMLs;
const
  Resours: array [1 .. 7] of procedure(var Stream: TStream) =
    (ToTheme, ToRels, ToApp, ToCore, ToRelsWorkbook, ToWorkbook, ToContentTypes);
  ResoursFile: array [1 .. 7] of string = ('xl/theme/theme1.xml', '_rels/.rels',
  'docProps/app.xml', 'docProps/core.xml', 'xl/_rels/workbook.xml.rels',
  'xl/workbook.xml', '[Content_Types].xml');

var
  i: Integer;
  Res: TStream;
begin

  for i := 1 to 7 do
  begin
    Res := TMemoryStream.Create;
    Resours[i](Res);
    Res.Seek(0, soFromBeginning);
    FZipFileProvider.AddFile(Res, ResoursFile[i]);
    Res.Free;
  end;
end;

procedure TXlsxFileWriterEh.CreateDynamicXMLs;
begin
  CreateSheetXml(FXMLSheet);

  SetLength(FFonts, 1);
  FFonts[0] := TFont.Create;
  FFonts[0].Name := 'Calibri';
  FFonts[0].Size := 11;

  SetLength(FBackColor, 2);
  FBackColor[0] := 0;
  FBackColor[1] := 0;

  SetLength(FBorder, 1);
  FBorder[0] := 0;

  SetLength(FStyles, 1);
  FStyles[0].NumFmt := 0;
  FStyles[0].NumFont := 0;
  FStyles[0].NumAlignment := 0;
  FStyles[0].NumBackColor := 0;
  FStyles[0].NumBorder := 0;
  FStyles[0].Vertical := False;
  FStyles[0].Wrap := False;

  CreateStylesXml(FXMLStyles);
  FNumFmt := 164;
  FFmtNode := FXMLStyles.CreateElement('numFmts',
    'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

  FCurRow := 1;
end;

procedure TXlsxFileWriterEh.WriteXMLs;
begin
  WriteSheetXML;
  WriteStylesXML;
end;

procedure TXlsxFileWriterEh.WriteSheetXML;
const
  ArrAttributes: array [1 .. 6] of string = ('0.7', '0.7', '0.75', '0.75',
    '0.3', '0.3');
  ArrNameAttributes: array [1 .. 6] of string = ('left', 'right', 'top',
    'bottom', 'header', 'footer');
var
  i: Integer;
  Res: TStream;
  Root: IXMLNode;
begin
  FRowNode := nil;

  if (FCurCol = 0) and (FDataRowCount = 0) then
    Exit;

  Root := FXMLSheet.CreateElement('pageMargins',
    'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
  for i := 1 to 6 do
    Root.Attributes[ArrNameAttributes[i]] := ArrAttributes[i];
  FXMLSheet.DocumentElement.ChildNodes.Add(Root);

  Root := FXMLSheet.DocumentElement.ChildNodes.FindNode('cols');

  WriteColumnsSec(Root);

  Res := TMemoryStream.Create;
  FXMLSheet.SaveToStream(Res);
  Res.Seek(0, soFromBeginning);
  FZipFileProvider.AddFile(Res, 'xl/worksheets/sheet1.xml');
  Res.Free;
end;

procedure TXlsxFileWriterEh.WriteColumnsSec(Root: IXMLNode);
var
  i: Integer;
  Node: IXMLNode;
begin
  for i := 0 to FColCount - 1 do
  begin
    Node := Root.AddChild('col');
    Node.Attributes['min'] := i + 1;
    Node.Attributes['max'] := i + 1;
    Node.Attributes['width'] := GetColWidth(i) / 7;
    Node.Attributes['customWidth'] := 1;
  end;
end;

procedure TXlsxFileWriterEh.WriteStylesXML;
var
  i, n: Integer;
  Res: TStream;
  Color: TColor;
  Root, Node, Node2: IXMLNode;
begin
  if FNumFmt > 164 then
    FXMLStyles.DocumentElement.ChildNodes.Insert(0, FFmtNode);

  n := Length(FFonts);
  Root := FXMLStyles.DocumentElement.ChildNodes.FindNode('fonts');
  Root.Attributes['count'] := n;
  FFonts[0].Free;
  for i := 1 to n - 1 do
  begin
    Node := Root.AddChild('font');
    if fsBold in FFonts[i].Style then
      Node.AddChild('b');
    if fsItalic in FFonts[i].Style then
      Node.AddChild('i');
    if fsUnderline in FFonts[i].Style then
      Node.AddChild('u');
    if fsStrikeOut in FFonts[i].Style then
      Node.AddChild('strike');
    Node.AddChild('sz').Attributes['val'] := FFonts[i].Size;
    Node.AddChild('color').Attributes['rgb'] := 'FF' + ColorToHex(FFonts[i].Color);
    Node.AddChild('name').Attributes['val'] := FFonts[i].Name;
    Node.AddChild('family').Attributes['val'] := 2;
    Node.AddChild('charset').Attributes['val'] := 204;
    FFonts[i].Free;
  end;
  SetLength(FFonts, 0);

  n := Length(FBackColor);
  Root := FXMLStyles.DocumentElement.ChildNodes.FindNode('fills');
  Root.Attributes['count'] := n;
  for i := 2 to n - 1 do
  begin
    Node := Root.AddChild('fill').AddChild('patternFill');
    Node.Attributes['patternType'] := 'solid';
    Node.AddChild('fgColor').Attributes['rgb'] := 'FF' + ColorToHex(FBackColor[i]);
    Node.AddChild('bgColor').Attributes['index'] := 0;
  end;
  SetLength(FBackColor, 0);

  n := Length(FBorder);
  Root := FXMLStyles.DocumentElement.ChildNodes.FindNode('borders');
  Root.Attributes['count'] := n;
  for i := 1 to n - 1 do
  begin
    Node := Root.AddChild('border');
    if FBorder[i] div 2 = 1 then
    begin
      Color := clBlack;
      Node2 := Node.AddChild('left');
      Node2.Attributes['style'] := 'thin';
      Node2.AddChild('color').Attributes['rgb'] := 'FF' + ColorToHex(Color);
      Node2 := Node.AddChild('right');
      Node2.Attributes['style'] := 'thin';
      Node2.AddChild('color').Attributes['rgb'] := 'FF' + ColorToHex(Color);
    end;
    if FBorder[i] mod 2 = 1 then
    begin
      Color := clBlack;
      Node2 := Node.AddChild('top');
      Node2.Attributes['style'] := 'thin';
      Node2.AddChild('color').Attributes['rgb'] := 'FF' + ColorToHex(Color);
      Node2 := Node.AddChild('bottom');
      Node2.Attributes['style'] := 'thin';
      Node2.AddChild('color').Attributes['rgb'] := 'FF' + ColorToHex(Color);
    end;
  end;
  SetLength(FBorder, 0);

  n := Length(FStyles);
  Root := FXMLStyles.DocumentElement.ChildNodes.FindNode('cellXfs');
  Root.Attributes['count'] := n;
  for i := 1 to n - 1 do
  begin
    Node := Root.AddChild('xf');
    Node.Attributes['numFmtId'] := FStyles[i].NumFmt;
    Node.Attributes['fontId'] := FStyles[i].NumFont;
    Node.Attributes['fillId'] := FStyles[i].NumBackColor;
    Node.Attributes['borderId'] := FStyles[i].NumBorder;
    Node.Attributes['xfId'] := 0;
    Node.Attributes['applyFont'] := 1;
    Node.Attributes['applyAlignment'] := 1;
    Node.Attributes['applyFill'] := 1;
    Node.Attributes['applyBorder'] := 1;
    Node2 := Node.AddChild('alignment');

    case FStyles[i].NumAlignment of
      0: Node2.Attributes['horizontal'] := 'left';
      1: Node2.Attributes['horizontal'] := 'right';
      2: Node2.Attributes['horizontal'] := 'center';
    end;
    if FStyles[i].Vertical then
      Node2.Attributes['textRotation'] := 90;
    if FStyles[i].Wrap then
      Node2.Attributes['wrapText'] := 1;
  end;
  SetLength(FStyles, 0);

  Res := TMemoryStream.Create;
  FXMLStyles.SaveToStream(Res);
  Res.Seek(0, soFromBeginning);
  FZipFileProvider.AddFile(Res, 'xl/styles.xml');
  Res.Free;

end;

procedure TXlsxFileWriterEh.WritetToFile(const AFileName: String);
var
  Eof: Boolean;
begin
  FFileName := AFileName;

  InitFileData;
  CreateXMLs;

  FColCount := GetColCount;
  Eof := InitTableCheckEof;
  FCurRow := 0;
  while not Eof do
  begin
    WriteRow(FCurRow, Eof);
    Inc(FCurRow);
  end;

  WriteXMLs;
  SaveDataToFile;
end;

procedure TXlsxFileWriterEh.SaveDataToFile;
begin
  FZipFileProvider.FinalizeZipFile;
  FreeAndNil(FZipFileProvider);
end;

function TXlsxFileWriterEh.SetStyle(NumFmt: Integer; Font: TFont;
  Alignment: Integer; BackColor: TColor; Border: Integer;
  Vert, Wrap: Boolean): Integer;
var
  i, n, NFont, NBackColor, NBorder: Integer;
begin
  i := 0;
  n := Length(FStyles);
  NFont := SetFont(Font);
  NBackColor := SetBackColor(BackColor);
  NBorder := SetBorder(Border);
  while (i < Length(FStyles)) and
        ( (NumFmt <> FStyles[i].NumFmt) or
          (NFont <> FStyles[i].NumFont) or
          (Alignment <> FStyles[i].NumAlignment) or
          (NBackColor <> FStyles[i].NumBackColor) or
          (NBorder <> FStyles[i].NumBorder) or
          (Vert <> FStyles[i].Vertical) or
          (Wrap <> FStyles[i].Wrap))
  do
    Inc(i);
  if i = n then
  begin
    SetLength(FStyles, n + 1);
    FStyles[n].NumFmt := NumFmt;
    FStyles[n].NumFont := NFont;
    FStyles[n].NumAlignment := Alignment;
    FStyles[n].NumBackColor := NBackColor;
    FStyles[n].NumBorder := NBorder;
    FStyles[n].Vertical := Vert;
    FStyles[n].Wrap := Wrap;
  end;
  Result := i;
end;

function TXlsxFileWriterEh.SysVarToStr(var Val: Variant): String;
var
  ASep: Char;
begin
  ASep := FormatSettings.DecimalSeparator;
  FormatSettings.DecimalSeparator := '.';
  try
    Result := VarToStr(Val);
  finally
    FormatSettings.DecimalSeparator := ASep;
  end;
end;

function TXlsxFileWriterEh.SetFont(Font: TFont): Integer;
var
  i, n: Integer;
begin
  i := 0;
  n := Length(FFonts);
  while (i < n) and
        ( (Font.Name <> FFonts[i].Name) or
          (Font.Size <> FFonts[i].Size) or
          (Font.Color <> FFonts[i].Color) or
          (Font.Style <> FFonts[i].Style)
        )
  do
    Inc(i);

  if i = n then
  begin
    SetLength(FFonts, n + 1);
    FFonts[n] := TFont.Create;
    FFonts[n].Name := Font.Name;
    FFonts[n].Size := Font.Size;
    FFonts[n].Color := Font.Color;
    FFonts[n].Style := Font.Style;
  end;
  Result := i;
end;

function TXlsxFileWriterEh.SetBackColor(Color: TColor): Integer;
var
  i, n: Integer;
begin
  i := 0;
  n := Length(FBackColor);
  while (i < Length(FBackColor)) and (Color <> FBackColor[i]) do
    Inc(i);
  if i = n then
  begin
    SetLength(FBackColor, n + 1);
    FBackColor[n] := Color;
  end;
  Result := i;
end;

function TXlsxFileWriterEh.SetBorder(Border: Integer): Integer;
var
  i, n: Integer;
begin
  i := 0;
  n := Length(FBorder);
  while (i < Length(FBorder)) and (Border <> FBorder[i]) do
    Inc(i);
  if i = n then
  begin
    SetLength(FBorder, n + 1);
    FBorder[n] := Border;
  end;
  Result := i;
end;

function TXlsxFileWriterEh.InitTableCheckEof: Boolean;
begin
  Result := True;
end;

function TXlsxFileWriterEh.InitRowCheckEof(ARow: Integer): Boolean;
begin
  Result := True;
end;

procedure TXlsxFileWriterEh.WriteRow(ARow: Integer; var Eof: Boolean);
var
  Root: IXMLNode;
  i: Integer;
begin
  Eof := InitRowCheckEof(ARow);
  if Eof then Exit;

  FCurCol := 0;

  Root := FXMLSheet.DocumentElement.ChildNodes.FindNode('sheetData');

  FRowNode := Root.AddChild('row');
  FRowNode.Attributes['r'] := IntToStr(CurRowNum);
  FRowNode.Attributes['spans'] := '1:' + IntToStr(FColCount);
  FRowNode.Attributes['x14ac:dyDescent'] := '0.25';

  for i := 0 to FColCount-1 do
    WriteValue(i, ARow);

end;

procedure TXlsxFileWriterEh.GetCellData(ACol, ARow: Integer; CellData: TXlsxCellDataEh);
begin

end;

function TXlsxFileWriterEh.GetColCount: Integer;
begin
  Result := -1;
end;

function TXlsxFileWriterEh.GetColWidth(ACol: Integer): Integer;
begin
  Result := -1;
end;

function TXlsxFileWriterEh.GetCurColNum: Integer;
begin
  Result := FCurCol+1;
end;

function TXlsxFileWriterEh.GetCurRowNum: Integer;
begin
  Result := FCurRow+1;
end;

procedure TXlsxFileWriterEh.WriteValue(ACol, ARow: Integer);
var
  t, f: Boolean;
  temp, b, NumFmt: Integer;
  Data: Variant;
  Node, Node2: IXMLNode;
  Color: TColor;
  AVarType: TVarType;
begin
  Node2 := FRowNode.AddChild('c');
  Node2.Attributes['r'] := ZEGetA1byCol(FCurCol) + IntToStr(CurRowNum);

  NumFmt := 0;
  b := 0;

  FCellData.Clear;
  GetCellData(ACol, ARow, FCellData);
  AVarType := VarType(FCellData.Value);

  if FCellData.HorzLine then
    b := b + 1;
  if FCellData.VertLine then
    b := b + 2;

{
  if dghAutoFitRowHeight in Column.Grid.OptionsEh
    then t := True
    else t := False;

  if xlsxColoredEh in FOptions
    then Color := FColCellParamsEh.Background
    else Color := TDBGridEh(DBGridEh).Color;
}
  t := True;
  Color := FCellData.Color;

  if AVarType in [varEmpty, varNull] then
    Node2.AddChild('v').Text := ''
  else
  begin
    begin
      if (AVarType = varDate) and (FCellData.Value < EncodeDate(1900, 1, 1)) then
      begin 
        Data := VarToStr(FCellData.DisplayText);
        Node2.Attributes['t'] := 'str';
      end
      else
      begin
        Data := FCellData.Value;
        f := False;
        if VarIsNumeric(FCellData.Value) and
           (FCellData.DisplayFormat <> '')
        then
          if IsNumber(FCellData.DisplayFormat) then
          begin
            Node := FFmtNode.AddChild('numFmt');
            Node.Attributes['numFmtId'] := FNumFmt;
            NumFmt := FNumFmt;
            Inc(FNumFmt);
            Node.Attributes['formatCode'] := FCellData.DisplayFormat;
          end
          else
            f := True;
        if f or not VarIsNumeric(FCellData.Value) then
          Node2.Attributes['t'] := 'str';
      end;

      temp := SetStyle(0, FCellData.Font, Integer(FCellData.Alignment), Color, b, False, t);
      if temp > 0 then
        Node2.Attributes['s'] := temp;
      Node2.AddChild('v').Text := SysVarToStr(Data);
    end;
  end;

  temp := SetStyle(NumFmt, FCellData.Font, Integer(FCellData.Alignment), Color, b, False, t);
  if temp > 0 then
    Node2.Attributes['s'] := temp;

  Inc(FCurCol);
end;

{ TCustomFileZipingProviderEh }

function ZipFileProviderClass: TCustomFileZipingProviderEhClass;
begin
  Result := FZipFileProviderClass;
end;

function RegisterZipFileProviderClass(AZipFileProviderClass: TCustomFileZipingProviderEhClass):
  TCustomFileZipingProviderEhClass;
begin
  Result := FZipFileProviderClass;
  FZipFileProviderClass := AZipFileProviderClass;
end;

class function TCustomFileZipingProviderEh.CreateInstance: TCustomFileZipingProviderEh;
begin
  Result := nil;
end;

{$IFDEF EH_LIB_16} 

{ TSystemZipFileProvider }

constructor TSystemZipFileProviderEh.Create;
begin
  inherited Create;
end;

destructor TSystemZipFileProviderEh.Destroy;
begin
  inherited Destroy;
  FreeAndNil(FZipFile);
  FreeAndNil(FStream);
end;

class function TSystemZipFileProviderEh.CreateInstance: TCustomFileZipingProviderEh;
begin
  Result := TSystemZipFileProviderEh.Create;
end;

function TSystemZipFileProviderEh.InitZipFile(FileName: String): TStream;
begin
  if FStream <> nil then
    raise Exception.Create('ZipFile is already Initialized.');
  if FZipFile <> nil then
    raise Exception.Create('ZipFile is already Initialized.');
  if FileName = '' then
    FStream := TMemoryStream.Create
  else
    FStream := TFileStream.Create(FileName, fmCreate);
  Result := FStream;
  FZipFile := TZipFile.Create;
  FZipFile.Open(Result, zmWrite);
end;

procedure TSystemZipFileProviderEh.AddFile(Data: TStream; const FilePathAndName: string);
begin
  FZipFile.Add(Data, FilePathAndName);
end;

procedure TSystemZipFileProviderEh.FinalizeZipFile;
begin
  if FZipFile <> nil then
    FZipFile.Close;
  FreeAndNil(FZipFile);
  FreeAndNil(FStream);
end;

procedure TSystemZipFileProviderEh.FinalizeZipFile(AStream: TStream);
begin
  if FZipFile <> nil then
    FZipFile.Close;
  AStream.CopyFrom(FStream, 0);
  FreeAndNil(FZipFile);
  FreeAndNil(FStream);
end;

{$ENDIF} 

procedure InitUnit;
begin
{$IFDEF EH_LIB_16} 
  RegisterZipFileProviderClass(TSystemZipFileProviderEh);
{$ENDIF} 
end;

procedure FinalizeUnit;
begin
{$IFDEF EH_LIB_16} 
  RegisterZipFileProviderClass(nil);
{$ENDIF} 
end;

{ TXlsxCellDataEh }

constructor TXlsxCellDataEh.Create;
begin
  inherited Create;
  FFont := TFont.Create;
end;

destructor TXlsxCellDataEh.Destroy;
begin
  FreeAndNil(FFont);
  inherited Destroy;
end;

procedure TXlsxCellDataEh.Clear;
begin
  FValue := Unassigned;
  Color := clNone;
  Font.Name := '';
  Font.Size := -1;
  Font.Color := -1;
  Font.Style := [];
  FHorzLine := False;
  FVertLine := False;
end;

initialization
  InitUnit;
finalization
  FinalizeUnit;

{$ENDIF} // not Lazarus

end.

