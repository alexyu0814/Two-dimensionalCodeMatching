unit XLSReadWriteII5;

{-
********************************************************************************
******* XLSReadWriteII V5.00                                             *******
*******                                                                  *******
******* Copyright(C) 1999,2013 Lars Arvidsson, Axolot Data               *******
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
{$I XLSRWII.inc}

interface

//* @anchor(a_ExcelFiles)About Excel files.
//* Excel files comes in two major models. Files up to Excel 2006 are binary
//* files and called Excel 97 as the file format was introduced in Excel 97.
//* Files after 2006 are called Excel 2007. These files are mostly XML files
//* stored in a ZIP files. You you change the extension on an Excel 2007+ file
//* to ZIP, you can open the file with WinZIP or similar. The Excel 2007 file
//* format is an official ISO/IEC standard called Office Open.
//* There is also binary variant of Excel 2007. These files has the extesnion
//* XLSB. This file format is hardly ever used and is not supported by
//* XLSReadWriteII5.

{.$define AX_SHAREWARE}

uses Classes, SysUtils,
{$ifndef BABOON}
     Dialogs, Windows,
{$endif}
{$ifdef DELPHI_XE3_OR_LATER}
     System.UITypes,
{$endif}
     xpgParseDrawing,
     Xc12Utils5, Xc12Manager5, Xc12DataWorkbook5, Xc12DefaultData5,
{$ifdef XLS_BIFF}
     BIFF5, BIFF_Utils5, BIFF_ReadII5, BIFF_Escher5, BIFF_EscherTypes5, BIFF_DrawingObj5,
     BIFF_SheetData5,
{$endif}
     XLSUtils5, XLSReadXLSX5, XLSWriteXLSX5, XLSSheetData5, XLSFormula5, XLSNames5,
     XLSMergedCells5, XLSHyperlinks5, XLSValidate5, XLSAutofilter5, XLSCondFormat5,
     XLSClassFactory5, XLSEvaluateFmla5, XLSDrawing5,
     XLSZip5;

type TXLSReadWriteII5 = class;

     TXLSClassFactoryImpl = class(TXLSClassFactory)
protected
     FOwner   : TXLSReadWriteII5;
public
     constructor Create(AOwner: TXLSReadWriteII5);

     function CreateAClass(AClassType: TXLSClassFactoryType; AClassOwner: TObject = Nil): TObject; override;
//     function GetAClass(AClassType: TXLSClassFactoryType): TObject; override;
     end;

{$ifdef DELPHI_XE2_OR_LATER}
     [ComponentPlatformsAttribute (pidWin32 or pidWin64 or pidOSX32)]
{$endif}
     TXLSReadWriteII5 = class(TXLSWorkbook)
private
     function  GetFilename: AxUCString;
     procedure SetFilename(const Value: AxUCString);
     function  GetProgressEvent: TXLSProgressEvent;
     procedure SetProgressEvent(const Value: TXLSProgressEvent);
     function  GetVersionNumber: AxUCString;
     procedure SetVerionNumber(const Value: AxUCString);
     function  GetVersion: TExcelVersion;
     procedure SetVersion(const Value: TExcelVersion);
     function  GetStrTRUE: AxUCString;
     procedure SetStrTRUE(const Value: AxUCString);
     function  GetStrFALSE: AxUCString;
     procedure SetStrFALSE(const Value: AxUCString);
     function  GetPassword: AxUCString;
     procedure SetPassword(const Value: AxUCString);
     function  GetDirectRead: boolean;
     function  GetDirectWrite: boolean;
     procedure SetDirectRead(const Value: boolean);
     procedure SetDirectWrite(const Value: boolean);
     function  GetCellReadEvent: TCellReadWriteEvent;
     function  GetCellWriteEvent: TCellReadWriteEvent;
     procedure SetCellReadEvent(const Value: TCellReadWriteEvent);
     procedure SetCellWriteEvent(const Value: TCellReadWriteEvent);
     function  GetUseAlternateZip: boolean;
     procedure SetUseAlternateZip(const Value: boolean);
     function  GetUserFunctionEvent: TUserFunctionEvent;
     procedure SetUserFunctionEvent(const Value: TUserFunctionEvent);
     function  GetManager: TXc12Manager;
     function  GetCompressStrings: boolean;
     procedure SetCompressStrings(const Value: boolean);
{$ifdef XLS_BIFF}
     function  GetBIFF: TBIFF5;
     function  GetFunctionEvent: TFunctionEvent;
     procedure SetFunctionEvent(const Value: TFunctionEvent);
{$endif}
protected
     FErrors        : TXLSErrorManager;
     FClassFactory  : TXLSClassFactoryImpl;
{$ifdef XLS_BIFF}
     FReadVBA       : boolean;
     FPasswordEvent : TPasswordEvent;
     FSkipExcel97Drw: boolean;
{$endif}

{$ifdef AX_SHAREWARE}
     FNagMsgShown: boolean;
{$endif}
{$ifdef XLS_BIFF}
     function  AddImage97(AAnchor: TCT_TwoCellAnchor; const ASheetIndex: integer; const AColOffs,ARowOffs: double): boolean;  override;
{$endif}

     procedure BeforeRead;
     procedure AfterRead;
     procedure BeforeWrite;
     procedure AddBIFFImages;
     procedure AddBIFFComments;
     procedure SaveBIFFComments;
public
     constructor Create(AOwner: TComponent); override;
     destructor Destroy; override;

     //* Deletes all data from the workbook and adds ADefaulSheetsCount
     //* worksheets.
     procedure Clear(ADefaulSheetsCount: integer = 1);

     //* Reads the Excel file give by @link(Filename).
     procedure Read;
     //* Does the same thing as @link(Filename) and @link(Read) but in one call.
     //* The @link(Filename) property is set to AFilname.
     procedure LoadFromFile(AFilename: AxUCString);
     //* Reads an Excel file from a stream. If AAutoDetect is True the method
     //* will try to detect if the stream is an Excel 2007 or an Excel 97 stream.
     //* If you explicite want to read an Excel 97 file, use @link(LoadFromStream97).
     procedure LoadFromStream(AStream: TStream; const AAutoDetect: boolean = True);
     //* Reads an Excel 97 file from a stream. If you want to read an
     //* Excel 2007+ file, use @link(LoadFromStream).
     procedure LoadFromStream97(AStream: TStream);

     //* Writes the data the file give by @link(Filename).
     procedure Write;
     //* Does the same thing as @link(Filename) and @link(Write) but in one call.
     //* The @link(Filename) property is set to AFilname.
     procedure SaveToFile(AFilename: AxUCString);
     //* Saves the data to a stream. If Version is xvExcel97, it will be an
     //* Excel97 (XLS) stream. If Version is xvExcel2007, it will be an
     //* Excel2007 (XLSX) stream.
     procedure SaveToStream(AStream: TStream);

     //* Sets the are you want to fill with cell values when using @link(DirectWrite).
     procedure SetDirectWriteArea(const ASheetIndex,ACol1,ARow1,ACol2,ARow2: integer);
{$ifdef _AXOLOT_DEBUG}
     procedure CheckIntegrity(const AList: TStrings);
{$endif}
{$ifdef XLS_BIFF}
     // The BIFF object contains data for Excel 97 files.
     property BIFF: TBIFF5 read GetBIFF;

     // Only Excel 97 macros are read.
     property ReadVBA: boolean read FReadVBA write FReadVBA;

     property SkipExcel97Drawing: boolean read FSkipExcel97Drw write FSkipExcel97Drw;
{$endif}
     //* Change StrTRUE property to change the string representation of the
     //* boolean value True. The default is 'TRUE'.
     //* See also @link(StrFALSE)
     property StrTRUE: AxUCString read GetStrTRUE write SetStrTRUE;
     //* Change StrFALSE property to change the string representation of the
     //* boolean value False. The default is 'FALSE'.
     //* See also @link(StrTRUE)
     property StrFALSE: AxUCString read GetStrFALSE write SetStrFALSE;
     //* The password for encrypted files.
     property Password: AxUCString read GetPassword write SetPassword;

     // Pictures will not work when using Delphi's ZIP.
     property UseAlternateZip : boolean read GetUseAlternateZip write SetUseAlternateZip;

     property CompressStrings: boolean read GetCompressStrings write SetCompressStrings;

     // TODO remove on final.
     property Manager: TXc12Manager read GetManager;
published
     //* The version of the TXLSReadWriteII5 component.
     property ComponentVersion: AxUCString read GetVersionNumber write SetVerionNumber;

     //* The Excel file version for the workbook.
     //* When reading files, Version is set according to the file extension.@br
     //* xls and xlm = xvExcel97
     //* xlsx and xlsm = xvExcel2007
     //* See also @link(Xc12Utils5.TExcelVersion), @link(a_ExcelFiles)
     property Version: TExcelVersion read GetVersion write SetVersion;

     //* The name of the Excel file.
     property Filename: AxUCString read GetFilename write SetFilename;
     //* When DirectRead mode is active, the cell values are not stored in
     //* memory. Instead is the OnReadCell event fired when a cell is read.
     //* You can then access the cell value in the @link(Xc12Manager5.TXLSEventCell)
     //* object, ACell parameter, in the event. @br
     //* The advantage with the DirectRead mode is that no memory is used for
     //* storing cell values. As Excel 2007+ files can have 100 of millions of
     //* cells this can be a greate memory save. If you know that you just
     //* want a few cells, or going to store them elsewhere, this is a good
     //* choice. See also @link(DirectWrite).
     property DirectRead: boolean read GetDirectRead write SetDirectRead;
     //* DirectWrite mode writes cell values directly to the disk file instead
     //* of first storing them in memory. This can save hughe ammount of memory
     //* if you are writing large files. See also @link(DirectRead).
     property DirectWrite: boolean read GetDirectWrite write SetDirectWrite;
     //* Error manager. See @link(XLSUtils5.TXLSErrorManager).
     property Errors: TXLSErrorManager read FErrors;

     property OnReadCell : TCellReadWriteEvent read GetCellReadEvent write SetCellReadEvent;
     property OnWriteCell: TCellReadWriteEvent read GetCellWriteEvent write SetCellWriteEvent;
{$ifdef XLS_BIFF}
     property OnPassword : TPasswordEvent read FPasswordEvent write FPasswordEvent;
     property OnFunction : TFunctionEvent read GetFunctionEvent write SetFunctionEvent;
{$endif}
     property OnProgress : TXLSProgressEvent read GetProgressEvent write SetProgressEvent;
     property OnUserFunction: TUserFunctionEvent read GetUserFunctionEvent write SetUserFunctionEvent;
     end;

implementation

{ TXLSReadWriteII5 }

{$ifdef XLS_BIFF}
function TXLSReadWriteII5.AddImage97(AAnchor: TCT_TwoCellAnchor; const ASheetIndex: integer; const AColOffs,ARowOffs: double): boolean;
var
  MSO: TMSOPicture;
  DrwPict: TDrwPicture;
begin
  Result := FBIFF <> Nil;

  if Result then begin
    MSO := FBIFF.MSOPictures.Add;
    MSO.Filename := AAnchor.Objects.Pic.NvPicPr.CNvPr.Descr;
    DrwPict := FBIFF.Sheets[ASheetIndex].DrawingObjects.Pictures.Add;
    DrwPict.PictureName := MSO.Filename;
    DrwPict.Col1       := AAnchor.From.Col;
    DrwPict.Col1Offset := AAnchor.From.ColOff;
    DrwPict.Row1       := AAnchor.From.Row;
    DrwPict.Row1Offset := AAnchor.From.RowOff;
    DrwPict.Col2       := AAnchor.To_.Col;
    DrwPict.Col2Offset := AColOffs;
    DrwPict.Row2       := AAnchor.To_.Row;
    DrwPict.Row2Offset := ARowOffs;
  end;
end;
{$endif}

procedure TXLSReadWriteII5.AddBIFFComments;
var
  i,j: integer;
  Sht: TSheet;
  N: TDrwNote;
begin
  for i := 0 to FBIFF.Sheets.Count - 1 do begin
    if i < Count - 1 then begin
      Sht := FBIFF[i];
      for j := 0 to Sht.DrawingObjects.Notes.Count - 1 do begin
        N := Sht.DrawingObjects.Notes[j];
        Items[i].Comments.AsPlainText[N.CellCol,N.CellRow] := N.Text;
      end;
    end;
  end;
end;

procedure TXLSReadWriteII5.AddBIFFImages;
var
  i,j: integer;
  Pict: TDrwPicture;
  MSOPict: TMSOPicture;
  Stream: TMemoryStream;
begin
  for i := 0 to Count - 1 do begin
    for j := 0 to Items[i].Drawing.BIFFDrawing.Pictures.Count - 1 do begin
      Pict := Items[i].Drawing.BIFFDrawing.Pictures[j];
      if Pict.PictureId > 0 then begin
        MSOPict := FBIFF.MSOPictures[Pict.PictureId - 1];
        Stream := TMemoryStream.Create;
        try
          MSOPict.SaveToStream(Stream);
          Stream.Seek(0,soFromBeginning);
          case MSOPict.PictType of
  //          msoblipERROR: ;
  //          msoblipUNKNOWN: ;
            msoblipEMF : Items[i].Drawing.InsertImage97(Format('Pict%d_%d.emf',[i + 1,j + 1]),Stream,Pict);
            msoblipWMF : Items[i].Drawing.InsertImage97(Format('Pict%d_%d.wmf',[i + 1,j + 1]),Stream,Pict);
  //          msoblipPICT: ;
            msoblipJPEG: Items[i].Drawing.InsertImage97(Format('Pict%d_%d.jpg',[i + 1,j + 1]),Stream,Pict);
            msoblipPNG : Items[i].Drawing.InsertImage97(Format('Pict%d_%d.png',[i + 1,j + 1]),Stream,Pict);
            msoblipDIB : Items[i].Drawing.InsertImage97(Format('Pict%d_%d.bmp',[i + 1,j + 1]),Stream,Pict);
          end;
        finally
          Stream.Free;
        end;
      end;
    end;
  end;
end;

procedure TXLSReadWriteII5.AfterRead;
var
  i: integer;
begin
  inherited AfterRead;
  FManager.AfterRead;
{$ifdef AX_SHAREWARE}
  for i := 0 to Count - 1 do
    Items[i].AsString[0,0] := 'XLSReadWriteII Copyright(c) 2014 Axolot Data';
{$endif}

{$ifdef XLS_BIFF}
  if FBIFF <> Nil then begin
    for i := 0 to Count - 1 do
      Items[i].Drawing.BIFFDrawing := FBIFF.Sheet[i].DrawingObjects;

    AddBIFFImages;
    AddBIFFComments;
  end;
{$endif}
end;

procedure TXLSReadWriteII5.BeforeRead;
{$ifdef AX_SHAREWARE}
var
  i: integer;
{$endif}
begin
  FManager.BeforeRead;
{$ifdef AX_SHAREWARE}
  for i := 0 to Count - 1 do
    Items[i].AsString[0,0] := 'XLSReadWriteII Copyright(c) 2014 Axolot Data';
{$endif}
end;

procedure TXLSReadWriteII5.BeforeWrite;
{$ifdef AX_SHAREWARE}
var
  i: integer;
{$endif}
begin
  inherited BeforeWrite;

{$ifdef AX_SHAREWARE}
  for i := 0 to Count - 1 do
    Items[i].AsString[0,0] := 'XLSReadWriteII Copyright(c) 2014 Axolot Data';
{$endif}

  SaveBIFFComments;
end;

{$ifdef _AXOLOT_DEBUG}
procedure TXLSReadWriteII5.CheckIntegrity(const AList: TStrings);
begin
  inherited CheckIntegrity(AList);
end;
{$endif}

procedure TXLSReadWriteII5.Clear(ADefaulSheetsCount: integer = 1);
var
  i: integer;
begin
  FManager.Clear;
  inherited Clear;
{$ifdef XLS_BIFF}
  if FBIFF <> Nil then begin
    FBIFF.Free;
    FBIFF := Nil;
  end;
{$endif}
  ADefaulSheetsCount := Fork(ADefaulSheetsCount,0,XLS_MAXSHEETS);

  for i := 0 to ADefaulSheetsCount - 1 do
    Add;
end;

constructor TXLSReadWriteII5.Create(AOwner: TComponent);
{$ifndef BABOON}
{$ifdef AX_SHAREWARE}
var
  A,B,C,D: boolean;
{$endif}
{$endif}
begin
  inherited Create(AOwner);

{$ifndef BABOON}
{$ifdef AX_SHAREWARE}
{$ifdef AX_SHAREWARE_DLL}
  FNagMsgShown := True;
{$endif}
  if not FNagMsgShown then begin
    A := (FindWindow('TApplication', nil) = 0);
//    B := (FindWindow('TAlignPalette', nil) = 0);
//    C := (FindWindow('TPropertyInspector', nil) = 0);
    B := False;
    C := False;
    D := (FindWindow('TAppBuilder', nil) = 0);
    if A or B or C or D then begin
      MessageDlg('This application was built with a demo version of' + #13 +
                  'the XLSReadWriteII components.' + #13 + #13 +
                  'Distributing an application based upon this version' + #13 +
                  'of the components are against the licensing agreement.' + #13 + #13 +
                  'Please see http://www.axolot.com for more information' + #13 +
                  'on purchasing.',mtInformation,[mbOk],0);
      FNagMsgShown := True;
    end;
  end;
{$endif}
{$endif}

  FErrors := TXLSErrorManager.Create;

  FClassFactory := TXLSClassFactoryImpl.Create(Self);

  FManager := TXc12Manager.Create(FClassFactory,FErrors);
  FFormulas := TXLSFormulaHandler.Create(FManager);

  FManager.CreateMembers;

  CreateMembers;

  Clear;
end;

destructor TXLSReadWriteII5.Destroy;
begin
  FManager.Free;
  FFormulas.Free;
{$ifdef XLS_BIFF}
  if FBIFF <> Nil then
    FBIFF.Free;
{$endif}
  FErrors.Free;

  FClassFactory.Free;
  inherited;
end;

{$ifdef XLS_BIFF}
function TXLSReadWriteII5.GetBIFF: TBIFF5;
begin
  Result := FBIFF;
end;
{$endif}

function TXLSReadWriteII5.GetCellReadEvent: TCellReadWriteEvent;
begin
  Result := FManager.OnReadCell;
end;

function TXLSReadWriteII5.GetCellWriteEvent: TCellReadWriteEvent;
begin
  Result := FManager.OnWriteCell;
end;

function TXLSReadWriteII5.GetCompressStrings: boolean;
begin
  Result := FManager.SST.CompressStrings;
end;

function TXLSReadWriteII5.GetDirectRead: boolean;
begin
  Result := FManager.DirectRead;
end;

function TXLSReadWriteII5.GetDirectWrite: boolean;
begin
  Result := FManager.DirectWrite;
end;

function TXLSReadWriteII5.GetFilename: AxUCString;
begin
  Result := FManager.Filename;
end;

{$ifdef XLS_BIFF}
function TXLSReadWriteII5.GetFunctionEvent: TFunctionEvent;
begin
  Result := Nil;
end;
{$endif}

function TXLSReadWriteII5.GetManager: TXc12Manager;
begin
  Result := FManager;
end;

function TXLSReadWriteII5.GetPassword: AxUCString;
begin
  Result := FManager.Password;
end;

function TXLSReadWriteII5.GetProgressEvent: TXLSProgressEvent;
begin
  Result := FManager.OnProgress;
end;

function TXLSReadWriteII5.GetStrFALSE: AxUCString;
begin
  Result := G_StrFALSE;
end;

function TXLSReadWriteII5.GetStrTRUE: AxUCString;
begin
  Result := G_StrTRUE;
end;

function TXLSReadWriteII5.GetUseAlternateZip: boolean;
begin
  Result := FManager.UseAlternateZip;
end;

function TXLSReadWriteII5.GetUserFunctionEvent: TUserFunctionEvent;
begin
  Result := FFormulas.OnUserFunction;
end;

function TXLSReadWriteII5.GetVersion: TExcelVersion;
begin
{$ifdef XLS_BIFF}
  if FBIFF <> Nil then
    Result := xvExcel97
  else
{$endif}
    Result := xvExcel2007;
end;

function TXLSReadWriteII5.GetVersionNumber: AxUCString;
begin
  Result := CurrentVersionNumber;
{$ifdef AX_SHAREWARE}
  Result := Result + 'a';
{$endif}
end;

procedure TXLSReadWriteII5.LoadFromFile(AFilename: AxUCString);
var
  Ext: AxUCString;
  Stream: TFileStream;
//  Strm: TStream;
begin
  SetFilename(AFilename);

  Ext := AnsiLowercase(ExtractFileExt(FManager.Filename));
  if (Ext = '.xls') or (Ext = '.xlm') then begin
    FManager.Version := xvExcel97;

    Clear(0);

    BeforeRead;
{$ifdef XLS_BIFF}
    FBIFF := TBIFF5.Create(FManager);
    FBIFF.Filename := AFilename;
    FBIFF.OnPassword := FPasswordEvent;
    FBIFF.ReadMacros := FReadVBA;
    FBIFF.SkipDrawing := FSkipExcel97Drw;
    FManager._ExtNames97 := FBIFF.FormulaHandler.ExternalNames;
    FManager.Names97 := FBIFF.FormulaHandler.InternalNames;
    FBIFF.Read;
{$endif}
    AfterRead;
  end
  else begin
    FManager.Version := xvExcel2007;
    Stream := TFileStream.Create(FManager.Filename,fmOpenRead or fmShareDenyNone);
    try
      LoadFromStream(Stream);
    finally
      Stream.Free;
    end;

//    if FReadVBA then begin
//      Strm := FManager.FileData.StreamByName('vbaProject.bin');
//      if Strm <> Nil then begin
//
//      end;
//    end;
  end;
end;

procedure TXLSReadWriteII5.LoadFromStream(AStream: TStream; const AAutoDetect: boolean = True);
var
  XLSX: TXLSReadXLSX;
begin
  if AAutoDetect and not StreamIsZIP(AStream) then
    LoadFromStream97(AStream)
  else begin
    FManager.Version := xvExcel2007;

    XLSX := TXLSReadXLSX.Create(FManager,FFormulas);
    try
      Clear(0);

      BeforeRead;

      XLSX.LoadFromStream(AStream);

      AfterRead;
    finally
      XLSX.Free;
    end;
  end;
end;

procedure TXLSReadWriteII5.LoadFromStream97(AStream: TStream);
begin
  FManager.Version := xvExcel97;

  Clear(0);

  BeforeRead;
{$ifdef XLS_BIFF}
  FBIFF := TBIFF5.Create(FManager);
  FBIFF.SkipDrawing := FSkipExcel97Drw;
  FBIFF.OnPassword := FPasswordEvent;
  FManager._ExtNames97 := FBIFF.FormulaHandler.ExternalNames;
  FManager.Names97 := FBIFF.FormulaHandler.InternalNames;
  FBIFF.LoadFromStream(AStream);
{$endif}
  AfterRead;
end;

procedure TXLSReadWriteII5.Read;
begin
  LoadFromFile(FManager.Filename);
end;

procedure TXLSReadWriteII5.SaveBIFFComments;
var
  i,j: integer;
  N : TDrwNote;
begin
{$ifdef XLS_BIFF}
  if FBIFF <> Nil then begin
    for i := 0 to Count - 1  do begin
      FBIFF[i].DrawingObjects.Notes.Clear;
      for j := 0 to Items[i].Comments.Count - 1 do begin
        N := FBIFF[i].DrawingObjects.Notes.Add;
        N.CellCol := Items[i].Comments[j].Col;
        N.CellRow := Items[i].Comments[j].Row;
        N.Text := Items[i].Comments[j].PlainText;
        N.Author := Items[i].Comments[j].Author;
      end;
    end;
  end;
{$endif}
end;

procedure TXLSReadWriteII5.SaveToFile(AFilename: AxUCString);
var
  Stream: TFileStream;
begin
  SetFilename(AFilename);
{$ifdef XLS_BIFF}
  if FBIFF <> Nil then begin
    BeforeWrite;
{$ifdef _AXOLOT_DEBUG}
    Filename := ChangeFileExt(Filename,'.xls');
{$endif}
    FBIFF.Filename := Filename;
    FBIFF.Write;
  end
  else begin
{$endif}
    SetFilename(Filename);

    Stream := TFileStream.Create(FManager.Filename,fmCreate);
    try
      SaveToStream(Stream);
    finally
      Stream.Free;
    end;
{$ifdef XLS_BIFF}
  end;
{$endif}
end;

procedure TXLSReadWriteII5.SaveToStream(AStream: TStream);
var
  XLSX: TXLSWriteXLSX;
begin
{$ifdef XLS_BIFF}
  if FBIFF <> Nil then
    FBIFF.WriteToStream(AStream)
  else begin
{$endif}
    XLSX := TXLSWriteXLSX.Create(FManager,FFormulas);
    try
      BeforeWrite;

      CalcDimensions;

      FManager.BeforeWrite;

      XLSX.SaveToStream(AStream);
    finally
      XLSX.Free;
    end;
{$ifdef XLS_BIFF}
  end;
{$endif}
end;

procedure TXLSReadWriteII5.SetCellReadEvent(const Value: TCellReadWriteEvent);
begin
  FManager.OnReadCell := Value;
end;

procedure TXLSReadWriteII5.SetCellWriteEvent(const Value: TCellReadWriteEvent);
begin
  FManager.OnWriteCell := Value;
end;

procedure TXLSReadWriteII5.SetCompressStrings(const Value: boolean);
begin
  FManager.SST.CompressStrings := Value;
end;

procedure TXLSReadWriteII5.SetDirectRead(const Value: boolean);
begin
  FManager.DirectRead := Value;
end;

procedure TXLSReadWriteII5.SetDirectWrite(const Value: boolean);
begin
  FManager.DirectWrite := Value;
end;

procedure TXLSReadWriteII5.SetDirectWriteArea(const ASheetIndex, ACol1, ARow1, ACol2, ARow2: integer);
begin
  FManager.EventCell.SetTargetArea(ASheetIndex, ACol1, ARow1, ACol2, ARow2);
end;

procedure TXLSReadWriteII5.SetFilename(const Value: AxUCString);
begin
  FManager.Filename := Value;
end;

{$ifdef XLS_BIFF}
procedure TXLSReadWriteII5.SetFunctionEvent(const Value: TFunctionEvent);
begin

end;
{$endif}

procedure TXLSReadWriteII5.SetPassword(const Value: AxUCString);
begin
  FManager.Password := Value;
end;

procedure TXLSReadWriteII5.SetProgressEvent(const Value: TXLSProgressEvent);
begin
  FManager.OnProgress := Value;
end;

procedure TXLSReadWriteII5.SetStrFALSE(const Value: AxUCString);
begin
  G_StrFALSE := Value;
end;

procedure TXLSReadWriteII5.SetStrTRUE(const Value: AxUCString);
begin
  G_StrTRUE := Value;
end;

procedure TXLSReadWriteII5.SetUseAlternateZip(const Value: boolean);
begin
  FManager.UseAlternateZip := Value;
end;

procedure TXLSReadWriteII5.SetUserFunctionEvent(const Value: TUserFunctionEvent);
begin
  FFormulas.OnUserFunction := Value;
end;

procedure TXLSReadWriteII5.SetVerionNumber(const Value: AxUCString);
begin

end;

procedure TXLSReadWriteII5.SetVersion(const Value: TExcelVersion);
var
  ZIP: TXLSZipArchive;
  Stream: TPointerMemoryStream;
  Stream97: TMemoryStream;
begin
  if Value <> GetVersion then begin

    if Value = xvExcel97 then begin
      if (Count > 1) or ((Count = 1) and not Items[0].IsEmpty) then
        FErrors.Warning('',XLSWARN_WORKBOOK_NOT_EMPTY);

      Clear;

      FManager.Version := xvExcel97;

      Stream := TPointerMemoryStream.Create;
      try
        Stream.SetStreamData(@XLS_DEFAULT_FILE_97,Length(XLS_DEFAULT_FILE_97));
        ZIP := TXLSZipArchive.Create;
        try
          ZIP.OpenRead(Stream);
          Stream97 := TMemoryStream.Create;
          try
            ZIP[0]._SaveToStream(Stream97);
            Stream97.Seek(0,soFromBeginning);
            LoadFromStream97(Stream97);

            FManager.StyleSheet.XFEditor.LockAll;
          finally
            Stream97.Free;
          end;
        finally
          ZIP.Free;
        end;
      finally
        Stream.Free;
      end;
    end
    else begin
{$ifdef XLS_BIFF}
      if FBIFF <> Nil then begin
        FBIFF.Free;
        FBIFF := Nil;
      end;
{$endif}
      FManager.Version := xvExcel2007;
    end;
  end;
end;

procedure TXLSReadWriteII5.Write;
begin
  SaveToFile(FManager.Filename);
end;

{ TXLSClassFactoryImpl }

constructor TXLSClassFactoryImpl.Create(AOwner: TXLSReadWriteII5);
begin
  FOwner := AOwner;
end;

function TXLSClassFactoryImpl.CreateAClass(AClassType: TXLSClassFactoryType; AClassOwner: TObject): TObject;
begin
  case AClassType of
    xcftNames                : Result := TXLSNames.Create(Self,FOwner.FManager,FOwner.FFormulas);
    xcftNamesMember          : Result := TXLSName.Create(FOwner.Names);
    xcftMergedCells          : Result := TXLSMergedCells.Create(Self);
    xcftMergedCellsMember    : Result := TXLSMergedCell.Create;
    xcftHyperlinks           : Result := TXLSHyperlinks.Create(Self,FOwner.Manager);
    xcftHyperlinksMember     : begin
      Result := TXLSHyperlink.Create;
      TXLSHyperlink(Result).Owner := TXLSHyperlinks(AClassOwner);
    end;
    xcftDataValidations      : Result := TXLSDataValidations.Create(Self);
    xcftDataValidationsMember: Result := TXLSDataValidation.Create;
    xcftConditionalFormat    : Result := TXLSConditionalFormat.Create(FOwner.FManager,FOwner.FFormulas);
    xcftConditionalFormats   : Result := TXLSConditionalFormats.Create(Self);
    xcftAutofilter           : Result := TXLSAutofilter.Create(Self,AClassOwner,FOwner.Names);
    xcftAutofilterColumn     : Result := TXLSAutofilterColumn.Create;
    xcftDrawing              : Result := TXPGDocXLSXDrawing.Create(FOwner.FManager.GrManager);
    else                       Result := Nil;
  end;
end;

//function TXLSClassFactoryImpl.GetAClass(AClassType: TXLSClassFactoryType): TObject;
//begin
//  case AClassType of
//    xcftDrawingManager       : Result := FOwner.FManager.Drawings;
//    else                       Result := Nil;
//  end;
//end;

end.
