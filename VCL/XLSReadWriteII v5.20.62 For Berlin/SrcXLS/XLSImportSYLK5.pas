unit XLSImportSYLK5;

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
** expressed or implied, including, without limitation, the warranties of     **
** merchantability and of fitness for any purpose. The author assumes no      **
** liability for damages, direct or consequential, which may result from the  **
** use of XLSReadWriteII.                                                     **
********************************************************************************
}

{$B-}
{$H+}
{$R-}
{$C-}
{$I AxCompilers.inc}
{$I XLSRWII.inc}

interface

uses Classes, SysUtils,
     Xc12Utils5, Xc12DataWorksheet5, Xc12Manager5,
     XLSUtils5, XLSFormulaTypes5, XLSEncodeFmla5;

type TXLSImportSYLK = class(TObject)
protected
     FManager: TXc12Manager;
     FSheet  : TXc12DataWorksheet;

     FCol    : integer;
     FRow    : integer;

     FStream : TStream;
     FCurrCol: integer;
     FCurrRow: integer;

     function  StrToVal(const S: AxUCString; const AMin,AMax: integer; out AVal: integer): boolean;
     function  FirstChar(const S: AxUCString): AxUCChar; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  GetNextStr(var S: AxUCString): boolean;
     function  DoRef(var S: AxUCString): boolean;
     function  DoCell(const AText: AxUCString): boolean;
public
     constructor Create(AManager: TXc12Manager);

     procedure Read(AStream: TStream; const ASheet,ACol,ARow: integer);
     end;

implementation

{ TXLSImportSYLK }

constructor TXLSImportSYLK.Create(AManager: TXc12Manager);
begin
  FManager := AManager;
end;

function TXLSImportSYLK.DoCell(const AText: AxUCString): boolean;
var
  S: AxUCString;
  sVal: AxUCString;
  Str: AxUCString;
  Err: TXc12CellError;
  Bool: boolean;
  Num: double;
  CT: TXLSCellType;
//  Ptgs: PXLSPtgs;
//  PtgsSz: integer;
begin
  S := AText;
  SplitAtChar(';',S);

  Result := DoRef(S);
  if not Result then
    Exit;

  CT := xctNone;
  Err := errUnknown;
  Bool := False;

  sVal := SplitAtChar(';',S);
  if Copy(sVal,1,1) = 'K' then begin
    sVal := Copy(sVal,2,MAXINT);
    Err := ErrorTextToCellError(sVAL);
    if Err <> errUnknown then
      CT := xctError
    else begin
      if Copy(sVal,1,1) = '"' then begin
        StripQuotes(sVal);
        Str := sVal;
        CT := xctString;
      end
      else if sVal = 'TRUE' then begin
        CT := xctBoolean;
        Bool := True;
      end
      else if sVal = 'FALSE' then begin
        CT := xctBoolean;
        Bool := False;
      end
      else begin
        if TryStrToFloat(sVal,Num) then
          CT := xctFloat;
      end;
    end;
  end;

  Result := CT <> xctNone;
  if not Result then
    Exit;

// Formulas don't work. R1C1 mode must be implemented in Tokenizer/Encoder.
//  if FirstChar(S) = 'E' then begin
//    S := Copy(S,2,MAXINT);
//
//    if XLSEncodeFormula(FManager,S,Ptgs,PtgsSz) then begin
//      FSheet.Cells.FormulaHelper.Clear;
//
//      FSheet.Cells.FormulaHelper.Style := XLS_STYLE_DEFAULT_XF;
//
//      FSheet.Cells.FormulaHelper.Col := FCurrCol;
//      FSheet.Cells.FormulaHelper.Row := FCurrRow;
//
//      FSheet.Cells.FormulaHelper.Ptgs := Ptgs;
//      FSheet.Cells.FormulaHelper.PtgsSize := PtgsSz;
//
//      case CT of
//        xctBoolean: FSheet.Cells.FormulaHelper.AsBoolean := Bool;
//        xctError  : FSheet.Cells.FormulaHelper.AsError := Err;
//        xctString : FSheet.Cells.FormulaHelper.AsString := Str;
//        xctFloat  : FSheet.Cells.FormulaHelper.AsFloat := Num;
//      end;
//
//      FSheet.Cells.StoreFormula;
//    end;
//  end
//  else begin F;P0;FG0C;SDLRBSM29;X41
    case CT of
      xctBoolean: FSheet.Cells.UpdateBoolean(FCurrCol,FCurrRow,Bool);
      xctError  : FSheet.Cells.UpdateError(FCurrCol,FCurrRow,Err);
      xctString : FSheet.Cells.UpdateString(FCurrCol,FCurrRow,Str);
      xctFloat  : FSheet.Cells.UpdateFloat(FCurrCol,FCurrRow,Num);
    end;
//  end;
end;

function TXLSImportSYLK.DoRef(var S: AxUCString): boolean;
var
  sRef: AxUCString;
begin
  Result := False;
  if FirstChar(S) = 'X' then begin
    sRef := SplitAtChar(';',S);
    Result := StrToVal(Copy(sRef,2,MAXINT),0,XLS_MAXCOL,FCurrCol);
    if Result  then begin
      if FirstChar(S) = 'Y' then begin
        sRef := SplitAtChar(';',S);
        Result := StrToVal(Copy(sRef,2,MAXINT),0,XLS_MAXROW,FCurrRow);
      end;
    end;
  end
  else if FirstChar(S) = 'Y' then begin
    sRef := SplitAtChar(';',S);
    Result := StrToVal(Copy(sRef,2,MAXINT),0,XLS_MAXROW,FCurrRow);
    if Result then begin
      if FirstChar(S) = 'X' then begin
        sRef := SplitAtChar(';',S);
        Result := StrToVal(Copy(sRef,2,MAXINT),0,XLS_MAXCOL,FCurrCol);
      end;
    end;
  end;
end;

function TXLSImportSYLK.FirstChar(const S: AxUCString): AxUCChar;
begin
  if Length(S) > 0 then
    Result := S[1]
  else
    Result := #0;
end;

function TXLSImportSYLK.GetNextStr(var S: AxUCString): boolean;
var
  n: integer;
  b: byte;
begin
  n := 0;
  while FStream.Read(b,SizeOf(byte)) = SizeOf(byte) do begin
    Inc(n);
    if n > Length(S) then
      SetLength(S,Length(S) + 256);
    if b = 13 then begin
      FStream.Read(b,SizeOf(byte)); // 10
      SetLength(S,n - 1);
      S := Trim(S);
      Result := Length(S) > 0;
      Exit;
    end;
    S[n] := AxUCChar(b);
  end;
  SetLength(S,n);
  Result := n > 0;
end;

procedure TXLSImportSYLK.Read(AStream: TStream; const ASheet,ACol,ARow: integer);
var
  S: AxUCString;
  w: word;
begin
  FStream := AStream;
  FSheet := FManager.Worksheets[ASheet];
  FCol := ACol;
  FRow := ARow;

  FStream.Read(w,SizeOf(word));
  //       ID
  if w <> $4449 then
    Exit;

  FStream.Seek(0,soFromBeginning);

  while GetNextStr(S) do begin
    case S[1] of
      'C': DoCell(S);
    end;
  end;
end;

function TXLSImportSYLK.StrToVal(const S: AxUCString; const AMin, AMax: integer; out AVal: integer): boolean;
begin
  Result := TryStrToInt(S,AVal);
  if Result then
    Result := (AVal >= AMin) and (AVal <= AMax);
end;

end.
