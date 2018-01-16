unit XLSAutofilter5;

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
{$I XLSRWII.inc}

interface

uses Classes, SysUtils, Contnrs,
     Xc12Utils5, Xc12DataAutofilter5, Xc12DataWorkbook5, Xc12DataWorksheet5, Xc12Manager5,
     XLSUtils5, XLSClassFactory5, XLSNames5;

type TXLSAutofilterColumn = class(TXc12FilterColumn)
protected
public
     end;

type TXLSAutofilter = class(TXc12AutoFilter)
protected
     FNames    : TXLSNames;
     FXc12Sheet: TXc12DataWorksheet;

     function GetCol1: integer;
     function GetCol2: integer;
     function GetRow1: integer;
     function GetRow2: integer;
public
     constructor Create(AClassFactory: TXLSClassFactory; AClassOwner: TObject; ANames: TXLSNames);

     procedure Clear;

     function IsFiltered: boolean;

     procedure Add(const AC1,AR1,AC2,AR2: integer);

     procedure Delete(ACol1,ARow1,ACol2,ARow2: integer);
     procedure Move(Col1, Row1, Col2, Row2, DeltaCol, DeltaRow: integer);

     property Col1: integer read GetCol1;
     property Row1: integer read GetRow1;
     property Col2: integer read GetCol2;
     property Row2: integer read GetRow2;
     end;

implementation

{ TXLSAutofilter }

procedure TXLSAutofilter.Add(const AC1,AR1,AC2,AR2: integer);
var
  N: TXLSName;
begin
  Clear;
  FRef := SetCellArea(AC1,AR1,AC2,AR2);

  N := TXLSName(FNames.FindBuiltIn(bnFilterDatabase,FXc12Sheet.Index));
  if N = Nil then
    N := TXLSName(FNames.Add(bnFilterDatabase,FXc12Sheet.Index));
  N.Definition := '''' + FXc12Sheet.Name + '''!' + AreaToRefStr(AC1,AR1,AC2,AR2,True,True,True,True);
  N.Hidden := True;
end;

procedure TXLSAutofilter.Clear;
begin

end;

constructor TXLSAutofilter.Create(AClassFactory: TXLSClassFactory; AClassOwner: TObject; ANames: TXLSNames);
begin
  inherited Create(AClassFactory);

  FXc12Sheet := TXc12DataWorksheet(AClassOwner);
  FNames := ANames;
end;

procedure TXLSAutofilter.Delete(ACol1, ARow1, ACol2, ARow2: integer);
begin

end;

function TXLSAutofilter.GetCol1: integer;
begin
  Result := 0;
end;

function TXLSAutofilter.GetCol2: integer;
begin
  Result := 0;
end;

function TXLSAutofilter.GetRow1: integer;
begin
  Result := 0;
end;

function TXLSAutofilter.GetRow2: integer;
begin
  Result := 0;
end;

function TXLSAutofilter.IsFiltered: boolean;
begin
  Result := False;
end;

procedure TXLSAutofilter.Move(Col1, Row1, Col2, Row2, DeltaCol, DeltaRow: integer);
begin

end;

end.
