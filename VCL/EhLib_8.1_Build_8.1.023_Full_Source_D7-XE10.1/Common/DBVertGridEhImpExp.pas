{*******************************************************}
{                                                       }
{                       EhLib v8.1                      }
{             DBGridEh import/export routines           }
{                      Build 8.1.18                     }
{                                                       }
{   Copyright (c) 2015-2015 by Dmitry V. Bolshakov      }
{                                                       }
{*******************************************************}

unit DBVertGridEhImpExp;

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
  XlsxFormatEh, DBVertGridsEh,
  Db, Clipbrd, ComObj, StdCtrls;

type


  TDBVertGridExportToXlsxEh = class(TXlsxFileWriterEh)
  private
    FDBVertGrid: TCustomDBVertGridEh;
  protected
    function GetColWidth(ACol: Integer): Integer; override;
    function InitTableCheckEof: Boolean; override;
    function GetColCount: Integer; override;
    function InitRowCheckEof(ARow: Integer): Boolean; override;

    procedure GetCellData(ACol, ARow: Integer; CellData: TXlsxCellDataEh); override;
  public
    property DBVertGrid: TCustomDBVertGridEh read FDBVertGrid write FDBVertGrid;
  end;

  procedure ExportDBVertGridEhToXlsx(DBVertGridEh: TCustomDBVertGridEh;
    const FileName: String);

implementation

type TDBVertGridEhCrack = class(TCustomDBVertGridEh);

procedure ExportDBVertGridEhToXlsx(DBVertGridEh: TCustomDBVertGridEh;
  const FileName: String);
var
  Xlsx: TDBVertGridExportToXlsxEh;
begin
  Xlsx := TDBVertGridExportToXlsxEh.Create;
  try
    Xlsx.DBVertGrid := DBVertGridEh;
//    Xlsx.FOptions := Options;
    Xlsx.WritetToFile(FileName);
  finally
    Xlsx.Free;
  end;
end;

{ TDBVertGridExportToXlsxEh }

function TDBVertGridExportToXlsxEh.GetColCount: Integer;
begin
  Result := 2;
end;

function TDBVertGridExportToXlsxEh.GetColWidth(ACol: Integer): Integer;
begin
  Result := TDBVertGridEhCrack(DBVertGrid).ColWidths[ACol];
end;

function TDBVertGridExportToXlsxEh.InitTableCheckEof: Boolean;
begin
  if (DBVertGrid.DataSource = nil) or
     (DBVertGrid.DataSource.DataSet = nil) or
     not DBVertGrid.DataSource.DataSet.Active or
     DBVertGrid.DataSource.DataSet.IsEmpty
  then
    Result := True
  else
    Result := False;
end;

function TDBVertGridExportToXlsxEh.InitRowCheckEof(ARow: Integer): Boolean;
begin
  Result := (ARow = DBVertGrid.Rows.Count);
end;

procedure TDBVertGridExportToXlsxEh.GetCellData(ACol, ARow: Integer;
  CellData: TXlsxCellDataEh);
begin
  if ACol = 0 then
  begin
    CellData.Value := DBVertGrid.Rows[ARow].RowLabel.Caption;
    CellData.Color := DBVertGrid.Rows[ARow].RowLabel.Color;
    CellData.Font.Name := DBVertGrid.Rows[ARow].RowLabel.Font.Name;
    CellData.Font.Size := DBVertGrid.Rows[ARow].RowLabel.Font.Size;
    CellData.Font.Color := DBVertGrid.Rows[ARow].RowLabel.Font.Color;
    CellData.Font.Style := DBVertGrid.Rows[ARow].RowLabel.Font.Style;
    CellData.HorzLine := True;
    CellData.VertLine := True;
  end
  else
  begin
    CellData.Value := DBVertGrid.Rows[ARow].EditValue;
    CellData.Color := DBVertGrid.Rows[ARow].Color;
    CellData.Font.Name := DBVertGrid.Rows[ARow].Font.Name;
    CellData.Font.Size := DBVertGrid.Rows[ARow].Font.Size;
    CellData.Font.Color := DBVertGrid.Rows[ARow].Font.Color;
    CellData.Font.Style := DBVertGrid.Rows[ARow].Font.Style;
    CellData.HorzLine := True;
    CellData.VertLine := True;
  end;
end;

end.
