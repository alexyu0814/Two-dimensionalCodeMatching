{*******************************************************}
{                                                       }
{                        EhLib v8.1                     }
{                                                       }
{                    TPlannerDataModule                 }
{                      Build 8.1.001                    }
{                                                       }
{   Copyright (c) 2015-2015 by Dmitry V. Bolshakov      }
{                                                       }
{*******************************************************}

unit PlannerToolCtrlsEh;

interface

uses
  SysUtils, Classes, ImgList, Controls, Dialogs;

type
  TPlannerDataMod = class(TDataModule)
    PlannerImList: TImageList;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  PlannerDataMod: TPlannerDataMod;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

initialization
  PlannerDataMod := TPlannerDataMod.CreateNew(nil, -1);
  InitInheritedComponent(PlannerDataMod, TDataModule);
//  PlannerDataMod := TPlannerDataMod.Create(nil);
finalization
//  ShowMessage('FinalizeUnit PlannerToolCtrlsEh');
  FreeAndNil(PlannerDataMod);
//  ShowMessage('FinalizeUnit PlannerToolCtrlsEh end');
end.
