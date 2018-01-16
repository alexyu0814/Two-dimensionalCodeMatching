{*******************************************************}
{                                                       }
{                        EhLib v8.1                     }
{                                                       }
{                  TCompoManEh component                }
{                      Build 8.1.001                    }
{                                                       }
{    Copyright (c) 2014-2015 by Dmitry V. Bolshakov     }
{                                                       }
{*******************************************************}

{$I EhLib.Inc}

unit CompoMansEh;

interface

uses
  Windows, SysUtils, Classes, Graphics, Dialogs,
  {$IFDEF EH_LIB_6} Variants, {$ENDIF}
  Db;

type

{ TCompoManEh }

  TCompoManEh = class(TComponent)
  private
  protected
    FVisibleComponentListPos: TStringList;
    procedure DefineProperties(Filer: TFiler); override;
    procedure ReadCompoPoses(Reader: TReader);
    procedure WriteCompoPoses(Writer: TWriter);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

//  TDesignTimeCompoListDeleteNotificationProcEh = procedure(Sender: TCompoManEh);

//var
//  DesignTimeCompoListDeleteNotification: TDesignTimeCompoListDeleteNotificationProcEh;

implementation

{ TCompoManEh }

constructor TCompoManEh.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FVisibleComponentListPos := TStringList.Create;
end;

destructor TCompoManEh.Destroy;
begin
//  if Assigned(DesignTimeCompoListDeleteNotification) then
//    DesignTimeCompoListDeleteNotification(Self);
  if Owner <> nil then Owner.RemoveComponent(Self);
  FreeAndNil(FVisibleComponentListPos);
  inherited Destroy;
end;

procedure TCompoManEh.ReadCompoPoses(Reader: TReader);
begin
  FVisibleComponentListPos.Clear;
  with Reader do
  begin
    ReadListBegin;
    while not Reader.EndOfList do
      FVisibleComponentListPos.Add(ReadString);
    ReadListEnd;
  end;
end;

procedure TCompoManEh.WriteCompoPoses(Writer: TWriter);
var
  I: Integer;
begin
  with Writer do
  begin
    WriteListBegin;
    for I := 0 to FVisibleComponentListPos.Count - 1 do
      WriteString(FVisibleComponentListPos[I]);
    WriteListEnd;
  end;
end;

procedure TCompoManEh.DefineProperties(Filer: TFiler);
begin
  inherited DefineProperties(Filer);
  Filer.DefineProperty('VisibleComponentListPos', ReadCompoPoses, WriteCompoPoses, True);
end;

end.
