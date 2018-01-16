{*******************************************************}
{                                                       }
{                        EhLib v8.1                     }
{                                                       }
{                      SpinGridsEh                      }
{                      Build 8.1.0001                    }
{                                                       }
{   Copyright (c) 2014-2015 by Dmitry V. Bolshakov      }
{                                                       }
{*******************************************************}

{$I EhLib.Inc}

unit PlannerItemDialog;

interface

uses
{$IFDEF CIL}
  EhLibVCLNET,
  WinUtils,
{$ELSE}
  {$IFDEF FPC}
  EhLibLCL, LMessages, LCLType, Win32Extra, MaskEdit,
  {$ELSE}
  EhLibVCL, DBConsts, RTLConsts, Mask,
  {$ENDIF}
{$ENDIF}
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, Dialogs, StdCtrls,
  DBCtrlsEh, PlannersEh, PlannerDataEh,
  DateUtils,
  ComCtrls, ExtCtrls;

type
  TPlannerItemForm = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Bevel1: TBevel;
    Bevel2: TBevel;
    AllDayCheck: TCheckBox;
    OKButton: TButton;
    CancelButton: TButton;
    eTitle: TDBEditEh;
    cbStartTimeEdit: TDBComboBoxEh;
    cbFinishTimeEdit: TDBComboBoxEh;
    eBody: TDBMemoEh;
    Bevel3: TBevel;
    cbxRecource: TDBComboBoxEh;
    Bevel4: TBevel;
    StartDateEdit: TDBDateTimeEditEh;
    EndDateEdit: TDBDateTimeEditEh;
    procedure cbStartTimeEditEnter(Sender: TObject);
    procedure cbStartTimeEditChange(Sender: TObject);
    procedure StartDateEditEnter(Sender: TObject);
    procedure StartDateEditChange(Sender: TObject);
  private
    { Private declarations }
    FDeltaTime: TDateTime;
    FDeltaDate: TDateTime;
  public
    procedure InitForm(Item: TPlannerDataItemEh);
    function FormStarDate: TDateTime;
    function FormEndDate: TDateTime;
  end;

var
  PlannerItemForm: TPlannerItemForm;

function EditPlanItem(Planner: TPlannerControlEh; Item: TPlannerDataItemEh): Boolean;
procedure EditNewItem(Planner: TPlannerControlEh);

implementation

{$R *.dfm}

function EditPlanItem(Planner: TPlannerControlEh; Item: TPlannerDataItemEh): Boolean;
var
  AForm: TPlannerItemForm;
  FDummyCheckPlanItem, FOldPlanItemState: TPlannerDataItemEh;
  ErrorText: String;
  CheckChange: Boolean;
begin
  Result := False;
  AForm := TPlannerItemForm.Create(Application);
  try
    AForm.InitForm(Item);
    if AForm.ShowModal = mrOK then
    begin
      FDummyCheckPlanItem := Item.Source.CreateTmpPlannerItem;
      FOldPlanItemState := Item.Source.CreateTmpPlannerItem;
      try
        FDummyCheckPlanItem.Assign(Item);
        FDummyCheckPlanItem.Title := AForm.eTitle.Text;
        FDummyCheckPlanItem.Body := AForm.eBody.Text;
        FDummyCheckPlanItem.StartTime := AForm.FormStarDate;
        FDummyCheckPlanItem.EndTime := AForm.FormEndDate;
        FDummyCheckPlanItem.AllDay := AForm.AllDayCheck.Checked;
        if AForm.cbxRecource.ItemIndex >= 0 then
          FDummyCheckPlanItem.ResourceID := AForm.cbxRecource.KeyItems[AForm.cbxRecource.ItemIndex];
        FDummyCheckPlanItem.EndEdit(True);

        CheckChange :=
          Planner.CheckPlannerItemInteractiveChanging(
            Planner.ActivePlannerView, Item, FDummyCheckPlanItem, ErrorText);

        if CheckChange then
        begin
          FOldPlanItemState.Assign(Item);
          Item.BeginEdit;
          Item.Assign(FDummyCheckPlanItem);
          Item.EndEdit(True);
          Planner.PlannerItemInteractiveChanged(Planner.ActivePlannerView, Item, FOldPlanItemState);
          Result := True;
        end else
        begin
          ShowMessage(ErrorText);
          Result := False;
        end;
      finally
        FDummyCheckPlanItem.Free;
        FOldPlanItemState.Free;
      end;
    end;
  finally
    AForm.Free;
  end;
end;

procedure EditNewItem(Planner: TPlannerControlEh);
var
  StartTime, EndTime: TDateTime;
  PlanItem: TPlannerDataItemEh;
  AResource: TPlannerResourceEh;
begin
  if Planner.NewItemParams(StartTime, EndTime, AResource) then
  begin
    PlanItem := Planner.PlannerDataSource.NewItem;
    PlanItem.Title := 'New Item';
    PlanItem.Body := '';
    PlanItem.AllDay := False;
    PlanItem.StartTime := StartTime;
    PlanItem.EndTime := EndTime;
    PlanItem.Resource := AResource;
    if EditPlanItem(Planner, PlanItem) then
    begin
      PlanItem.EndEdit(True);
    end else
      PlanItem.EndEdit(False);
  end;
end;

{ TPlannerItemForm }

procedure TPlannerItemForm.StartDateEditEnter(Sender: TObject);
begin
  if not VarIsNull(StartDateEdit.Value) and not VarIsNull(EndDateEdit.Value) then
    try
      FDeltaDate := VarToDateTime(EndDateEdit.Value) - VarToDateTime(StartDateEdit.Value);
    except
      on EConvertError do FDeltaTime := 0;
    end;
end;

procedure TPlannerItemForm.StartDateEditChange(Sender: TObject);
begin
 if (FDeltaDate <> 0) and not VarIsNull(StartDateEdit.Value) then
   EndDateEdit.Value := StartDateEdit.Value + FDeltaDate;
end;

procedure TPlannerItemForm.cbStartTimeEditEnter(Sender: TObject);
begin
  if (cbStartTimeEdit.Text <> '') and (cbFinishTimeEdit.Text <> '') then
    try
      FDeltaTime := StrToTime(cbFinishTimeEdit.Text) - StrToTime(cbStartTimeEdit.Text)
    except
      on EConvertError do FDeltaTime := 0;
    end;
end;

function TPlannerItemForm.FormStarDate: TDateTime;
begin
  Result := Trunc(VarToDateTime(StartDateEdit.Value)) + Frac(StrToTime(cbStartTimeEdit.Text));
end;

function TPlannerItemForm.FormEndDate: TDateTime;
begin
  Result := Trunc(VarToDateTime(EndDateEdit.Value)) + Frac(StrToTime(cbFinishTimeEdit.Text));
end;

procedure TPlannerItemForm.cbStartTimeEditChange(Sender: TObject);
var
  s: String;
  ATime: TDateTime;

  function IsDigit(c: Char): Boolean;
  begin
    Result := CharInSetEh(c, ['0','1','2','3','4','5','6','7','8','9',' '])
  end;

begin
  s := cbStartTimeEdit.Text;
  if (Length(s) = 5) and
     IsDigit(s[1]) and
     IsDigit(s[2]) and
     (s[3] = ':') and
     IsDigit(s[4]) and
     IsDigit(s[5])
  then
   if FDeltaTime <> 0 then
   begin
     ATime := StrToTime(cbStartTimeEdit.Text);
     cbFinishTimeEdit.Text := FormatDateTime('HH:MM', ATime + FDeltaTime);
   end;
end;

procedure TPlannerItemForm.InitForm(Item: TPlannerDataItemEh);
var
  i: Integer;
begin
  eTitle.Text := Item.Title;
  eBody.Text := Item.Body;

  StartDateEdit.OnChange := nil;
  StartDateEdit.Value := Item.StartTime;
  StartDateEdit.OnChange := StartDateEditChange;

//  StartTimeEdit.DateTime := Item.StartTime;
  EndDateEdit.Value := Item.EndTime;
//  EndTimeEdit.DateTime := Item.EndTime;
  AllDayCheck.Checked := Item.AllDay;

  cbStartTimeEdit.OnChange := nil;
  cbStartTimeEdit.Text := FormatDateTime('HH:MM', Item.StartTime);
//  cbStartTimeEdit.OnClosingUp := cbStartTimeEditClosingUp;
  cbStartTimeEdit.OnChange := cbStartTimeEditChange;

  cbFinishTimeEdit.Text := FormatDateTime('HH:MM', Item.EndTime);

  if (Item.Source.Resources.Count > 0) then
  begin
    for i := 0 to Item.Source.Resources.Count-1 do
    begin
      cbxRecource.Items.Add(Item.Source.Resources[i].Name);
      cbxRecource.KeyItems.Add(VarToStr(Item.Source.Resources[i].ResourceID));
    end;
  end;

  cbxRecource.ItemIndex := cbxRecource.KeyItems.IndexOf(VarToStr(Item.ResourceID));
end;

end.
