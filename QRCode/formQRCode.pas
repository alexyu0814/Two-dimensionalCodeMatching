unit formQRCode;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Data.DB,
  Data.Win.ADODB, System.DateUtils, Vcl.ComCtrls, Winapi.ActiveX, StrUtils, Vcl.Mask,
  XLSSheetData5, XLSReadWriteII5, Winapi.ShellAPI, Vcl.ExtCtrls,
  DBGridEhGrouping, ToolCtrlsEh, DBGridEhToolCtrls, DynVarsEh, EhLibVCL, GridsEh,
  DBAxisGridsEh, DBGridEh, ComObj, FileCtrl, Contnrs, RegularExpressions, Vcl.CheckLst,
  Vcl.Buttons, System.Math, IdGlobalProtocols, DBGridEhImpExp, XLSWriteXLSX5,
  System.IniFiles;

const
  WM_MYUSER = WM_USER + 100;

type
  TfrmQRCode = class(TForm)
    ADOQuery1: TADOQuery;
    ADOConnection1: TADOConnection;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    btnStart: TButton;
    lblResult: TLabel;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    btnMatch: TButton;
    Memo1: TMemo;
    btnCreateFile: TButton;
    Memo2: TMemo;
    TabSheet4: TTabSheet;
    GroupBox1: TGroupBox;
    Label23: TLabel;
    Label24: TLabel;
    edtA: TEdit;
    edtB: TEdit;
    btnImportA: TButton;
    btnImportB: TButton;
    btnClearA: TButton;
    btnClearB: TButton;
    OpenDialogA: TOpenDialog;
    OpenDialogB: TOpenDialog;
    XLS: TXLSReadWriteII5;
    TabSheet5: TTabSheet;
    grp1: TGroupBox;
    pnl1: TPanel;
    btnImportCustFile: TButton;
    btnAddCustFile: TButton;
    btnDelCustFile: TButton;
    btnClearCustFile: TButton;
    lbl2: TLabel;
    gridCustFile: TDBGridEh;
    qryCustFile: TADOQuery;
    dsCustFile: TDataSource;
    qryCustFileselected: TBooleanField;
    qryCustFilefileName: TStringField;
    conOA: TADOConnection;
    grp2: TGroupBox;
    mmoCustFile: TMemo;
    qryCustFileisContainFirst: TBooleanField;
    TabCheck: TTabSheet;
    grp3: TGroupBox;
    lbl1: TLabel;
    lbl3: TLabel;
    edtCheckA: TEdit;
    edtCheckB: TEdit;
    btnImportCheckA: TButton;
    btnImportCheckB: TButton;
    btnClearCheckA: TButton;
    btnClearCheckB: TButton;
    chk1: TCheckBox;
    TabMulti: TTabSheet;
    grp4: TGroupBox;
    pnl2: TPanel;
    lbl7: TLabel;
    btnImportMulti: TButton;
    btnAddMulti: TButton;
    btnDelMulti: TButton;
    btnClearMulti: TButton;
    gridMulti: TDBGridEh;
    grp5: TGroupBox;
    mmoMulti: TMemo;
    dsMulti: TDataSource;
    btnRepeatMulti: TButton;
    btnNotRepeatMulti: TButton;
    qryMulti: TADOQuery;
    BooleanField1: TBooleanField;
    StringField1: TStringField;
    BooleanField2: TBooleanField;
    btnAddMultiDir: TButton;
    edtCodeLen: TEdit;
    lbl8: TLabel;
    lblFileCount: TLabel;
    chkAuto: TCheckBox;
    Label25: TLabel;
    edtCustNy: TEdit;
    Label27: TLabel;
    Label28: TLabel;
    Panel1: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label16: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    edtNo: TEdit;
    edtMode: TEdit;
    edtWebAddr: TEdit;
    edtBrandCode: TEdit;
    edtPrinterCode: TEdit;
    edtNy: TEdit;
    edtCount: TEdit;
    edtThreadCount: TEdit;
    edtType: TComboBox;
    edtMax: TMaskEdit;
    edtGdbh: TEdit;
    edtYpbh: TEdit;
    edtYpmc: TEdit;
    Panel2: TPanel;
    Label15: TLabel;
    edtPath: TEdit;
    CheckBox1: TCheckBox;
    Label29: TLabel;
    Label30: TLabel;
    Panel3: TPanel;
    btnPublic: TButton;
    btnRepeat: TButton;
    GroupBox2: TGroupBox;
    mmoLog: TMemo;
    Panel4: TPanel;
    GroupBox3: TGroupBox;
    mmoLog1: TMemo;
    tabServer: TTabSheet;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label26: TLabel;
    hostIP: TEdit;
    DBUser: TEdit;
    DBPass: TEdit;
    DB: TEdit;
    chk3: TCheckBox;
    chk2: TCheckBox;
    edtBatchNo: TEdit;
    lbl6: TLabel;
    btnCheck: TButton;
    GroupBox4: TGroupBox;
    CheckListBoxA: TCheckListBox;
    GroupBox5: TGroupBox;
    CheckListBoxB: TCheckListBox;
    Panel5: TPanel;
    btnUpA: TSpeedButton;
    btnDownA: TSpeedButton;
    Panel6: TPanel;
    btnUpB: TSpeedButton;
    btnDownB: TSpeedButton;
    TabSheet6: TTabSheet;
    GroupBox6: TGroupBox;
    GroupBox7: TGroupBox;
    mmoLog2: TMemo;
    grpA1: TGroupBox;
    lbl4: TLabel;
    lbl5: TLabel;
    edtPath1: TEdit;
    edtSPath1: TEdit;
    btnDb1: TButton;
    btnDbS1: TButton;
    btnClear1: TButton;
    btnSClear1: TButton;
    Check1: TCheckBox;
    CheckS1: TCheckBox;
    lbl9: TLabel;
    lbl10: TLabel;
    edtPath2: TEdit;
    edtSPath2: TEdit;
    btnDb2: TButton;
    btnDbS2: TButton;
    btnClear2: TButton;
    btnSClear2: TButton;
    Check2: TCheckBox;
    CheckS2: TCheckBox;
    grpA2: TGroupBox;
    lbl11: TLabel;
    lbl12: TLabel;
    edtPath3: TEdit;
    edtSPath3: TEdit;
    btnDb3: TButton;
    btnDbS3: TButton;
    btnSClear3: TButton;
    btnClear3: TButton;
    Check3: TCheckBox;
    CheckS3: TCheckBox;
    lbl13: TLabel;
    lbl14: TLabel;
    edtPath4: TEdit;
    edtSPath4: TEdit;
    btnDb4: TButton;
    btnDbS4: TButton;
    btnSClear4: TButton;
    btnClear4: TButton;
    Check4: TCheckBox;
    CheckS4: TCheckBox;
    grpA3: TGroupBox;
    lbl15: TLabel;
    lbl16: TLabel;
    edtPath5: TEdit;
    edtSPath5: TEdit;
    btnDb5: TButton;
    btnDbS5: TButton;
    btnSClear5: TButton;
    btnClear5: TButton;
    Check5: TCheckBox;
    CheckS5: TCheckBox;
    lbl17: TLabel;
    lbl18: TLabel;
    edtPath6: TEdit;
    edtSPath6: TEdit;
    btnDb6: TButton;
    btnDbS6: TButton;
    btnSClear6: TButton;
    btnClear6: TButton;
    Check6: TCheckBox;
    CheckS6: TCheckBox;
    grpA4: TGroupBox;
    lbl19: TLabel;
    lbl20: TLabel;
    edtPath7: TEdit;
    edtSPath7: TEdit;
    btnDb7: TButton;
    btnDbS7: TButton;
    btnSClear7: TButton;
    btnClear7: TButton;
    Check7: TCheckBox;
    CheckS7: TCheckBox;
    lbl21: TLabel;
    lbl22: TLabel;
    edtPath8: TEdit;
    edtSPath8: TEdit;
    btnDb8: TButton;
    btnDbS8: TButton;
    btnSClear8: TButton;
    btnClear8: TButton;
    Check8: TCheckBox;
    CheckS8: TCheckBox;
    grpA5: TGroupBox;
    lbl23: TLabel;
    lbl24: TLabel;
    edtPath9: TEdit;
    edtSPath9: TEdit;
    btnDb9: TButton;
    btnDbS9: TButton;
    btnSClear9: TButton;
    btnClear9: TButton;
    Check9: TCheckBox;
    CheckS9: TCheckBox;
    lbl25: TLabel;
    lbl26: TLabel;
    edtPath10: TEdit;
    edtSPath10: TEdit;
    btnDb10: TButton;
    btnDbS10: TButton;
    btnClear10: TButton;
    btnSClear10: TButton;
    Check10: TCheckBox;
    CheckS10: TCheckBox;
    btnCheck2: TButton;
    btnClear: TButton;
    adoCheck: TADOQuery;
    GroupBox8: TGroupBox;
    edtTxt: TEdit;
    btnTxt: TButton;
    Label31: TLabel;
    Button2: TButton;
    adoCheck3: TADOQuery;
    StatusBar1: TStatusBar;
    ProgressBar1: TProgressBar;
    XLS2: TXLSReadWriteII5;
    btnImportMulti2: TButton;
    procedure btnStartClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnMatchClick(Sender: TObject);
    procedure btnCreateFileClick(Sender: TObject);
    procedure edtNoExit(Sender: TObject);
    procedure btnImportAClick(Sender: TObject);
    procedure btnClearAClick(Sender: TObject);
    procedure btnImportBClick(Sender: TObject);
    procedure btnClearBClick(Sender: TObject);
    procedure btnRepeatClick(Sender: TObject);
    procedure btnAddCustFileClick(Sender: TObject);
    procedure btnImportCustFileClick(Sender: TObject);
    procedure btnDelCustFileClick(Sender: TObject);
    procedure btnClearCustFileClick(Sender: TObject);
    procedure qryCustFileAfterOpen(DataSet: TDataSet);
    procedure gridCustFileMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure btnImportCheckAClick(Sender: TObject);
    procedure btnClearCheckAClick(Sender: TObject);
    procedure btnClearCheckBClick(Sender: TObject);
    procedure btnImportCheckBClick(Sender: TObject);
    procedure chk3Click(Sender: TObject);
    procedure chk2Click(Sender: TObject);
    procedure btnAddMultiClick(Sender: TObject);
    procedure qryMultiAfterOpen(DataSet: TDataSet);
    procedure gridMultiMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure btnDelMultiClick(Sender: TObject);
    procedure btnClearMultiClick(Sender: TObject);
    procedure btnImportMultiClick(Sender: TObject);
    procedure btnAddMultiDirClick(Sender: TObject);
    procedure btnRepeatMultiClick(Sender: TObject);
    procedure btnNotRepeatMultiClick(Sender: TObject);
    procedure btnPublicClick(Sender: TObject);
    procedure chkAutoClick(Sender: TObject);
    procedure btnCheck2Click(Sender: TObject);
    procedure btnConnectClick(Sender: TObject);
    procedure CheckListBoxAClickCheck(Sender: TObject);
    procedure CheckListBoxBClickCheck(Sender: TObject);
    procedure btnUpAClick(Sender: TObject);
    procedure btnUpBClick(Sender: TObject);
    procedure btnDownAClick(Sender: TObject);
    procedure btnDownBClick(Sender: TObject);
    procedure btnDb1Click(Sender: TObject);
    procedure btnDbS1Click(Sender: TObject);
    procedure btnClear1Click(Sender: TObject);
    procedure btnSClear1Click(Sender: TObject);
    procedure Check1Click(Sender: TObject);
    procedure CheckS1Click(Sender: TObject);
    procedure btnClearClick(Sender: TObject);
    procedure btnCheckClick(Sender: TObject);
    procedure btnTxtClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure btnImportMulti2Click(Sender: TObject);
  private
    FCheckListBoxSelectCountA, FCheckListBoxSelectCountB: Integer;
    FIsImported: Boolean;
    function ShowAsk(MHandle: Thandle; tText: Pchar): Boolean;
    procedure Searchfile(path: PChar; fileExt: string; fileList: TStringList);
    function MyCheck2(IndexNo: Integer; FileA, FileB: string; DelimA, DelimB: Boolean): Boolean;
    function MyCheck3(FileA, FileB, Delim: string): Boolean;
    function AutoCheck(FileA, FileB: string; DelimA, DelimB: Boolean): Boolean;
    function LastPos(const SubStr, Str: string): Integer;
    function DbGridEhToExcel(fileNameout: string): string;
    function DataToExcel(fileNamein, fileNameout: string): string;
    procedure MyWndProc(var Msg: Tmessage); message WM_MYUSER;
    { Private declarations }
  public
    procedure WmDropFiles(var Msg: TMessage); message WM_DROPFILES; //处理WM_DROPFILES消息的过程
    { Public declarations }
  end;

var
  frmQRCode: TfrmQRCode;
  Cs: TRTLCriticalSection;

implementation
{$R *.dfm}

//文件名生成规则：主表单据号(工单号【】,印品编号【】,印品名称【】,盒型【】,二维码数量【】)

procedure TfrmQRCode.btnImportCheckAClick(Sender: TObject);
var
  iQuery: TADOQuery;
  tableName, isql: string;
  i: Integer;
begin
  if FCheckListBoxSelectCountA <= 0 then
  begin
    ShowMessage('请选择子文件的格式！');
    Exit;
  end;

  OpenDialogA.FileName := '文本文件(子文件)';
  OpenDialogA.Filter := '*.txt|*.txt';
  iQuery := TADOQuery.Create(Self);
  try
    if OpenDialogA.Execute then
    begin
      edtCheckA.Text := OpenDialogA.FileName;
      with iQuery do
      begin
        CommandTimeout := 0;
        Connection := ADOConnection1;
        Close;
        isql := 'if Exists(select 1 from sysobjects where name=''tb_child'' and xtype=''U'') drop table tb_child;';
        isql := isql + 'create table tb_child(';
        //根据子文件格式动态创建列
        for i := 0 to CheckListBoxA.Items.Count - 1 do
        begin
          if CheckListBoxA.Checked[i] then
          begin
            isql := isql + CheckListBoxA.Items.Strings[i] + ' varchar(100),';
          end;
        end;
        if RightStr(isql, 1) = ',' then
          isql := Copy(isql, 1, length(isql) - 1);

        isql := isql + ');BULK INSERT tb_child FROM ' + QuotedStr(Trim(edtCheckA.Text)) + 'WITH(FIELDTERMINATOR ='' '',ROWTERMINATOR =''\n'')';
        SQL.Add(isql);
        ExecSQL;
        Close;
        SQL.Text := 'select count(1) as count from tb_child';
        Open;
        mmoLog1.Lines.Add('子文件总记录[' + FieldByName('count').AsString + ']行');
      end;
    end;
  finally
    iQuery.Free;
    btnCheck.Enabled := True;
  end;
end;

procedure TfrmQRCode.btnImportCheckBClick(Sender: TObject);
var
  iQuery: TADOQuery;
  tableName, s, fileName, isql: string;
  sheetIndex, FirstRow, Row, Col: Integer;
  list: TStringList;
  i: Integer;
begin
  if FCheckListBoxSelectCountB <= 0 then
  begin
    ShowMessage('请选择父文件的格式！');
    Exit;
  end;

  OpenDialogB.FileName := '文本文件';
  OpenDialogB.Filter := '*.txt;*.xls;*.xlsx|*.txt;*.xls;*.xlsx';
  iQuery := TADOQuery.Create(Self);
  list := TStringList.Create;
  try
    if OpenDialogB.Execute then
    begin
      edtCheckB.Text := OpenDialogB.FileName;
      with iQuery do
      begin
        iQuery.CommandTimeout := 0;
        XLS.Filename := Trim(edtCheckB.Text);
        //导入xls、xlsx
        if (ExtractFileExt(XLS.Filename) = '.xls') or (ExtractFileExt(XLS.Filename) = '.xlsx') then
        begin
          chk1.Visible := True;
          chk1.Checked := True;
          //  XLS.DirectRead := True; //取消此参数
          XLS.Read;
          for sheetIndex := 0 to 99 do //最多可导入100个Sheet
          begin
            try
              if not XLS.Sheets[sheetIndex].IsEmpty then
              begin
                if chk1.Checked then
                  FirstRow := 1
                else
                  FirstRow := XLS.Sheets[sheetIndex].FirstRow;
                for Row := FirstRow to XLS.Sheets[sheetIndex].LastRow do
                begin
                  s := '';
                  for Col := XLS.Sheets[sheetIndex].FirstCol to XLS.Sheets[sheetIndex].LastCol do
                  begin
                    s := s + ' ' + XLS.Sheets[sheetIndex].AsString[Col, Row];
                  end;
                  list.Append(s);
                end;
              end;
            except
              Break;  //屏蔽异常
            end;
          end;
          fileName := ExtractFilePath(Trim(edtCheckB.Text)) + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '(临时文件).txt';
          //保存txt文件
          list.SaveToFile(fileName);
          tableName := Trim(ChangeFileExt(ExtractFileName(XLS.Filename), ''));
          //从txt文件保存到数据库
          Connection := ADOConnection1;
          Close;
          isql := 'if Exists(select 1 from sysobjects where name=''tb_father'' and xtype=''U'') drop table tb_father;';
          //根据子文件格式动态创建列
          isql := isql + 'create table tb_father(';
          for i := 0 to CheckListBoxB.Items.Count - 1 do
          begin
            if CheckListBoxB.Checked[i] then
            begin
              isql := isql + CheckListBoxB.Items.Strings[i] + ' varchar(100),';
            end;
          end;
          if RightStr(isql, 1) = ',' then
            isql := Copy(isql, 1, length(isql) - 1);
          isql := isql + ');BULK INSERT tb_father FROM ' + QuotedStr(fileName) + 'WITH(FIELDTERMINATOR ='' '',ROWTERMINATOR =''\n'')';
          try
            SQL.Add(isql);
            ExecSQL;
          except
            on e: Exception do
            begin
              mmoCustFile.Lines.Add('导入文件【' + fileName + '】失败，异常信息：' + e.message);
              btnCheck.Enabled := True;
              exit;
            end;
          end;
        end
        else if (ExtractFileExt(XLS.Filename) = '.txt') then
        begin
          chk1.Visible := False;
          chk1.Checked := False;
          Connection := ADOConnection1;
          Close;
          isql := 'if Exists(select 1 from sysobjects where name=''tb_father'' and xtype=''U'') drop table tb_father;';
          //根据子文件格式动态创建列
          isql := isql + 'create table tb_father(';
          for i := 0 to CheckListBoxB.Items.Count - 1 do
          begin
            if CheckListBoxB.Checked[i] then
            begin
              isql := isql + CheckListBoxB.Items.Strings[i] + ' varchar(100),';
            end;
          end;
          if RightStr(isql, 1) = ',' then
            isql := Copy(isql, 1, length(isql) - 1);
          isql := isql + ');BULK INSERT tb_father FROM ' + QuotedStr(Trim(edtCheckB.Text)) + 'WITH(FIELDTERMINATOR ='' '',ROWTERMINATOR =''\n'')';
          SQL.Add(isql);
//          SQL.Add('Update tb_father set 条码内容=子批次号,子批次号=''''');
          ExecSQL;
        end;
        Close;
        SQL.Text := 'select count(1) as count from tb_father';
        Open;
        mmoLog1.Lines.Add('父文件总记录[' + FieldByName('count').AsString + ']行');
      end;
    end;
  finally
    btnCheck.Enabled := True;
    iQuery.Free;
    list.Free;
  end;
end;

 //返回子字符串在字符串最后一次出现的位置 alex add 2017-09-18
function TfrmQRCode.LastPos(const SubStr, Str: string): Integer;
var
  Idx: Integer; // an index of SubStr in Str
begin
  Result := 0;
  Idx := StrUtils.PosEx(SubStr, Str);
  if Idx = 0 then
    Exit;
  while Idx > 0 do
  begin
    Result := Idx;
    Idx := StrUtils.PosEx(SubStr, Str, Idx + 1);
  end;
end;

procedure TfrmQRCode.btnImportCustFileClick(Sender: TObject);
var
  Col, Row, sheetIndex, FirstRow, Flag, index: integer;
  s, s2, isql, fileName, fileName2, tableName: string;
  SPath, Sfilename, txtname, txtpath: string;
  list: TStringList;
  list2: TStringList;
  ListArray: array[1..6] of TStringList;
  i: Integer;
  sDateTime: TDateTime;
begin
  if not qryCustFile.Active then
  begin
    ShowMessage('没有选择要导入的文件');
    Exit;
  end;
  if qryCustFile.IsEmpty then
    Exit;


  //alex yu edit 2017 09-18
  {
    第一部分为根据导入的excel生成四个文件(txt)：临时，检_，喷_,品检_
  }
  mmoCustFile.Lines.Clear;
  mmoCustFile.Lines.Add('导入可能需要几分钟，正在导入中，请耐心等待...');
  sDateTime := now;
  qryCustFile.DisableControls;
  list := TStringList.Create;
  list2 := TStringList.Create;
  try
    qryCustFile.First;
    while not qryCustFile.Eof do
    begin
      Application.ProcessMessages;
      list.Clear;
      list2.Clear;
      if qryCustFile.FieldByName('selected').AsBoolean then
      begin
        XLS.Filename := qryCustFile.FieldByName('fileName').AsString;
        if (ExtractFileExt(XLS.Filename) = '.xls') or (ExtractFileExt(XLS.Filename) = '.xlsx') then
        begin
//          XLS.DirectRead := True; //取消此参数
          XLS.Read;
          for sheetIndex := 0 to 99 do //最多可导入100个Sheet
          begin
            try
              if not XLS.Sheets[sheetIndex].IsEmpty then
              begin
                if qryCustFile.FieldByName('isContainFirst').AsBoolean then
                  FirstRow := 1
                else
                  FirstRow := XLS.Sheets[sheetIndex].FirstRow;
                for Row := FirstRow to XLS.Sheets[sheetIndex].LastRow do
                begin
                  s := '';
                  s2 := '';
                  Flag := 1;
                  for Col := XLS.Sheets[sheetIndex].FirstCol to XLS.Sheets[sheetIndex].LastCol do
                  begin
                    s := s + ' ' + XLS.Sheets[sheetIndex].AsString[Col+1 , Row];
                    //s2:=StringReplace(s,' ',',',[rfReplaceAll,rfIgnoreCase]);
                    //品检格式 条码内容,验证码
                    //其他格式 子批次号 序号 条码内容 验证码  （注意已空格隔开并且最后有两个空格，凑足四个栏位）
                    if Flag <= 2 then
                    begin
                      if s2 = '' then
                        s2 := XLS.Sheets[sheetIndex].AsString[Col + 1, Row]
                      else
                        s2 := s2 + ',' + XLS.Sheets[sheetIndex].AsString[Col + 1, Row];
                    end;
                    Inc(Flag);
                    //  s:=copy(s,Pos(' ',TrimLeft(s))+2,length(s));
                  end;
                  list.Append(TrimLeft(s));
                  list2.Append(TrimLeft(s2));
                end;
              end;
            except
              Break;  //屏蔽异常
            end;
          end;
          //fileName := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '(无序号文件).txt';
          fileName := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + '喷_' + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '.txt';
          //保存txt文件
          list.SaveToFile(fileName);
          fileName2 := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + '检_' + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '.txt';
          CopyFile(PChar(fileName), PChar(fileName2), False);
          fileName := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + '品检_' + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '.txt';
          //保存txt文件
          list2.SaveToFile(fileName);
        end;

        //导入xls、xlsx
        list.Clear;
        if (ExtractFileExt(XLS.Filename) = '.xls') or (ExtractFileExt(XLS.Filename) = '.xlsx') then
        begin
//          XLS.DirectRead := True; //取消此参数
          XLS.Read;
          for sheetIndex := 0 to 99 do //最多可导入100个Sheet
          begin
            try
              if not XLS.Sheets[sheetIndex].IsEmpty then
              begin
                if qryCustFile.FieldByName('isContainFirst').AsBoolean then
                  FirstRow := 1
                else
                  FirstRow := XLS.Sheets[sheetIndex].FirstRow;
                for Row := FirstRow to XLS.Sheets[sheetIndex].LastRow do
                begin
                  s := '';
                  for Col := XLS.Sheets[sheetIndex].FirstCol to XLS.Sheets[sheetIndex].LastCol do
                  begin
                    s := s + ' ' + XLS.Sheets[sheetIndex].AsString[Col, Row];
                  //  s:=copy(s,Pos(' ',TrimLeft(s))+2,length(s));
                  end;
                  list.Append(s);
                end;
              end;
            except
              Break;  //屏蔽异常
            end;
          end;
          fileName := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '(临时文件).txt';
          //保存txt文件
          list.SaveToFile(fileName);
          tableName := Trim(ChangeFileExt(ExtractFileName(XLS.Filename), ''));
          //从txt文件保存到数据库

          isql := 'if not Exists(select 1 from sysobjects where name=' + QuotedStr(tableName) + ' and xtype=''U'') create table [' + tableName + '](子批次号 varchar(100),序号 varchar(50),条码内容 varchar(100),验证码 varchar(100)) else TRUNCATE TABLE [' + tableName + '];';
          isql := isql + 'BULK INSERT [' + tableName + '] FROM ' + QuotedStr(fileName) + 'WITH(FIELDTERMINATOR ='' '',ROWTERMINATOR =''\n'')';
          try
            ADOConnection1.Execute(isql);
          except
            on e: Exception do
            begin
              mmoCustFile.Lines.Add('导入文件【' + fileName + '】失败，异常信息：' + e.message);
              exit;
            end;
          end;
        end
        else if (ExtractFileExt(XLS.Filename) = '.txt') then //导入txt
        begin
          fileName := qryCustFile.FieldByName('fileName').AsString;
          tableName := Trim(ChangeFileExt(ExtractFileName(qryCustFile.FieldByName('fileName').AsString), ''));
          //从txt文件保存到数据库
          isql := 'if not Exists(select 1 from sysobjects where name=' + QuotedStr(tableName) + ' and xtype=''U'') create table [' + tableName + '](条码内容 varchar(100),验证码 varchar(100)) else TRUNCATE TABLE [' + tableName + '];';
          isql := isql + 'BULK INSERT [' + tableName + '] FROM ' + QuotedStr(fileName) + 'WITH(FIELDTERMINATOR ='' '',ROWTERMINATOR =''\n'')';
          try
            ADOConnection1.Execute(isql);
          except
            on e: Exception do
            begin
              mmoCustFile.Lines.Add('导入文件【' + fileName + '】失败，异常信息：' + e.message);
              exit;
            end;
          end;
        end;
        mmoCustFile.Lines.Add('导入文件【' + fileName + '】成功!!!');
      end;
      qryCustFile.Next;
    end;
    mmoCustFile.Lines.Add('恭喜，文件已全部导入成功!!!');
  finally
    qryCustFile.EnableControls;
    list.Free;
    FreeAndNil(list2);
  end;
  for i := Low(ListArray) to High(ListArray) do
    ListArray[i] := TStringList.Create;
  ListArray[2].Clear;
  ListArray[4].Clear;
  ListArray[6].Clear;
  try
    mmoCustFile.Lines.Add('现在开始匹配文件!!!');
  //ShowMessage(IntToStr(qryCustFile.RecordCount));
    qryCustFile.First;
    while not qryCustFile.Eof do
    begin
    //ShowMessage('aaaaaa');
    //取最右边\的位数
      index := LastPos('\', qryCustFilefileName.Value);
    //取文件路径
      spath := Copy(qryCustFilefileName.Value, 0, index);
    //取文件名
      Sfilename := Copy(qryCustFilefileName.Value, index + 1, Length(qryCustFilefileName.Value));
    //取去扩展名的文件名。用于后面生成txt文件名（只需加扩展名，已经加了.）
      txtname := Copy(Sfilename, 0, LastPos('.', Sfilename));

      txtpath := spath + '检_' + txtname + 'txt';
      {
      ListArray[1].Clear;
      ListArray[1].LoadFromFile(txtpath);
      ListArray[2].AddStrings(ListArray[1]);
      }
      autoCheck(txtpath, qryCustFilefileName.Text, true, True);

      txtpath := spath + '喷_' + txtname + 'txt';
      {
      ListArray[3].Clear;
      ListArray[3].LoadFromFile(txtpath);
      ListArray[4].AddStrings(ListArray[3]);
      }
      autoCheck(txtpath, qryCustFilefileName.Text, true, True);

      txtpath := spath + '品检_' + txtname + 'txt';
      ListArray[5].Clear;
      ListArray[5].LoadFromFile(txtpath);
      ListArray[6].AddStrings(ListArray[5]);
      autoCheck(txtpath, qryCustFilefileName.Text, False, True);
      qryCustFile.Next;
    end;
    //ListArray[2].SaveToFile(spath+'检_合并.txt');
    //ListArray[4].SaveToFile(spath+'喷_合并.txt');
    ListArray[6].SaveToFile(spath + '品检_合并.txt');
    mmoCustFile.Lines.Add('恭喜，已全部导入和对比完成，共耗时：' + IntToStr(SecondsBetween(sDateTime, Now)) + '秒！');
  finally
    for i := Low(ListArray) to High(ListArray) do
      FreeAndNil(ListArray[i]);
  end;
end;

procedure TfrmQRCode.btnImportMultiClick(Sender: TObject);
var
  Col, Row, sheetIndex, FirstRow: integer;
  s, isql, fileName, tableName: string;
  list: TStringList;
  iXLS: TXLSReadWriteII5;
  sDate, eDate: TDateTime;
  iQuery: TADOQuery;
begin
  FIsImported := false;
  if not qryMulti.Active then
  begin
    ShowMessage('没有选择要导入的文件');
    Exit;
  end;
  if qryMulti.IsEmpty then
  begin
    ShowMessage('没有选择要导入的文件');
    Exit;
  end;
  mmoMulti.Lines.Clear;
  qryMulti.DisableControls;
  list := TStringList.Create;
  iXLS := TXLSReadWriteII5.Create(Self);
  try
    sDate := Now;
    ADOConnection1.Execute('if object_id(''tempdb..##tb_Multi'') is not null drop table ##tb_Multi create table ##tb_Multi(code varchar(200))');
    qryMulti.First;
    while not qryMulti.Eof do
    begin
      Application.ProcessMessages;
      if qryMulti.FieldByName('selected').AsBoolean then
      begin
        XLS.Filename := qryMulti.FieldByName('fileName').AsString;
        //导入xls、xlsx
        if (ExtractFileExt(XLS.Filename) = '.xls') or (ExtractFileExt(XLS.Filename) = '.xlsx') then
        begin
        //  XLS.DirectRead := True; //取消此参数
          {XLS.Read;
          for sheetIndex := 0 to 99 do //最多可导入100个Sheet
          begin
            try
              if not XLS.Sheets[sheetIndex].IsEmpty then
              begin
                if qryCustFile.FieldByName('isContainFirst').AsBoolean then
                  FirstRow := 1
                else
                  FirstRow := XLS.Sheets[sheetIndex].FirstRow;
                for Row := FirstRow to XLS.Sheets[sheetIndex].LastRow do
                begin
                  s := '';
                  for Col := XLS.Sheets[sheetIndex].FirstCol to XLS.Sheets[sheetIndex].LastCol do
                  begin
                    s := s + ' ' + XLS.Sheets[sheetIndex].AsString[Col, Row];
                  end;
                  list.Append(s);
                end;
              end;
            except
              Break;  //屏蔽异常
            end;
          end;
          fileName := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '(临时文件).txt';
          //保存txt文件
          list.SaveToFile(fileName);
          tableName := Trim(ChangeFileExt(ExtractFileName(XLS.Filename), ''));
          //从txt文件保存到数据库
          isql := 'if not Exists(select 1 from sysobjects where name=' + QuotedStr(tableName) + ' and xtype=''U'') create table [' + tableName + '](子批次号 varchar(100),序号 varchar(50),条码内容 varchar(100),验证码 varchar(100)) else TRUNCATE TABLE [' + tableName + '];';
          isql := isql + 'BULK INSERT [' + tableName + '] FROM ' + QuotedStr(fileName) + 'WITH(FIELDTERMINATOR ='' '',ROWTERMINATOR =''\n'')';
          try
            ADOConnection1.Execute(isql);
          except
            on e: Exception do
            begin
              mmoCustFile.Lines.Add('导入文件【' + fileName + '】失败，异常信息：' + e.message);
              exit;
            end;
          end; }
        end
        else if (ExtractFileExt(XLS.Filename) = '.txt') then //导入txt
        begin
          fileName := qryMulti.FieldByName('fileName').AsString;
          //从txt文件保存到数据库
          isql := 'BULK INSERT ##tb_multi FROM ' + QuotedStr(fileName) + 'WITH(FIELDTERMINATOR ='' '',ROWTERMINATOR =''\n'')';
          try
            ADOConnection1.Execute(isql);
          except
            on e: Exception do
            begin
              mmoMulti.Lines.Add('导入文件【' + fileName + '】失败，异常信息：' + e.message);
              exit;
            end;
          end;
        end;
        mmoMulti.Lines.Add('导入文件【' + fileName + '】成功!!!');
      end;
      qryMulti.Next;
    end;
    ADOConnection1.Execute('UPDATE ##tb_Multi SET code = SUBSTRING(code,CHARINDEX(''http'',code),' + Trim(edtCodeLen.Text) + ')');
    eDate := Now;
    iQuery := TADOQuery.Create(Self);
    with iQuery do
    begin
      CommandTimeout := 0;
      Connection := ADOConnection1;
      Close;
      SQL.Text := 'select count(*) as count from ##tb_Multi';
      Open;
    end;
    mmoMulti.Lines.Add('文件已全部导入成功，共导入【' + iQuery.FieldByName('count').AsString + '】笔，耗时：' + IntToStr(SecondsBetween(eDate, sDate)) + '秒!!!');
    FIsImported := True;
  finally
    qryMulti.EnableControls;
    list.Free;
    iXLS.Free;
    iQuery.Free;
  end;
end;

procedure TfrmQRCode.btnImportMulti2Click(Sender: TObject);
var
  Hash1: THashedStringList;
  i: Integer;
  temp: string;
  eDate, sDate:TDateTime;
begin
  if not qryMulti.Active then
  begin
    ShowMessage('没有选择要导入的文件');
    Exit;
  end;
  if qryMulti.IsEmpty then
  begin
    ShowMessage('没有选择要导入的文件');
    Exit;
  end;
  sDate:=Now;
  Hash1 := THashedStringList.Create;
  try
    qryMulti.First;
    while not qryMulti.Eof do
    begin
      Application.ProcessMessages;
      if qryMulti.FieldByName('selected').AsBoolean then
      begin
        Hash1.Clear;
        Hash1.LoadFromFile(qryMulti.FieldByName('fileName').AsString);
        Hash1.SaveToFile(qryMulti.FieldByName('fileName').AsString);
        mmoMulti.Lines.Add('补回车符号完成【'+qryMulti.FieldByName('fileName').AsString+'】');
      end;
      qryMulti.Next;
    end;
    eDate:=Now;
    mmoMulti.Lines.Add('补回车符号完成，共耗时：' + IntToStr(SecondsBetween(eDate, sDate)) + '秒!!!');
  finally
    FreeAndNil(Hash1);
  end;
end;

procedure TfrmQRCode.btnAddMultiClick(Sender: TObject);
var
  iQuery: TADOQuery;
  OpenDialog: TOpenDialog;
  i: Integer;
begin
  OpenDialog := TOpenDialog.Create(Self);
  OpenDialog.FileName := '文本文件';
  OpenDialog.Filter := '*.txt|*.txt';
  OpenDialog.Options := [ofHideReadOnly, ofEnableSizing, ofAllowMultiSelect];
  iQuery := TADOQuery.Create(Self);
  try
    if OpenDialog.Execute then
    begin
      for i := 0 to OpenDialog.Files.Count - 1 do
      begin
        with qryMulti do
        begin
          //如果存在相同的文件
          DisableControls;
          try
            First;
            while not Eof do
            begin
              if Trim(FieldByName('fileName').AsString) = Trim(OpenDialog.Files[i]) then
              begin
                ShowMessage('已经存在文件【' + Trim(FieldByName('fileName').AsString) + '】!!!');
                Exit;
              end;
              Next;
            end;
          finally
            EnableControls;
          end;

          if not (State in [dsEdit, dsInsert]) then
            Append;
          FieldByName('selected').AsBoolean := True;
          FieldByName('fileName').AsString := OpenDialog.files[i];
          FieldByName('isContainFirst').AsBoolean := True;
          if not (State in [dsBrowse]) then
            Post;
        end;
      end;
    end;
  finally
    iQuery.Free;
    OpenDialog.Free;
  end;
end;

procedure TfrmQRCode.btnClearAClick(Sender: TObject);
begin
  edtA.Clear;
end;

procedure TfrmQRCode.btnClearBClick(Sender: TObject);
begin
  edtB.Clear;
end;

procedure TfrmQRCode.btnClearCheckAClick(Sender: TObject);
begin
  edtCheckA.Clear;
end;

procedure TfrmQRCode.btnClearCheckBClick(Sender: TObject);
begin
  edtCheckB.Clear;
end;

procedure TfrmQRCode.btnClearCustFileClick(Sender: TObject);
begin
  qryCustFile.DisableControls;
  try
    qryCustFile.First;
    while not qryCustFile.Eof do
    begin
      qryCustFile.Delete;
    end;
  finally
    qryCustFile.EnableControls;
  end;
  edtTxt.Clear;
end;

procedure TfrmQRCode.btnClearMultiClick(Sender: TObject);
begin
  qryMulti.DisableControls;
  try
    qryMulti.First;
    while not qryMulti.Eof do
    begin
      qryMulti.Delete;
    end;
  finally
    qryMulti.EnableControls;
  end;
end;

procedure TfrmQRCode.btnConnectClick(Sender: TObject);
begin
  with ADOConnection1 do
  begin
    Connected := False;
   // ConnectionString := 'Provider=SQLOLEDB.1;Password=' + trim(DBPass.Text) + ';Persist Security Info=True;User ID=' + trim(DBUser.Text) + ';Initial Catalog=' + trim(DB.Text) + ';Data Source=' + trim(hostIP.Text) + '';
    ConnectionString:='Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;User ID=sa;Initial Catalog=DB_RQCode;Data Source=.;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;'+'Workstation ID=.;Use Encryption for Data=False;Tag with column collation when possible=False';
    Connected := True;
  end;
end;

procedure TfrmQRCode.btnCreateFileClick(Sender: TObject);
var
  isql, path, iTable, ppMc: string;
  sDate, eDate: TDateTime;
  iQuery: TADOQuery;
  filename: string;
begin
  if Trim(edtNo.Text) = '' then
  begin
    ShowMessage('主表单据号不能为空!!!');
    Exit;
  end;
  if Trim(edtPath.Text) = '' then
  begin
    ShowMessage('生成文件路径不能为空!!!');
    Exit;
  end;
  Memo2.Lines.Clear;

  if not DirectoryExists(path) then
    CreateDirectory(PWideChar(path), nil);

  if Trim(edtBrandCode.Text) = '01' then
    ppMc := '双喜（经典工坊）'
  else if Trim(edtBrandCode.Text) = '02' then
    ppMc := '双喜（彩悦）'
  else if Trim(edtBrandCode.Text) = '03' then
    ppMc := '双喜（喜百年）'
  else if Trim(edtBrandCode.Text) = '04' then
    ppMc := '双喜（百年经典）'
  else if Trim(edtBrandCode.Text) = '05' then
    ppMc := '双喜（和喜）'
  else if Trim(edtBrandCode.Text) = '06' then
    ppMc := '双喜（国喜细支）'
  else if Trim(edtBrandCode.Text) = '07' then
    ppMc := '双喜（国喜）双喜（彩悦）';
  path := Trim(edtPath.Text);
  filename := trim(iTable) + '(工单号【' + edtGdbh.Text + '】,印品编号【' + edtYpbh.Text + '】,印品名称【' + edtYpmc.Text + '】,盒型【' + edtType.Text + '】,二维码数量【' + edtCount.Text + '】)' + '.txt';

  //遍历本主表单据号下的多版本数据表
  iQuery := TADOQuery.Create(Self);
  try
    with iQuery do
    begin
      CommandTimeout := 0;
      Connection := ADOConnection1;
      Close;
      SQL.Text := 'SELECT name FROM sysobjects WHERE NAME LIKE ' + QuotedStr(Trim(edtNo.Text) + '%') + ' AND xtype=''U''';
      Open;

      isql := '';
      first;
      while not Eof do
      begin
        sDate := Now;
        iTable := FieldByName('name').AsString;
        if CheckBox1.Checked then
        begin  //文件名生成规则：主表单据号(工单号【】,印品编号【】,印品名称【】,盒型【】,二维码数量【】)
          isql := 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('SELECT wz+ppdm+yscdm+ny+sjm+'' ''+jym FROM DB_QRCode..' + trim(iTable)) + ',' + QuotedStr(path + '\' + iTable + filename) + ';'
        end
        else
        begin
          isql := 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('SELECT wz+ppdm+yscdm+ny+sjm+'' ''+jym FROM DB_QRCode..' + trim(iTable)) + ',' + QuotedStr(path + '\' + iTable + filename) + ';';
        end;

        try
          ADOConnection1.Execute(isql);
        except
          on e: Exception do
          begin
            Memo1.Lines.Add('生成文件【' + path + '\' + iTable + filename + '】失败，异常信息：' + e.Message);
            Exit;
          end;
        end;
        eDate := Now;
        Memo2.Lines.Add('生成文件【' + path + '\' + iTable + filename + '】成功，耗时：' + IntToStr(SecondsBetween(eDate, sDate)) + '秒!!!');
        Next;
      end;
    end;
  finally
    iQuery.Free;
  end;
end;

procedure TfrmQRCode.btnDb1Click(Sender: TObject);
var
  path: string;
  flag: Integer;
begin
  OpenDialogA.FileName := '文本文件';
  OpenDialogA.Filter := '*.txt|*.txt';
  if OpenDialogA.Execute then
  begin
    path := Trim(OpenDialogA.Files.Text);
    //path:=StringReplace(path,#13#10,'',[rfReplaceAll,rfIgnoreCase]);//文件名中有回车符号，后面sql语句报错
    flag := TButton(Sender).Tag;
    if FindComponent('edtPath' + inttostr(flag)) <> nil then
    begin
      TEdit(FindComponent('edtPath' + inttostr(flag))).Text := path;
      TEdit(FindComponent('edtPath' + inttostr(flag))).Hint := path;
    end;
  end;
end;

procedure TfrmQRCode.btnDbS1Click(Sender: TObject);
var
  path: string;
  flag: Integer;
begin
  OpenDialogA.FileName := 'Excel文件';
  OpenDialogA.Filter := '*.xls|*.xlsx';
  if OpenDialogA.Execute then
  begin
    path := Trim(OpenDialogA.Files.Text);
    flag := TButton(Sender).Tag;
    if FindComponent('edtSPath' + inttostr(flag)) <> nil then
    begin
      TEdit(FindComponent('edtSPath' + inttostr(flag))).Text := path;
      TEdit(FindComponent('edtSPath' + inttostr(flag))).Hint := path;
    end;
  end;
end;

procedure TfrmQRCode.btnClear1Click(Sender: TObject);
var
  flag: Integer;
begin
  flag := TButton(Sender).Tag;
  if FindComponent('edtPath' + inttostr(flag)) <> nil then
    TEdit(FindComponent('edtPath' + inttostr(flag))).Text := '';
end;

procedure TfrmQRCode.btnSClear1Click(Sender: TObject);
var
  flag: Integer;
begin
  flag := TButton(Sender).Tag;
  if FindComponent('edtSPath' + inttostr(flag)) <> nil then
    TEdit(FindComponent('edtSPath' + inttostr(flag))).Text := '';
end;

procedure TfrmQRCode.Check1Click(Sender: TObject);
var
  flag: Integer;
begin
  flag := TCheckBox(Sender).Tag;
  if TCheckBox(Sender).Checked then
  begin
    if FindComponent('CheckS' + inttostr(flag)) <> nil then
      TCheckBox(FindComponent('CheckS' + inttostr(flag))).Checked := False;
  end;
end;

procedure TfrmQRCode.CheckS1Click(Sender: TObject);
var
  flag: Integer;
begin
  flag := TCheckBox(Sender).Tag;
  if TCheckBox(Sender).Checked then
  begin
    if FindComponent('Check' + inttostr(flag)) <> nil then
      TCheckBox(FindComponent('Check' + inttostr(flag))).Checked := False;
  end;
end;

procedure TfrmQRCode.btnClearClick(Sender: TObject);
var
  i: Integer;
begin
  for I := 1 to 10 do
  begin
    if FindComponent('edtPath' + inttostr(i)) <> nil then
      TEdit(FindComponent('edtPath' + inttostr(i))).Text := '';
    if FindComponent('edtSPath' + inttostr(i)) <> nil then
      TEdit(FindComponent('edtSPath' + inttostr(i))).Text := '';

    if FindComponent('Check' + inttostr(i)) <> nil then
      TCheckBox(FindComponent('Check' + inttostr(i))).Checked := False;
    if FindComponent('CheckS' + inttostr(i)) <> nil then
      TCheckBox(FindComponent('CheckS' + inttostr(i))).Checked := False;
  end;
end;

function TfrmQRCode.MyCheck2(IndexNo: Integer; FileA, FileB: string; DelimA, DelimB: Boolean): Boolean;
var
  ErrorStr, fileName, fileName1, tableName1, fileName2, tableName2, isql, Delim, fileNameout, s: string;
  zu, yu: Integer;
  Col, Row, sheetIndex, FirstRow: Integer;
  list: TStringList;
begin
 { ErrorStr := '';
  zu := ceil(IndexNo / 2);
  yu := (IndexNo mod 2);
  if yu = 0 then
    yu := 2;
  if (FileA = '') and (FileB = '') then
    Exit;
  if (DelimA = False) and (DelimB = False) then
  begin
    ErrorStr := Format('第%d组第%d个未选择匹配符！！！', [zu, yu]);
    mmoLog2.Lines.Add(ErrorStr);
    Exit;
  end;
  if (FileA = '') and (FileB <> '') then
  begin
    ErrorStr := Format('第%d组第%d个待匹配文件文件不能为空！！！', [zu, yu]);
    mmoLog2.Lines.Add(ErrorStr);
    Exit;
  end;
  if (FileA <> '') and (FileB = '') then
  begin
    ErrorStr := Format('第%d组第%d个源不能为空！！！', [zu, yu]);
    mmoLog2.Lines.Add(ErrorStr);
    Exit;
  end;
  if DelimA then
    Delim := ' '
  else if DelimB then
    Delim := ' ';

  Application.ProcessMessages;
  mmoLog2.Lines.Add(Format('开始匹配第%d组第%d个，请等待...', [zu, yu]));

  fileName1 := FileA;
  tableName1 := '##' + Trim(ChangeFileExt(ExtractFileName(fileName1), ''));
    //从txt文件保存到数据库
  isql := 'drop table [' + tableName1 + ']';
  try
    ADOConnection1.Execute(isql);
  except
  end;
  isql := 'create table [' + tableName1 + '](条码内容 varchar(200),验证码 varchar(100));';
  isql := isql + 'BULK INSERT [' + tableName1 + '] FROM ' + QuotedStr(fileName1) + 'WITH(FIELDTERMINATOR =' + quotedstr(Delim) + ',ROWTERMINATOR =''\n'')';
  try
    ADOConnection1.Execute(isql);
  except
    on e: Exception do
    begin
      mmoLog2.Lines.Add(Format('导入第%d组第%d个文件【%s】失败，异常信息%s', [zu, yu, FileA, e.Message]));
      exit;
    end;
  end;

  list := TStringList.Create;
  try
    fileName2 := FileB;
    tableName2 := '##' + Trim(ChangeFileExt(ExtractFileName(fileName2), ''));
    XLS.Filename := FileB;
    if (ExtractFileExt(XLS.Filename) = '.xls') or (ExtractFileExt(XLS.Filename) = '.xlsx') then
    begin
      XLS.Read;
      for sheetIndex := 0 to 99 do //最多可导入100个Sheet
      begin
        try
          if not XLS.Sheets[sheetIndex].IsEmpty then
          begin
            for Row := 1 to XLS.Sheets[sheetIndex].LastRow do
            begin
              s := '';
              for Col := XLS.Sheets[sheetIndex].FirstCol + 1 to XLS.Sheets[sheetIndex].LastCol - 1 do
              begin
                if s = '' then
                  s := XLS.Sheets[sheetIndex].AsString[Col, Row]
                else
                  s := s + Delim + XLS.Sheets[sheetIndex].AsString[Col, Row];
              end;
              list.Append(Trim(s));
            end;
          end;
        except
          Break;  //屏蔽异常
        end;
      end;
          //fileName := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '(无序号文件).txt';
      fileName := ExtractFilePath(FileB) + tableName2 + '.txt';
      list.SaveToFile(fileName);
    end;
  finally
    FreeAndNil(List);
  end;

  //从txt文件保存到数据库
  if FileExists(fileName) = False then
  begin
    mmoLog2.Lines.Add(Format('导入第%d组第%d个文件【%s】失败', [zu, yu, FileB]));
    exit;
  end;

  isql := 'drop table [' + tableName2 + ']';
  try
    ADOConnection1.Execute(isql);
  except
  end;
  try
    isql := 'create table [' + tableName2 + '](条码内容 varchar(200),验证码 varchar(100));';
    isql := isql + 'BULK INSERT [' + tableName2 + '] FROM ' + QuotedStr(fileName) + 'WITH(FIELDTERMINATOR =' + quotedstr(Delim) + ',ROWTERMINATOR =''\n'')';
    try
      ADOConnection1.Execute(isql);
    except
      on e: Exception do
      begin
        mmoLog2.Lines.Add(Format('导入第%d组第%d个文件【%s】失败，异常信息%s', [zu, yu, FileB, e.Message]));
        exit;
      end;
    end;
  finally
    if FileExists(fileName) then
      DeleteFile(fileName);
  end;

  fileNameout := ExtractFilePath(FileA) + Trim(ChangeFileExt(ExtractFileName(FileA), '')) + '(未匹配记录).txt';
  isql := Format('SELECT A.* FROM [%s] AS A LEFT JOIN [%s] AS B ON A.条码内容=B.条码内容 AND A.验证码=B.验证码 WHERE B.条码内容 IS NULL AND B.验证码  IS NULL', [tableName1, tableName2]);
  isql := 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr(isql) + ',' + QuotedStr(fileNameout) + ';';
  try
    ADOConnection1.Execute(isql);
  except
    on e: Exception do
    begin
      mmoLog2.Lines.Add(Format('对比第%d组第%d个文件失败，异常信息%s', [zu, yu, e.Message]));
      exit;
    end;
  end;
  if FileSizeByName(fileNameout) > 0 then
    mmoLog2.Lines.Add(Format('第%d组第%d个文件存在不匹配的记录，请到【%s】查看', [zu, yu, fileNameout]))
  else
    DeleteFile(fileNameout);
  mmoLog2.Lines.Add(Format('匹配第%d组第%d完成', [zu, yu]));  }
end;

function TfrmQRCode.AutoCheck(FileA, FileB: string; DelimA, DelimB: Boolean): Boolean;
var
  ErrorStr, fileName, fileName1, tableName1, fileName2, tableName2, isql, Delim, fileNameout, s: string;
  zu, yu: Integer;
  Col, Row, sheetIndex, FirstRow: Integer;
  list: TStringList;
begin
  if (FileA = '') and (FileB = '') then
    Exit;

  if DelimA then
    Delim := ' '
  else if DelimB then
    Delim := ',';

  Application.ProcessMessages;
  mmoCustFile.Lines.Add('开始匹配' + FileA + '和' + FileB + '请等待...');

  fileName1 := FileA;
  tableName1 := '##' + Trim(ChangeFileExt(ExtractFileName(fileName1), ''));
    //从txt文件保存到数据库
  isql := 'drop table [' + tableName1 + ']';
  try
    ADOConnection1.Execute(isql);
  except
  end;
  isql := 'create table [' + tableName1 + '](条码内容 varchar(200),验证码 varchar(100));';
  isql := isql + 'BULK INSERT [' + tableName1 + '] FROM ' + QuotedStr(fileName1) + 'WITH(FIELDTERMINATOR =' + quotedstr(Delim) + ',ROWTERMINATOR =''\n'')';
  try
    ADOConnection1.Execute(isql);
  except
    on e: Exception do
    begin
      mmoCustFile.Lines.Add(Format('导入文件【%s】失败，异常信息%s', [FileA, e.Message]));
      exit;
    end;
  end;

  list := TStringList.Create;
  try
    fileName2 := FileB;
    tableName2 := '##' + Trim(ChangeFileExt(ExtractFileName(fileName2), ''));
    XLS.Filename := FileB;
    if (ExtractFileExt(XLS.Filename) = '.xls') or (ExtractFileExt(XLS.Filename) = '.xlsx') then
    begin
      XLS.Read;
      for sheetIndex := 0 to 99 do //最多可导入100个Sheet
      begin
        try
          if not XLS.Sheets[sheetIndex].IsEmpty then
          begin
            for Row := 1 to XLS.Sheets[sheetIndex].LastRow do
            begin
              s := '';
              for Col := XLS.Sheets[sheetIndex].FirstCol + 1 to XLS.Sheets[sheetIndex].LastCol - 1 do
              begin
                if s = '' then
                  s := XLS.Sheets[sheetIndex].AsString[Col, Row]
                else
                  s := s + Delim + XLS.Sheets[sheetIndex].AsString[Col, Row];
              end;
              list.Append(Trim(s));
            end;
          end;
        except
          Break;  //屏蔽异常
        end;
      end;
          //fileName := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '(无序号文件).txt';
      fileName := ExtractFilePath(FileB) + tableName2 + '.txt';
      list.SaveToFile(fileName);
    end;
  finally
    FreeAndNil(List);
  end;

  //从txt文件保存到数据库
  if FileExists(fileName) = False then
  begin
    mmoCustFile.Lines.Add('导入' + FileA + '和' + FileB + '失败');
    exit;
  end;

  isql := 'drop table [' + tableName2 + ']';
  try
    ADOConnection1.Execute(isql);
  except
  end;
  try
    isql := 'create table [' + tableName2 + '](条码内容 varchar(200),验证码 varchar(100));';
    isql := isql + 'BULK INSERT [' + tableName2 + '] FROM ' + QuotedStr(fileName) + 'WITH(FIELDTERMINATOR =' + quotedstr(Delim) + ',ROWTERMINATOR =''\n'')';
    try
      ADOConnection1.Execute(isql);
    except
      on e: Exception do
      begin
        mmoCustFile.Lines.Add(Format('导入文件【%s】失败，异常信息%s', [FileB, e.Message]));
        exit;
      end;
    end;
  finally
    if FileExists(fileName) then
      DeleteFile(fileName);
  end;

  fileNameout := ExtractFilePath(FileA) + Trim(ChangeFileExt(ExtractFileName(FileA), '')) + '(未匹配记录).txt';
  isql := Format('SELECT A.* FROM [%s] AS A LEFT JOIN [%s] AS B ON A.条码内容=B.条码内容 AND A.验证码=B.验证码 WHERE B.条码内容 IS NULL AND B.验证码  IS NULL', [tableName1, tableName2]);
  isql := 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr(isql) + ',' + QuotedStr(fileNameout) + ';';
  try
    ADOConnection1.Execute(isql);
  except
    on e: Exception do
    begin
      mmoCustFile.Lines.Add(Format('对比' + FileA + '和' + FileB + '文件失败，异常信息%s', [e.Message]));
      exit;
    end;
  end;
  if FileSizeByName(fileNameout) > 0 then
    mmoCustFile.Lines.Add(Format(FileA + '和' + FileB + '存在不匹配的记录，请到【%s】查看', [fileNameout]))
  else
    DeleteFile(fileNameout);
  mmoCustFile.Lines.Add('匹配' + FileA + '和' + FileB + '完成');
end;

procedure TfrmQRCode.btnCheck2Click(Sender: TObject);
var
  i, zu, yu: Integer;
  Delim: Boolean;
  sDateTime: TDateTime;
begin
  try
    btnCheck2.Enabled := False;
    mmoLog2.Clear;
    mmoLog2.Lines.Add('可能耗费较多时间，正在匹配中...');
    sDateTime := Now;
    for I := 1 to 10 do
    begin
      zu := ceil(i / 2);
      yu := (i mod 2);
      if yu = 0 then
        yu := 2;
      if ((FindComponent('edtPath' + inttostr(i)) <> nil) and (FindComponent('edtSPath' + inttostr(i)) <> nil) and (FindComponent('Check' + inttostr(i)) <> nil) and (FindComponent('CheckS' + inttostr(i)) <> nil)) then
        MyCheck2(i, TEdit(FindComponent('edtPath' + inttostr(i))).Text, TEdit(FindComponent('edtSPath' + inttostr(i))).Text, TCheckBox(FindComponent('Check' + inttostr(i))).Checked, TCheckBox(FindComponent('CheckS' + inttostr(i))).Checked)
      else
        mmoLog2.Lines.Add(Format('运算到第%d组第%d个发生错误', [zu, yu]));
    end;
    mmoLog2.Lines.Add('恭喜，已全部匹配完成，共耗时：' + IntToStr(SecondsBetween(sDateTime, Now)) + '秒！');
  finally
    btnCheck2.Enabled := True;
  end;
end;

procedure TfrmQRCode.btnDelCustFileClick(Sender: TObject);
begin
  inherited;
  if not qryCustFile.IsEmpty then
    qryCustFile.Delete;
end;

procedure TfrmQRCode.btnDelMultiClick(Sender: TObject);
begin
  inherited;
  if not qryMulti.IsEmpty then
    qryMulti.Delete;
end;

procedure TfrmQRCode.btnDownAClick(Sender: TObject);
var
  i: Integer;
begin
  if FCheckListBoxSelectCountA = 0 then
  begin
    ShowMessage('未选中行时不允许移动！');
    Exit;
  end
  else if FCheckListBoxSelectCountA > 1 then
  begin
    ShowMessage('同时选中多行时不允许移动！');
    Exit;
  end;
  for i := 0 to CheckListBoxA.Items.Count - 1 do
  begin
    if CheckListBoxA.Checked[i] then
    begin
      if i < CheckListBoxA.Items.Count - 1 then
        CheckListBoxA.Items.Move(i, i + 1);
      Break;
    end;
  end;
end;

procedure TfrmQRCode.btnDownBClick(Sender: TObject);
var
  i: Integer;
begin
  if FCheckListBoxSelectCountB = 0 then
  begin
    ShowMessage('未选中行时不允许移动！');
    Exit;
  end
  else if FCheckListBoxSelectCountB > 1 then
  begin
    ShowMessage('同时选中多行时不允许移动！');
    Exit;
  end;
  for i := 0 to CheckListBoxB.Items.Count - 1 do
  begin
    if CheckListBoxB.Checked[i] then
    begin
      if i < CheckListBoxB.Items.Count - 1 then
        CheckListBoxB.Items.Move(i, i + 1);
      Break;
    end;
  end;
end;

procedure TfrmQRCode.btnAddMultiDirClick(Sender: TObject);
var
  strCaption, strDirectory: string;
  wstrRoot: WideString;
  fileList: TStringList;
  i: Integer;
begin
  strCaption := '可直接选择文件夹。';
  //该参数是浏览文件夹窗口的显示说明部分
  wstrRoot := Trim(edtPath.Text);
  //这个参数表示所显示的浏览文件夹窗口中的根目录，默认或空表示“我的电脑”。
  SelectDirectory(strCaption, wstrRoot, strDirectory);
  lbl7.Caption := strDirectory;
  //传递结果，其中参数strDirectory表示函数的返回值
  fileList := TStringList.Create;
  try
    Searchfile(PChar(strDirectory), '.txt', fileList);
    for i := 0 to fileList.Count - 1 do
    begin
      if UpperCase(ExtractFileExt(fileList[i])) = UpperCase('.txt') then
      begin
        with qryMulti do
        begin
          if not (State in [dsEdit, dsInsert]) then
            Append;
          FieldByName('selected').AsBoolean := True;
          FieldByName('fileName').AsString := fileList[i];
          FieldByName('isContainFirst').AsBoolean := True;
          if not (State in [dsBrowse]) then
            Post;
        end;
      end;
    end;
    lblFileCount.Caption := '文件数量：' + Inttostr(fileList.Count);
  finally
    fileList.Free;
  end;
end;

procedure TfrmQRCode.btnAddCustFileClick(Sender: TObject);
var
  iQuery: TADOQuery;
  OpenDialog: TOpenDialog;
  i: Integer;
begin
  OpenDialog := TOpenDialog.Create(Self);
  OpenDialog.FileName := '客户文件';
  OpenDialog.Filter := '*.txt;*.xls;*.xlsx|*.txt;*.xls;*.xlsx';
  OpenDialog.Options := [ofHideReadOnly, ofEnableSizing, ofAllowMultiSelect];
  iQuery := TADOQuery.Create(Self);
  try
    if OpenDialog.Execute then
    begin
      for i := 0 to OpenDialog.Files.Count - 1 do
      begin
        with qryCustFile do
        begin
          //如果存在相同的文件
          DisableControls;
          try
            First;
            while not Eof do
            begin
              if Trim(FieldByName('fileName').AsString) = Trim(OpenDialog.Files[i]) then//if Trim(ChangeFileExt(ExtractFileName(FieldByName('fileName').AsString), '')) = Trim(ChangeFileExt(ExtractFileName(OpenDialog.FileName), '')) then
              begin
                ShowMessage('已经存在文件【' + Trim(FieldByName('fileName').AsString) + '】!!!'); //ShowMessage('已经存在文件【' + Trim(ChangeFileExt(ExtractFileName(FieldByName('fileName').AsString), '')) + '】!!!');
                Exit;
              end;
              Next;
            end;
          finally
            EnableControls;
          end;

          if not (State in [dsEdit, dsInsert]) then
            Append;
          FieldByName('selected').AsBoolean := True;
          FieldByName('fileName').AsString := OpenDialog.files[i];
          FieldByName('isContainFirst').AsBoolean := True;
          if not (State in [dsBrowse]) then
            Post;
        end;
      end;
    end;
  finally
    iQuery.Free;
    OpenDialog.Free;
  end;
end;

procedure TfrmQRCode.btnImportAClick(Sender: TObject);
var
  iQuery: TADOQuery;
begin
  OpenDialogA.FileName := '文本文件';
  OpenDialogA.Filter := '*.txt|*.txt';
  iQuery := TADOQuery.Create(Self);
  try
    if OpenDialogA.Execute then
    begin
      edtA.Text := OpenDialogA.FileName;
      with iQuery do
      begin
        Connection := ADOConnection1;
        CommandTimeout := 0;
        Close;
        SQL.Clear;
        SQL.Add('if not Exists(select 1 from sysobjects where name=''tableA'' and xtype=''U'') create table tableA(code varchar(100)) else TRUNCATE TABLE tableA;');
        SQL.Add('BULK INSERT tableA FROM ' + QuotedStr(Trim(edtA.Text)));
        ExecSQL;
        Close;
        SQL.Text := 'select count(1) as count from tableA';
        Open;
        mmoLog.Lines.Add('A文件总记录[' + FieldByName('count').AsString + ']行');
      end;
    end;
  finally
    iQuery.Free;
  end;
end;

procedure TfrmQRCode.btnImportBClick(Sender: TObject);
var
  iQuery: TADOQuery;
begin
  OpenDialogB.FileName := '文本文件';
  OpenDialogB.Filter := '*.txt|*.txt';
  iQuery := TADOQuery.Create(Self);
  try
    if OpenDialogB.Execute then
    begin
      edtB.Text := OpenDialogB.FileName;
      with iQuery do
      begin
        Connection := ADOConnection1;
        CommandTimeout := 0;
        Close;
        SQL.Clear;
        SQL.Add('if not Exists(select 1 from sysobjects where name=''tableB'' and xtype=''U'') create table tableB(code varchar(100)) else TRUNCATE TABLE tableB;');
        SQL.Add('BULK INSERT tableB FROM ' + QuotedStr(Trim(edtB.Text)));
        ExecSQL;
        Close;
        SQL.Text := 'select count(1) as count from tableB';
        Open;
        mmoLog.Lines.Add('B文件总记录[' + FieldByName('count').AsString + ']行');
      end;
    end;
  finally
    iQuery.Free;
  end;
end;

procedure TfrmQRCode.btnMatchClick(Sender: TObject);
var
  isql, iTable: string;
  ssDate, sDate, eDate: TDateTime;
  iQuery, iQuery2, iQuery3: TADOQuery;
begin
  if Trim(edtNo.Text) = '' then
  begin
    ShowMessage('主表单据号不能为空!!!');
    Exit;
  end;
  sDate := Now;
  ssDate := Now;
  Memo1.Lines.Clear;
  //检测本主表单据号下的多版本数据表的重复数据
  iQuery := TADOQuery.Create(Self);
  iQuery2 := TADOQuery.Create(Self);
  iQuery3 := TADOQuery.Create(Self);
  try
    with iQuery do
    begin
      CommandTimeout := 0;
      Connection := ADOConnection1;
      Close;
      SQL.Text := 'SELECT name FROM sysobjects WHERE NAME LIKE ' + QuotedStr(Trim(edtNo.Text) + '%') + ' AND xtype=''U''';
      Open;

      isql := '';
      //遍历本主表单据号下的多版本数据表
      first;
      while not Eof do
      begin
        iTable := FieldByName('name').AsString;
        isql := 'select sjm from ' + iTable + ' where ppdm+yscdm+Ny+sjm in (select ppdm+yscdm+Ny+sjm from ' + Trim(iTable) + ' group by ppdm,yscdm,Ny,sjm having count(*) > 1) and id not in (select min(id) from ' + Trim(iTable) + ' group by ppdm,yscdm,Ny,sjm having count(ppdm+yscdm+Ny+sjm)>1)';

        //检测本包重复随机码
        with iQuery2 do
        begin
          CommandTimeout := 0;
          Connection := ADOConnection1;
          Close;
          SQL.Text := isql;
          Open;
          if not IsEmpty then
          begin
            Memo1.Lines.Add('本包【' + iTable + '】检测到重复随机码：');
            First;
            while not Eof do
            begin
              Memo1.Lines.Add(FieldByName('sjm').AsString);
              Next;
            end;
            isql := 'delete from ' + Trim(iTable) + ' where ppdm+yscdm+Ny+sjm in (select ppdm+yscdm+Ny+sjm from ' + Trim(iTable) + ' group by ppdm,yscdm,Ny,sjm having count(*) > 1) and id not in (select min(id) from ' + Trim(iTable) + ' group by ppdm,yscdm,Ny,sjm having count(ppdm+yscdm+Ny+sjm)>1)';
            try
              ADOConnection1.Execute(isql);
            except
              on e: Exception do
              begin
                Memo1.Lines.Add('删除本包重复数据失败，异常信息：' + e.Message);
                Exit;
              end;
            end;
            eDate := Now;
            Memo1.Lines.Add('删除本包重复数据成功，耗时：' + IntToStr(SecondsBetween(eDate, sDate)) + '秒!!!');
          end
          else
            Memo1.Lines.Add('本包【' + iTable + '】未检测到重复随机码');
        end;
        iQuery.Next;
      end;
    end;
    //检测历史库重复数据
    isql := 'SELECT name FROM sysobjects WHERE NAME LIKE''QR%'' AND xtype=''U'' AND NAME<>' + QuotedStr(Trim(iTable)) + ' and convert(char(10),crDate,120)=convert(char(10),getdate(),120)';
    with ADOQuery1 do
    begin
      Close;
      SQL.Text := isql;
      Open;
      if not IsEmpty then
      begin
        First;
        while not Eof do
        begin
          sDate := Now;
          Memo1.Lines.Add('检测历史库【' + ADOQuery1.fieldbyname('name').AsString + '】...');
          isql := 'SELECT * from ' + trim(iTable) + ' WHERE EXISTS(SELECT 1 from ' + ADOQuery1.fieldByName('name').AsString + ' WHERE ' + ADOQuery1.fieldByName('name').AsString + '.ppdm+' + ADOQuery1.fieldByName('name').AsString + '.yscdm+' + ADOQuery1.fieldByName('name').AsString + '.NY+' + ADOQuery1.fieldByName('name').AsString + '.sjm=' + trim(iTable) + '.ppdm+' + trim(iTable) + '.yscdm+' + trim(iTable) + '.NY+' + trim(iTable) + '.sjm)';
          with iQuery3 do
          begin
            CommandTimeout := 0;
            Connection := ADOConnection1;
            Close;
            SQL.Text := isql;
            Open;
            if not IsEmpty then
            begin
              Memo1.Lines.Add('检测到历史库【' + ADOQuery1.fieldbyname('name').AsString + '】有重复数据：');
              First;
              while not Eof do
              begin
                Memo1.Lines.Add('品牌代码【' + FieldByName('ppdm').AsString + '】，印刷厂代码【' + FieldByName('yscdm').AsString + '】，年月【' + FieldByName('ny').AsString + '】，随机码【' + FieldByName('sjm').AsString + '】');
                Next;
              end;
              Memo1.Lines.Add('删除历史库【' + ADOQuery1.fieldbyname('name').AsString + '】中重复数据......');
              isql := 'delete from ' + trim(iTable) + ' WHERE EXISTS(SELECT 1 from ' + ADOQuery1.fieldByName('name').AsString + ' WHERE ' + ADOQuery1.fieldByName('name').AsString + '.ppdm+' + ADOQuery1.fieldByName('name').AsString + '.yscdm+' + ADOQuery1.fieldByName('name').AsString + '.NY+' + ADOQuery1.fieldByName('name').AsString + '.sjm=' + trim(iTable) + '.ppdm+' + trim(iTable) + '.yscdm+' + trim(iTable) + '.NY+' + trim(iTable) + '.sjm)';
              try
                ADOConnection1.Execute(isql);
              except
                on e: Exception do
                begin
                  Memo1.Lines.Add('删除历史库【' + ADOQuery1.fieldbyname('name').AsString + '】中重复数据失败，异常信息：' + e.Message);
                  Exit;
                end;
              end;
              eDate := Now;
              Memo1.Lines.Add('删除历史库【' + ADOQuery1.fieldbyname('name').AsString + '】中重复数据成功，耗时：' + IntToStr(SecondsBetween(eDate, sDate)) + '秒!!!');
            end
            else
              Memo1.Lines.Add('检测历史库【' + ADOQuery1.fieldbyname('name').AsString + '】无重复数据!');
          end;
          Next;
        end;
      end
      else
        Memo1.Lines.Add('没有检测到历史库');
    end;
  finally
    iQuery.Free;
    iQuery2.Free;
    iQuery3.Free;
  end;
  eDate := Now;
  Memo1.Lines.Add('检测完成，耗时：' + IntToStr(SecondsBetween(eDate, ssDate)) + '秒!!!');
end;

procedure TfrmQRCode.btnNotRepeatMultiClick(Sender: TObject);
var
  fileName,aa: string;
  sDate, eDate: TDateTime;
begin
  if not FIsImported then
  begin
    mmoMulti.Lines.Add('没有要处理的数据，请先导入后再处理!!!');
    Exit;
  end;
  if Trim(lbl7.Caption) = '' then
    fileName := Trim(edtPath.Text) + '\' + FormatDateTime('yyyymmddhh', Now) + '【不重复】.txt'
  else
    fileName := Trim(lbl7.Caption) + '\' + FormatDateTime('yyyymmddhh', Now) + '【不重复】.txt';
  try
    sDate := now;
    aa:='exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('select DISTINCT code from ##tb_Multi') + ',' + QuotedStr(fileName) + ';';
    //ShowMessage(aa);
    ADOConnection1.Execute('exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('select DISTINCT code from ##tb_Multi') + ',' + QuotedStr(fileName) + ';');
  except
    on e: Exception do
    begin
      mmoMulti.Lines.Add('处理不重复失败，异常信息：' + E.Message);
      Exit;
    end;
  end;
  eDate := Now;
  mmoMulti.Lines.Add('处理不重复成功，生成文件【' + fileName + '】,耗时：' + IntToStr(SecondsBetween(eDate, sDate)) + '秒!!!');
end;

procedure TfrmQRCode.btnPublicClick(Sender: TObject);
var
  datePointer: TDateTime;
  iQuery: TADOQuery;
  filename: string;
begin
  if Trim(edtA.Text) = '' then
    MessageBox(Self.Handle, 'A文件不能为空，请检查！', '提示', MB_OK)
  else if Trim(edtB.Text) = '' then
    MessageBox(Self.Handle, 'B文件不能为空，请检查！', '提示', MB_OK)
  else if ExtractFileExt(edtA.Text) <> '.txt' then
    MessageBox(Self.Handle, '打开的A文件类型不对，请检查！', '提示', MB_OK)
  else if ExtractFileExt(edtB.Text) <> '.txt' then
    MessageBox(Self.Handle, '打开的B文件类型不对，请检查！', '提示', MB_OK)
  else if not FileExists(edtA.Text) then
    MessageBox(Self.Handle, 'A文件不存在，请检查！', '提示', MB_OK)
  else if not FileExists(edtB.Text) then
    MessageBox(Self.Handle, '父文件不存在，请检查！', '提示', MB_OK)
  else
  begin
    btnPublic.Enabled := not btnPublic.Enabled;
    mmoLog.Lines.Add('可能耗费较多时间，正在匹配中...');
    datePointer := Now;

    iQuery := TADOQuery.Create(Self);
    try
      iQuery.CommandTimeout := 0;
      try
        with iQuery do
        begin
          CommandTimeout := 0;
          Connection := ADOConnection1;
          Close;
          SQL.Text := 'if not exists(select 1 from sysobjects where name=''tableC'' and xtype=''U'') create table tableC(code varchar(100)) else TRUNCATE TABLE tableC;' + 'INSERT INTO tableC(code) SELECT code FROM tableA WHERE exists(SELECT code FROM tableB where tableB.code=tableA.code) order by code;';
          ExecSQL;
        end;
      except
        mmoLog.Lines.Add('匹配失败，请稍后重试！');
        Exit;
      end;
      with iQuery do
      begin
        CommandTimeout := 0;
        Connection := ADOConnection1;
        Close;
        SQL.Text := 'select count(1) as count from tableC';
        Open;
        mmoLog.Lines.Add('匹配完成！共匹配：' + FieldByName('count').AsString + '条记录，共耗时：' + IntToStr(SecondsBetween(Now, datePointer)) + '秒！');
        Close;
        filename := Trim(edtPath.Text) + '\' + Trim(btnPublic.Caption) + '.txt';
        sql.Text := 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('SELECT * from DB_QRCode..tableC') + ',' + QuotedStr(filename);
        try
          ExecSQL;
        except
          on e: Exception do
          begin
            mmoLog.Lines.Add('文件生成失败，异常信息：' + E.Message);
            Exit;
          end;
        end;
        mmoLog.Lines.Add('文件生成成功，文件名：' + filename);
      end;
    finally
      iQuery.Free;
    end;
    btnPublic.Enabled := not btnPublic.Enabled;
  end;
end;

procedure TfrmQRCode.btnRepeatClick(Sender: TObject);
var
  datePointer: TDateTime;
  iQuery: TADOQuery;
  filename: string;
begin
  if Trim(edtA.Text) = '' then
    MessageBox(Self.Handle, 'A文件不能为空，请检查！', '提示', MB_OK)
  else if Trim(edtB.Text) = '' then
    MessageBox(Self.Handle, 'B文件不能为空，请检查！', '提示', MB_OK)
  else if ExtractFileExt(edtA.Text) <> '.txt' then
    MessageBox(Self.Handle, '打开的A文件类型不对，请检查！', '提示', MB_OK)
  else if ExtractFileExt(edtB.Text) <> '.txt' then
    MessageBox(Self.Handle, '打开的B文件类型不对，请检查！', '提示', MB_OK)
  else if not FileExists(edtA.Text) then
    MessageBox(Self.Handle, 'A文件不存在，请检查！', '提示', MB_OK)
  else if not FileExists(edtB.Text) then
    MessageBox(Self.Handle, '父文件不存在，请检查！', '提示', MB_OK)
  else
  begin
    btnRepeat.Enabled := not btnRepeat.Enabled;
    mmoLog.Lines.Add('可能耗费较多时间，正在匹配中...');
    datePointer := Now;

    iQuery := TADOQuery.Create(Self);
    try
      iQuery.CommandTimeout := 0;
      try
        with iQuery do
        begin
          Connection := ADOConnection1;
          Close;
          SQL.Text := 'if not exists(select 1 from sysobjects where name=''tableC'' and xtype=''U'') create table tableC(code varchar(100)) else TRUNCATE TABLE tableC;' + 'INSERT INTO tableC(code) SELECT code FROM tableA WHERE not exists(SELECT code FROM tableB where tableB.code=tableA.code) order by code;';
          ExecSQL;
        end;
      except
        mmoLog.Lines.Add('匹配失败，请稍后重试！');
        Exit;
      end;
      with iQuery do
      begin
        Connection := ADOConnection1;
        Close;
        SQL.Text := 'select count(1) as count from tableC';
        Open;
        mmoLog.Lines.Add('匹配完成！共匹配：' + FieldByName('count').AsString + '条记录，共耗时：' + IntToStr(SecondsBetween(Now, datePointer)) + '秒！');
        Close;
        filename := Trim(edtPath.Text) + '\' + Trim(btnRepeat.Caption) + '.txt';
        sql.Text := 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('SELECT * from DB_QRCode..tableC') + ',' + QuotedStr(filename);
        try
          ExecSQL;
        except
          on e: Exception do
          begin
            mmoLog.Lines.Add('文件生成失败，异常信息：' + E.Message);
            Exit;
          end;
        end;
        mmoLog.Lines.Add('文件生成成功，文件名：' + filename);
      end;
    finally
      iQuery.Free;
    end;
    btnRepeat.Enabled := not btnRepeat.Enabled;
  end;
end;

procedure TfrmQRCode.btnRepeatMultiClick(Sender: TObject);
var
  fileName: string;
  sDate, eDate: TDateTime;
begin
  if not FIsImported then
  begin
    mmoMulti.Lines.Add('没有要处理的数据，请先导入后再处理!!!');
    Exit;
  end;
  if Trim(lbl7.Caption) = '' then
    fileName := Trim(edtPath.Text) + '\' + FormatDateTime('yyyymmddhh', Now) + '【重复】.txt'
  else
    fileName := Trim(lbl7.Caption) + '\' + FormatDateTime('yyyymmddhh', Now) + '【重复】.txt';
  try
    sDate := Now;
    ADOConnection1.Execute('exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('select code from ##tb_Multi GROUP BY code HAVING COUNT(code)>1') + ',' + QuotedStr(fileName) + ';');
  except
    on e: Exception do
    begin
      mmoMulti.Lines.Add('处理重复失败，异常信息：' + E.Message);
      Exit;
    end;
  end;
  eDate := Now;
  mmoMulti.Lines.Add('处理重复成功，生成文件【' + fileName + '】,耗时：' + IntToStr(SecondsBetween(eDate, sDate)) + '秒!!!');
end;

procedure TfrmQRCode.btnStartClick(Sender: TObject);
var
  i, threadCount: Integer;
  myThread: TThread;
  Handles: TWOHandleArray;
  QRNo, QRType, QRMode, QRWebAddr, QRBrandCode, QRPrinterCode, QRNy, QRCount, QRThreadCount, QRMax: string;
  err: string;
begin
  lblResult.Caption := '';
  if Trim(edtNo.Text) = '' then
  begin
    ShowMessage('主表单据号不能为空!!!');
    Exit;
  end;
  if Trim(Copy(edtNo.Text, 1, 2)) <> 'QR' then
  begin
    ShowMessage('主表单据号必须以QR开头!!!');
    Exit;
  end;
  if Trim(edtType.Text) = '' then
  begin
    ShowMessage('条盒/小盒不能为空!!!');
    Exit;
  end;
  if Trim(edtMode.Text) = '' then
  begin
    ShowMessage('厂家类型不能为空!!!');
    Exit;
  end;
  if Trim(edtWebAddr.Text) = '' then
  begin
    ShowMessage('网址不能为空!!!');
    Exit;
  end;
  if Trim(edtBrandCode.Text) = '' then
  begin
    ShowMessage('品牌代码不能为空!!!');
    Exit;
  end;
  if Trim(edtPrinterCode.Text) = '' then
  begin
    ShowMessage('印刷厂代码不能为空!!!');
    Exit;
  end;
  if Trim(edtNy.Text) = '' then
  begin
    ShowMessage('年月不能为空!!!');
    Exit;
  end;
  if Length(edtNy.Text) <> 4 then
  begin
    ShowMessage('年月位数不对，并且格式必须是YYMM!!!');
    Exit;
  end;
  if (Trim(edtCount.Text) = '') or (Trim(edtCount.Text) = '0') then
  begin
    ShowMessage('二维码数量不能为空或0!!!');
    Exit;
  end;
  if (Trim(edtThreadCount.Text) = '') or (Trim(edtThreadCount.Text) = '0') then
  begin
    ShowMessage('工作线程数不能为空或0!!!');
    Exit;
  end;
  if (Trim(edtMax.Text) = '') or (Trim(edtMax.Text) = '0') then
  begin
    ShowMessage('单个数据包上限不能为空或0!!!');
    Exit;
  end;

  if not ShowAsk(0, PChar('请确定您输入的信息都正确，确定要生成二维码吗？')) then
    Exit;

  with ADOQuery1 do
  begin
    Close;
    SQL.Text := 'select 1 from sysobjects where name like' + QuotedStr(Trim(edtNo.Text) + '%') + ' and xtype=''U''';
    Open;
    if not IsEmpty then
    begin
      if not ShowAsk(0, PChar('此单的二维码已经生成过，您确定还要继续生成吗？')) then
        Exit;
    end;
  end;

  lblResult.Caption := '生成中，请稍后......';
  //加锁
//  EnterCriticalSection(Cs);
  //到UI取值
  threadCount := StrToInt(frmQRCode.edtThreadCount.Text);
  QRNo := Trim(frmQRCode.edtNo.Text);
  QRType := Trim(frmQRCode.edtType.Text);
  QRMode := Trim(frmQRCode.edtMode.Text);
  QRWebAddr := Trim(frmQRCode.edtWebAddr.Text);
  QRBrandCode := Trim(frmQRCode.edtBrandCode.Text);
  QRPrinterCode := Trim(frmQRCode.edtPrinterCode.Text);
  QRNy := Trim(frmQRCode.edtNy.Text);
  QRCount := Trim(frmQRCode.edtCount.Text);
  QRThreadCount := Trim(frmQRCode.edtThreadCount.Text);
  QRMax := Trim(frmQRCode.edtMax.Text);
  //释放锁
//  LeaveCriticalSection(Cs);

  for i := 0 to StrToInt(QRThreadCount) - 1 do
  begin
    myThread := TThread.CreateAnonymousThread(
      procedure
      var
        adoConnection: TADOConnection;
      begin
        CoInitialize(nil);
        try
          //创建新的ADOConnection，给线程设置参数，并执行
          adoConnection := TADOConnection.Create(nil);
          adoConnection.Connected := False;
          adoConnection.CommandTimeout := 0;
          adoConnection.ConnectionString := 'Provider=SQLOLEDB.1;Password=qwe@123;Persist Security Info=True;User ID=sa;Initial Catalog=DB_QRCode;Data Source=127.0.0.1';
          adoConnection.LoginPrompt := False;
          adoConnection.Connected := True;
          try
            adoConnection.Execute('exec pWriteQRCode ' + Quotedstr(QRNo) + ',' + Quotedstr(QRMode) + ',' + Quotedstr(QRWebAddr) + ',' + Quotedstr(QRBrandCode) + ',' + Quotedstr(QRPrinterCode) + ',' + Quotedstr(QRNy) + ',' + IntToStr(StrToInt(QRCount) div StrToInt(QRThreadCount)) + ',' + Quotedstr(QRType) + ',' + QRMax);
          except
            on E: Exception do
              err := e.Message;
          end;
        finally
          CoUninitialize;
          adoConnection.Free;
        end;
      end);
    Handles[i] := myThread.Handle;
    myThread.Start;
  end;

  try
    TThread.CreateAnonymousThread(
      procedure
      var
        sDate, eDate: TDateTime;
      begin
        sDate := Now;
        WaitForMultipleObjects(threadCount, @Handles, TRUE, INFINITE);

        TThread.Synchronize(nil,
          procedure
          begin
            eDate := Now;
            if Trim(err) = '' then
              frmQRCode.lblResult.Caption := '二维码生成成功，耗时：' + IntToStr(SecondsBetween(eDate, sDate)) + '秒。为了避免二维码和历史库中的重复，请继续匹配历史库二维码!!!'
            else
              frmQRCode.lblResult.Caption := '二维码生成失败，异常信息：' + err;
            frmQRCode.lblResult.Refresh;
          end);
      end).Start;
  except
    on E: Exception do
    begin
      lblResult.Caption := '二维码生成失败，异常信息:' + e.Message;
      Exit;
    end;
  end;
end;

procedure TfrmQRCode.btnTxtClick(Sender: TObject);
var
  path: string;
begin
  OpenDialogA.FileName := '文本文件';
  OpenDialogA.Filter := '*.txt|*.txt';
  if OpenDialogA.Execute then
  begin
    path := Trim(OpenDialogA.Files.Text);
    //path:=StringReplace(path,#13#10,'',[rfReplaceAll,rfIgnoreCase]);//文件名中有回车符号，后面sql语句报错
    edtTxt.Text := path;
  end;
end;

procedure TfrmQRCode.btnUpAClick(Sender: TObject);
var
  i: Integer;
begin
  if FCheckListBoxSelectCountA = 0 then
  begin
    ShowMessage('未选中行时不允许移动！');
    Exit;
  end
  else if FCheckListBoxSelectCountA > 1 then
  begin
    ShowMessage('同时选中多行时不允许移动！');
    Exit;
  end;

  for i := 0 to CheckListBoxA.Items.Count - 1 do
  begin
    if CheckListBoxA.Checked[i] then
    begin
      if i > 0 then
        CheckListBoxA.Items.Move(i, i - 1);
      Break;
    end;
  end;
end;

procedure TfrmQRCode.btnUpBClick(Sender: TObject);
var
  i: Integer;
begin
  if FCheckListBoxSelectCountB = 0 then
  begin
    ShowMessage('未选中行时不允许移动！');
    Exit;
  end
  else if FCheckListBoxSelectCountB > 1 then
  begin
    ShowMessage('同时选中多行时不允许移动！');
    Exit;
  end;

  for i := 0 to CheckListBoxB.Items.Count - 1 do
  begin
    if CheckListBoxB.Checked[i] then
    begin
      if i > 0 then
        CheckListBoxB.Items.Move(i, i - 1);
      Break;
    end;
  end;
end;

function TfrmQRCode.DbGridEhToExcel(fileNameout: string): string;
var
  ExpClass: TDBGridEhExportclass;
  ADgEh: TDBGridEh;
  MyDataSource: TDataSource;
begin
  Result := '';
  MyDataSource := TDataSource.Create(nil);
  ADgEh := TDBGridEh.Create(nil);
  try
    try
      MyDataSource.DataSet := adoCheck3;
      ADgEh.DataSource := MyDataSource;
      if ADgEh.DataSource.DataSet.IsEmpty then
        exit;
      ExpClass := TDBGridEhExportAsXLS;
      SaveDBGridEhToExportFile(ExpClass, ADgEh, fileNameout, true);
    except
      on e: exception do
      begin
        Result := Format('导出到【%s】失败，失败原因：%s', [fileNameout, e.Message]);
        Exit;
      end;
    end;

  finally
    FreeAndNil(ADgEh);
    FreeAndNil(MyDataSource);
  end;
end;

procedure TfrmQRCode.MyWndProc(var Msg: TMessage);
begin
  if Msg.Msg = WM_MYUSER then
  begin
    if Msg.LParam = 1 then
    begin
      StatusBar1.Panels[0].Text := string(Msg.WParam);
    end
    else if Msg.LParam = 2 then
    begin
      ProgressBar1.Position := Msg.WParam;
    end
    else if Msg.LParam = 3 then
    begin
      StatusBar1.Panels[0].Text := string(Msg.WParam);
    end;
  end;
end;

function TfrmQRCode.DataToExcel(fileNamein, fileNameout: string): string;
var
  txtpath: string;
  sheetIndex, i, y, MaxY: Integer;
  MsgStr, A1, A2: string;
  AList: TStringList;
begin
  {
  AList := TStringList.Create;
  try
    AList.LoadFromFile(fileNamein);
    for I := 0 to AList.Count - 1 do
    begin
      MsgStr := Trim(AList[i]);
      A1 := Copy(MsgStr, 1, Pos(',', MsgStr) - 1);
      A2 := Copy(MsgStr, Pos(',', MsgStr) + 1, Length(MsgStr));
      XLS2.Sheets[0].AsString[0, i+1] := IntToStr(i);
      XLS2.Sheets[0].AsString[1, i+1] := A1;
      XLS2.Sheets[0].AsString[2, i+1] := A2;
    end;
    for i := 0 to 2 do
      XLS2.Sheets[0].AutoWidthCol(i + 1);
     XLS2.SaveToFile(fileNameout);
  finally
    FreeAndNil(AList);
  end;
  }

  try
    if adoCheck3.IsEmpty then
      Exit;
    XLS2.Clear();
    Application.ProcessMessages;
    adoCheck3.DisableControls;
    XLS2.Sheets[0].AsString[0, 0] := '序号';
    for i := 0 to adoCheck3.FieldCount - 1 do
    begin
      XLS2.Sheets[0].AsString[i + 1, 0] := adoCheck3.Fields[i].DisplayName;
    end;
    MaxY := adoCheck3.RecordCount;
    MsgStr := Format('正在转换文件：【%s】', [FileNameout]);
    //mmoCustFile.Lines.Add(MsgStr);

    y := 1;
    adoCheck3.First;
    for I := 1 to MaxY do
    begin
      XLS2.Sheets[0].AsString[0, i] := IntToStr(i);
      XLS2.Sheets[0].AsString[1, i] := adoCheck3.FieldByName('条码内容').AsString;
      XLS2.Sheets[0].AsString[2, i] := adoCheck3.FieldByName('验证码').AsString;

      adoCheck3.Next;
    end;
    for i := 0 to adoCheck3.FieldCount - 1 do
    begin
      XLS2.Sheets[0].AutoWidthCol(i + 1);
    end;
    XLS2.SaveToFile(fileNameout);
    MsgStr := Format('转换文件：【%s】完成', [FileNameout]);
    //mmoCustFile.Lines.Add(MsgStr);
  except
    on E: Exception do
    begin
      MsgStr := Format('导出到【%s】失败，失败原因：%s', [fileNameout, e.Message]);
      mmoCustFile.Lines.Add(MsgStr);
      adoCheck3.EnableControls;
      Exit;
    end;
  end;
end;

function TfrmQRCode.MyCheck3(FileA, FileB, Delim: string): Boolean;
var
  ErrorStr, fileName, fileName1, tableName1, fileName2, tableName2, isql, fileNamein, fileNameout, s, temp: string;
  zu, yu: Integer;
  Col, Row, sheetIndex, FirstRow: Integer;
  list: TStringList;
begin
  if (FileA = '') and (FileB = '') then
    Exit;

  Application.ProcessMessages;
  mmoCustFile.Lines.Add('开始匹配' + FileA + '和' + FileB + '请等待...');

  fileName1 := FileA;
  tableName1 := '##' + Trim(ChangeFileExt(ExtractFileName(fileName1), ''));
    //从txt文件保存到数据库
  isql := 'drop table [' + tableName1 + ']';
  try
    ADOConnection1.Execute(isql);
  except
  end;
  isql := 'create table [' + tableName1 + '](条码内容 varchar(200),验证码 varchar(100));';
  isql := isql + 'BULK INSERT [' + tableName1 + '] FROM ' + QuotedStr(fileName1) + 'WITH(FIELDTERMINATOR =' + quotedstr(Delim) + ',ROWTERMINATOR =''\n'')';
  try
    ADOConnection1.Execute(isql);
  except
    on e: Exception do
    begin
      mmoCustFile.Lines.Add(Format('导入文件【%s】失败，异常信息%s', [FileA, e.Message]));
      exit;
    end;
  end;

  list := TStringList.Create;
  try
    fileName2 := FileB;
    tableName2 := '##' + Trim(ChangeFileExt(ExtractFileName(fileName2), ''));
    XLS.Filename := FileB;
    if (ExtractFileExt(XLS.Filename) = '.xls') or (ExtractFileExt(XLS.Filename) = '.xlsx') then
    begin
      XLS.Read;

      for sheetIndex := 0 to 99 do //最多可导入100个Sheet
      begin
        try
          if not XLS.Sheets[sheetIndex].IsEmpty then
          begin
            for Row := 1 to XLS.Sheets[sheetIndex].LastRow do
            begin
              s := '';
              for Col := XLS.Sheets[sheetIndex].FirstCol + 1 to XLS.Sheets[sheetIndex].LastCol - 1 do
              begin
                if s = '' then
                  s := XLS.Sheets[sheetIndex].AsString[Col, Row]
                else
                  s := s + Delim + XLS.Sheets[sheetIndex].AsString[Col, Row];
              end;
              list.Append(Trim(s));
            end;
          end;
        except
          Break;  //屏蔽异常
        end;
      end;
          //fileName := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '(无序号文件).txt';
      fileName := ExtractFilePath(FileB) + tableName2 + '.txt';
      list.SaveToFile(fileName);
    end;
  finally
    FreeAndNil(List);
  end;

  //从txt文件保存到数据库
  if FileExists(fileName) = False then
  begin
    mmoCustFile.Lines.Add('导入' + FileA + '和' + FileB + '失败');
    exit;
  end;

  isql := 'drop table [' + tableName2 + ']';
  try
    ADOConnection1.Execute(isql);
  except
  end;
  try
    isql := 'create table [' + tableName2 + '](条码内容 varchar(200),验证码 varchar(100));';
    isql := isql + 'BULK INSERT [' + tableName2 + '] FROM ' + QuotedStr(fileName) + 'WITH(FIELDTERMINATOR =' + quotedstr(Delim) + ',ROWTERMINATOR =''\n'')';
    try
      ADOConnection1.Execute(isql);
    except
      on e: Exception do
      begin
        mmoCustFile.Lines.Add(Format('导入文件【%s】失败，异常信息%s', [FileB, e.Message]));
        exit;
      end;
    end;
  finally
    if FileExists(fileName) then
      DeleteFile(fileName);
  end;

  isql := Format('UPDATE [%s] SET 验证码=B.验证码 FROM [%s] A   INNER JOIN [%s] AS B ON A.条码内容=B.条码内容 WHERE ISNULL(A.验证码,'''')='''' ', [tableName1, tableName1, tableName2]);
  try
    ADOConnection1.Execute(isql);
  except
    on e: Exception do
    begin
      mmoCustFile.Lines.Add(Format(FileA + '和' + FileB + '更新验证码失败，异常信息%s', [e.Message]));
      exit;
    end;
  end;

  fileNamein := ExtractFilePath(FileB) + Trim(ChangeFileExt(ExtractFileName(FileB), '')) + '(未匹配记录).txt';
  fileNameout := ExtractFilePath(FileB) + Trim(ChangeFileExt(ExtractFileName(FileB), '')) + '(未匹配记录).xlsx';
  isql := Format('SELECT * FROM [%s] AS A LEFT JOIN [%s] AS B ON A.条码内容=B.条码内容 AND A.验证码=B.验证码 WHERE B.条码内容 IS NULL AND B.验证码  IS NULL', [tableName2, tableName1]);
  //isql := 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr(isql) + ',' + QuotedStr(fileNamein) + ';';
  try
    adoCheck3.SQL.Text := isql;
    adoCheck3.Open;
    DataToExcel(fileNamein, fileNameout);
    //DbGridEhToExcel(fileNameout);

  except
    on e: Exception do
    begin
      mmoCustFile.Lines.Add(Format('对比' + FileA + '和' + FileB + '文件失败，异常信息%s', [e.Message]));
      if FileExists(fileNameout) then
        DeleteFile(fileNameout);
      exit;
    end;
  end;
  if FileSizeByName(fileNameout) > 0 then
  begin
    mmoCustFile.Lines.Add(Format(FileA + '和' + FileB + '存在不匹配的记录，请到【%s】查看', [fileNameout]));
  end
  else
    DeleteFile(fileNameout);

  fileNamein := ExtractFilePath(FileB) + Trim(ChangeFileExt(ExtractFileName(FileB), '')) + '(匹配成功记录).txt';
  fileNameout := ExtractFilePath(FileB) + Trim(ChangeFileExt(ExtractFileName(FileB), '')) + '(匹配成功记录).xlsx';
  isql := Format('SELECT A.* FROM [%s] AS A INNER JOIN [%s] AS B ON A.条码内容=B.条码内容 AND A.验证码=B.验证码', [tableName1, tableName2]);
  //isql := 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr(isql) + ',' + QuotedStr(fileNamein) + ';';
  try
    //ADOConnection1.Execute(isql);
    adoCheck3.SQL.Text := isql;
    adoCheck3.Open;
    DataToExcel(fileNamein, fileNameout);
    //DbGridEhToExcel(fileNameout);
  except
    on e: Exception do
    begin
      mmoCustFile.Lines.Add(Format('对比' + FileA + '和' + FileB + '文件失败，异常信息%s', [e.Message]));
      if FileExists(fileNameout) then
        DeleteFile(fileNameout);
      exit;
    end;
  end;
  if FileSizeByName(fileNameout) > 0 then
    mmoCustFile.Lines.Add(Format(FileA + '和' + FileB + '存在匹配成功的记录，请到【%s】查看', [fileNameout]))
  else
    DeleteFile(fileNameout);

  mmoCustFile.Lines.Add('匹配' + FileA + '和' + FileB + '完成');
end;

procedure TfrmQRCode.Button2Click(Sender: TObject);
var
  txtpath: string;
  sDateTime: TDateTime;
  sheetIndex, i, Row, Col: Integer;
  temp: string;
begin
  txtpath := Trim(edtTxt.Text);
  if txtpath = '' then
  begin
    ShowMessage('没有选择品检文件');
    Exit;
  end;
  if not qryCustFile.Active then
  begin
    ShowMessage('没有选择要导入的文件');
    Exit;
  end;
  if qryCustFile.IsEmpty then
  begin
    ShowMessage('没有选择要导入的文件');
    Exit;
  end;
  mmoCustFile.Lines.Clear;
  mmoCustFile.Lines.Add('匹配可能需要几分钟，请耐心等待...');
  sDateTime := now;
  Application.ProcessMessages;
  try
    qryCustFile.DisableControls;
    qryCustFile.First;
    while not qryCustFile.Eof do
    begin
      if qryCustFile.FieldByName('selected').AsBoolean then
      begin
        if (ExtractFileExt(qryCustFile.FieldByName('fileName').AsString) = '.xls') or (ExtractFileExt(qryCustFile.FieldByName('fileName').AsString) = '.xlsx') then
          MyCheck3(txtpath, qryCustFile.FieldByName('fileName').AsString, ' ');
      end;
      qryCustFile.Next;
    end;
    mmoCustFile.Lines.Add('恭喜，已全部匹配完成，共耗时：' + IntToStr(SecondsBetween(sDateTime, Now)) + '秒！');
  finally
    qryCustFile.EnableControls;
  end;

end;

procedure TfrmQRCode.CheckListBoxAClickCheck(Sender: TObject);
begin
  if CheckListBoxA.Checked[CheckListBoxA.ItemIndex] then
    FCheckListBoxSelectCountA := FCheckListBoxSelectCountA + 1
  else
    FCheckListBoxSelectCountA := FCheckListBoxSelectCountA - 1;
end;

procedure TfrmQRCode.CheckListBoxBClickCheck(Sender: TObject);
begin
  if CheckListBoxB.Checked[CheckListBoxB.ItemIndex] then
    FCheckListBoxSelectCountB := FCheckListBoxSelectCountB + 1
  else
    FCheckListBoxSelectCountB := FCheckListBoxSelectCountB - 1;
end;

procedure TfrmQRCode.btnCheckClick(Sender: TObject);
var
  datePointer: TDateTime;
  fileName, NotFileName, isql, tableName: string;
  strList: TStringList;
  iXLS: TXLSReadWriteII5;
  Row: Integer;
  iQuery: TADOQuery;
  HFileRes: HFILE;
  ExcelApp: Variant;
  num: Double;
  totalcon: Double;
begin
  totalcon := 0;
  if Trim(edtCheckA.Text) = '' then
  begin
    MessageBox(Self.Handle, '子文件不能为空，请检查！', '提示', MB_OK);
    Exit;
  end
  else if not FileExists(edtCheckA.Text) then
  begin
    MessageBox(Self.Handle, '子文件不存在，请检查！', '提示', MB_OK);
    Exit;
  end
  else if not chkAuto.Checked then
  begin
    if Trim(edtCheckB.Text) = '' then
    begin
      MessageBox(Self.Handle, '非自动匹配父文件模式下，父文件不能为空，请检查！', '提示', MB_OK);
      Exit;
    end
    else if not FileExists(edtCheckB.Text) then
    begin
      MessageBox(Self.Handle, '非自动匹配父文件模式下，父文件不存在，请检查！', '提示', MB_OK);
      Exit;
    end
  end
  else if chkAuto.Checked then
  begin
    if Trim(edtCustNy.Text) = '' then
    begin
      MessageBox(Self.Handle, '自动匹配父文件模式下，导入客户文件年月不能为空，请检查！', '提示', MB_OK);
      Exit;
    end;
  end;
  if not ShowAsk(0, PChar('您确认对比的文件和生成的文件格式正确无误吗？')) then
    Exit;
  btnCheck.Enabled := not btnCheck.Enabled;
  mmoLog1.Lines.Add('可能耗费较多时间，正在匹配中...');
  datePointer := Now;

  try
    iQuery := TADOQuery.Create(Self);
    try
      //非自动匹配父文件模式，只有一个父包
      if not chkAuto.Checked then
      begin
        with iQuery do
        begin
          CommandTimeout := 0;
          Connection := ADOConnection1;
          Close;
          SQL.Clear;
          SQL.Add('TRUNCATE TABLE tb_middle;' + 'INSERT INTO tb_middle(code,number) SELECT 条码内容,验证码 FROM tb_father WHERE exists(SELECT 条码内容 FROM tb_child WHERE tb_child.条码内容=tb_father.条码内容) order by tb_father.条码内容;');
          SQL.Add('TRUNCATE TABLE tb_middle_Not;' + 'INSERT INTO tb_middle_Not(code) SELECT 条码内容 FROM tb_child WHERE not exists(SELECT 条码内容 FROM tb_father WHERE tb_child.条码内容=tb_father.条码内容) order by tb_child.条码内容;');
          ExecSQL;
        end;
        if chk2.Checked then//保存txt文件,文件名以子文件名命名
        begin
          fileName := ExtractFilePath(edtCheckB.Text) + Trim(ChangeFileExt(ExtractFileName(edtCheckB.Text), '')) + '【匹配成功的文件】.txt';
          NotFileName := ExtractFilePath(edtCheckB.Text) + Trim(ChangeFileExt(ExtractFileName(edtCheckB.Text), '')) + '【未匹配成功的文件】.txt';
          isql := 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('SELECT code+'' ''+number FROM DB_QRCode..tb_middle') + ',' + QuotedStr(fileName) + ';';
          isql := isql + 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('SELECT code FROM DB_QRCode..tb_middle_Not') + ',' + QuotedStr(NotFileName) + ';';
          ADOConnection1.Execute(isql);
        end
        else if chk3.Checked then//保存xlsx文件,文件名以子文件名命名
        begin
          iXLS := TXLSReadWriteII5.Create(self);
          try
            with iQuery do
            begin
              Connection := ADOConnection1;
              CommandTimeout := 0;
              Close;
              SQL.Text := 'select * from tb_middle';
              Open;
              if IsEmpty then
              begin
                mmoLog1.Lines.Add('没有匹配成功的数据，未生成文件!!!');
                btnCheck.Enabled := True;
                Exit;
              end
              else
              begin
                fileName := ExtractFilePath(edtCheckB.Text) + Trim(ChangeFileExt(ExtractFileName(edtCheckB.Text), '')) + '【匹配成功的文件】.xlsx';
                if FileExists(fileName) then
                  DeleteFile(fileName);
                ExcelApp := CreateOleObject('Excel.Application');
                try
                  ExcelApp.WorkBooks.Add;
                  ExcelApp.workbooks[1].SaveAs(fileName);
                finally
                  ExcelApp.WorkBooks.Close;
                  ExcelApp.Quit;
                  ExcelApp := unassigned;
                end;
                Row := 0;
                XLS.Filename := fileName;
                XLS.Read;
                XLS.Sheets[0].AsString[0, Row] := '序号';
                XLS.Sheets[0].AsString[1, Row] := '二维码';
                XLS.Sheets[0].AsString[2, Row] := '验证码';
                XLS.Sheets[0].AsString[3, Row] := '子批次号';
                Row := 1;

                DisableControls;
                try
                  First;
                  while not Eof do
                  begin
                    XLS.Sheets[0].AsString[0, Row] := FieldByName('index1').AsString;
                    XLS.Sheets[0].AsString[1, Row] := FieldByName('code').AsString;
                    XLS.Sheets[0].AsString[2, Row] := FieldByName('number').AsString;
                    XLS.Sheets[0].AsString[3, Row] := Trim(edtBatchNo.Text);
                    Next;
                    Inc(Row);
                  end;
                  XLS.Write;
                finally
                  EnableControls;
                end;
              end;
            end;
          finally
            iXLS.Free;
          end;
        end;

        with ADOQuery1 do
        begin
          Close;
          SQL.Text := 'select count(1) as count from tb_middle';
          Open;
          if FieldByName('count').AsInteger > 0 then
            mmoLog1.Lines.Add('匹配完成，生成文件【' + fileName + '】！共匹配：' + ADOQuery1.FieldByName('count').AsString + '条记录，共耗时：' + IntToStr(SecondsBetween(Now, datePointer)) + '秒！');
        end;

        if chk2.Checked then  //生成文本-生成未匹配成功的文件
        begin
          with ADOQuery1 do
          begin
            Close;
            SQL.Text := 'select count(1) as count from tb_middle_Not';
            Open;
            if FieldByName('count').AsInteger > 0 then
              mmoLog1.Lines.Add('有未匹配成功的二维码，生成文件【' + notFileName + '】！未匹配成功的二维码共：' + ADOQuery1.FieldByName('count').AsString + '条记录，共耗时：' + IntToStr(SecondsBetween(Now, datePointer)) + '秒！');
          end;
        end;
      end
      //自动匹配父文件模式，只有多个父包，遍历父包
      else
      begin
        //找到指定父包
        ADOQuery1.Connection := ADOConnection1;
        ADOQuery1.CommandTimeout := 0;
        ADOQuery1.Close;
        ADOQuery1.SQL.Clear;
        ADOQuery1.SQL.Add('select name from sysobjects where xtype=''U'' and convert(char(6),crdate,112)=' + QuotedStr(Trim(edtCustNy.Text)));
        ADOQuery1.SQL.Add('and name not like''QR%'' and name<>''tableA'' and name<>''tableB'' and name<>''tableC''');
        ADOQuery1.SQL.Add('and name<>''tb_child'' and name<>''tb_father'' and name<>''tb_middle'' and name<>''tb_middle_Not'' order by crdate desc');
        ADOQuery1.Open;

        ADOQuery1.DisableControls;
        try
          ADOQuery1.First;
          while not ADOQuery1.Eof do
          begin
            with iQuery do
            begin
              tableName := ADOQuery1.FieldByName('name').AsString;
              Connection := ADOConnection1;
              CommandTimeout := 0;
              Close;
              SQL.Text := 'TRUNCATE TABLE tb_middle;' + 'INSERT INTO tb_middle(code,number) SELECT 条码内容,验证码 FROM dbo.[' + tableName + '] WHERE exists(SELECT 条码内容 FROM tb_child WHERE tb_child.条码内容=dbo.[' + tableName + '].条码内容) order by 条码内容;';
              ExecSQL;

              Close;
              SQL.Text := 'select * from tb_middle';
              Open;

              if not IsEmpty then
              begin
                mmoLog1.Lines.Add('匹配到父包【' + tableName + '】，共匹配到【' + inttostr(iQuery.RecordCount) + '】条记录!');
                totalcon := totalcon + iQuery.RecordCount;
                if chk2.Checked then //保存txt文件,文件名以父文件名命名
                begin
                  NotFileName := ExtractFilePath(edtCheckA.Text) + Trim(ChangeFileExt(ExtractFileName(tableName), '')) + '【未匹配成功的文件】.txt';
                  fileName := ExtractFilePath(edtCheckA.Text) + Trim(ChangeFileExt(ExtractFileName(tableName), '')) + '【匹配成功的文件】.txt';
                  isql := 'TRUNCATE TABLE tb_middle_not;' + 'INSERT INTO tb_middle_not(code) SELECT 条码内容 FROM dbo.[' + tableName + '] WHERE not exists(SELECT 条码内容 FROM tb_child WHERE tb_child.条码内容=dbo.[' + tableName + '].条码内容) order by 条码内容;';
                  isql := isql + 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('SELECT code+'' ''+number FROM DB_QRCode..tb_middle') + ',' + QuotedStr(fileName) + ';';
                  isql := isql + 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('SELECT code FROM DB_QRCode..tb_middle_Not') + ',' + QuotedStr(NotFileName) + ';';
                  ADOConnection1.Execute(isql);
                end
                else if chk3.Checked then  //保存xlsx文件,文件名以父文件名命名
                begin
                  //把数据库的数据写到xlsx
                  fileName := ExtractFilePath(edtCheckA.Text) + Trim(ChangeFileExt(ExtractFileName(tableName), '')) + '【匹配成功的文件】.xlsx';
                  if FileExists(fileName) then
                    DeleteFile(fileName);
                  ExcelApp := CreateOleObject('Excel.Application');
                  try
                    ExcelApp.WorkBooks.Add;
                    ExcelApp.workbooks[1].SaveAs(fileName);
                  finally
                    ExcelApp.WorkBooks.Close;
                    ExcelApp.Quit;
                    ExcelApp := unassigned;
                  end;
                  Row := 0;
                  XLS.Filename := fileName;
                  XLS.Read;
                  XLS.Sheets[0].AsString[0, Row] := '序号';
                  XLS.Sheets[0].AsString[1, Row] := '二维码';
                  XLS.Sheets[0].AsString[2, Row] := '验证码';
                  XLS.Sheets[0].AsString[3, Row] := '子批次号';
                  Row := 1;
                  num := 0;
                  DisableControls;
                  try
                    First;
                    while not Eof do
                    begin                                 //按序号重新排列
                      XLS.Sheets[0].AsString[0, Row] := FloatToStr(num); //FieldByName('index1').AsString;
                      XLS.Sheets[0].AsString[1, Row] := FieldByName('code').AsString;
                      XLS.Sheets[0].AsString[2, Row] := FieldByName('number').AsString;
                      XLS.Sheets[0].AsString[3, Row] := Trim(edtBatchNo.Text);
                      //XLS.Sheets[0].AsString[4, Row] := FloatToStr(num);
                      num := num + 1;
                      Next;
                      Inc(Row);
                    end;
                    XLS.Write;
                  finally
                    EnableControls;
                  end;
                end;
                mmoLog1.Lines.Add('生成文件【' + fileName + '】成功!!!');
                if chk2.Checked then  //生成文本-生成未匹配成功的文件
                begin
                  with iQuery do
                  begin
                    Connection := ADOConnection1;
                    Close;
                    SQL.Text := 'select count(1) as count from tb_middle_Not';
                    Open;
                    if FieldByName('count').AsInteger > 0 then
                      mmoLog1.Lines.Add('有未匹配成功的二维码，生成文件【' + notFileName + '】！未匹配成功的二维码共：' + FieldByName('count').AsString + '条记录，共耗时：' + IntToStr(SecondsBetween(Now, datePointer)) + '秒！');
                  end;
                end;
              end;
            end;
            ADOQuery1.Next;
          end;
        finally
          ADOQuery1.EnableControls;
        end;
        if Trim(fileName) = '' then
          mmoLog1.Lines.Add('未匹配到任何数据！')
        else
        begin
          mmoLog1.Lines.Add('子包共匹配到【' + floattostr(totalcon) + '】条记录!');
          mmoLog1.Lines.Add('恭喜，已全部匹配完成，共耗时：' + IntToStr(SecondsBetween(Now, datePointer)) + '秒！');
        end;
      end;
    finally
      iQuery.Free;
    end;
  except
    on e: Exception do
    begin
      mmoLog1.Lines.Add('匹配失败，异常信息:' + E.Message + '！');
      btnCheck.Enabled := not btnCheck.Enabled;
      Exit;
    end;
  end;

  btnCheck.Enabled := True;
end;

procedure TfrmQRCode.chk2Click(Sender: TObject);
begin
  if chk2.Checked then
    chk3.Checked := False
  else
    chk3.Checked := True;
end;

procedure TfrmQRCode.chk3Click(Sender: TObject);
begin
  if chk3.Checked then
    chk2.Checked := False
  else
    chk2.Checked := True;
end;

procedure TfrmQRCode.chkAutoClick(Sender: TObject);
begin
  if chkAuto.Checked then
  begin
    lbl3.Visible := False;
    edtCheckB.Visible := False;
    btnImportCheckB.Visible := False;
    btnClearCheckB.Visible := False;
    Label25.Visible := True;
    edtCustNy.Visible := True;
    Label27.Visible := True;
  end
  else
  begin
    lbl3.Visible := True;
    edtCheckB.Visible := True;
    btnImportCheckB.Visible := True;
    btnClearCheckB.Visible := True;
    Label25.Visible := False;
    edtCustNy.Visible := False;
    Label27.Visible := False;
  end;
end;

procedure TfrmQRCode.gridCustFileMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if (UpperCase(gridCustFile.SelectedField.FieldName) = UpperCase('selected')) or (UpperCase(gridCustFile.SelectedField.FieldName) = UpperCase('isContainFirst')) then
    gridCustFile.ReadOnly := False
  else
    gridCustFile.ReadOnly := True;
end;

procedure TfrmQRCode.gridMultiMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if (UpperCase(gridMulti.SelectedField.FieldName) = UpperCase('selected')) or (UpperCase(gridMulti.SelectedField.FieldName) = UpperCase('isContainFirst')) then
    gridMulti.ReadOnly := False
  else
    gridMulti.ReadOnly := True;
end;

procedure TfrmQRCode.edtNoExit(Sender: TObject);
var
  iQuery: TADOQuery;
begin
  if Trim(edtNo.Text) = '' then
    Exit;
  if Trim(Copy(edtNo.Text, 1, 2)) <> 'QR' then
  begin
    PageControl1.ActivePageIndex := 0;
    ShowMessage('主表单据号必须以QR开头!!!');
    edtNo.SetFocus;
    Exit;
  end;
  iQuery := TADOQuery.Create(Self);
  try
    with iQuery do
    begin
      CommandTimeout := 0;
      Connection := conOA;
      Close;
      SQL.Text := 'select * from CustomModule_201611erweimasq where MainID=' + QuotedStr(Trim(edtNo.Text));
      Open;
      if IsEmpty then
      begin
        ShowMessage('主表单据号输入有误，请检查！');
        edtNo.SetFocus;
      end
      else
      begin
        if Trim(FieldByName('hexing').AsString) = '条盒' then
          edtType.ItemIndex := 0
        else if Trim(FieldByName('hexing').AsString) = '小盒' then
          edtType.ItemIndex := 1;
        if FieldByName('gcbh').AsString = '广东中烟' then
          edtMode.Text := '0';
        edtWebAddr.Text := Trim(FieldByName('wz').AsString);
        edtBrandCode.Text := Copy(Trim(FieldByName('ppdm').AsString), 1, Pos('：', FieldByName('ppdm').AsString) - 1);
        edtPrinterCode.Text := Copy(Trim(FieldByName('yscdm').AsString), 1, Pos('：', FieldByName('yscdm').AsString) - 1);
        edtNy.Text := Trim(FieldByName('scnn').AsString);
        edtCount.Text := Trim(FieldByName('scsl').AsString);
        edtCount.Text := Trim(FieldByName('scsl').AsString);
        edtGdbh.Text := Trim(FieldByName('gdbh').AsString);
        edtYpbh.Text := Trim(FieldByName('ypbh').AsString);
        edtYpmc.Text := Trim(FieldByName('ypmc').AsString);
      end;
    end;
  finally
    iQuery.Free;
  end;
end;

procedure TfrmQRCode.FormCreate(Sender: TObject);
begin

  if DateToStr(Now())>'2018-03-01' then
  begin
    ShowMessage('试用日期已到，请联系客服！');
    Application.Terminate;
  end;

  InitializeCriticalSection(Cs);
  DragAcceptFiles(self.Handle, True); //调用windows API
  ProgressBar1.Parent := StatusBar1;
  ProgressBar1.Left := StatusBar1.Panels[0].Width + 2;
  ProgressBar1.Top := 0;
  ProgressBar1.Width := StatusBar1.Width - StatusBar1.Panels[0].Width - 2;
  ProgressBar1.Height := StatusBar1.Height;
end;

procedure TfrmQRCode.FormDestroy(Sender: TObject);
begin
  DeleteCriticalSection(Cs);
end;

procedure TfrmQRCode.FormResize(Sender: TObject);
begin
  ProgressBar1.Parent := StatusBar1;
  ProgressBar1.Left := StatusBar1.Panels[0].Width + 2;
  ProgressBar1.Top := 0;
  ProgressBar1.Width := StatusBar1.Width - StatusBar1.Panels[0].Width - 2;
  ProgressBar1.Height := StatusBar1.Height;
end;

procedure TfrmQRCode.FormShow(Sender: TObject);
begin
  PageControl1.ActivePageIndex := 3;
  btnConnectClick(nil);
  if ADOConnection1.Connected then
  begin
    with qryCustFile do
    begin
      ADOConnection1.Execute('if not exists(select 1 from sysobjects where name=''##tmpCustFile'' and xtype=''U'') create table ##tmpCustFile(selected bit DEFAULT 1,fileName varchar(500),isContainFirst bit DEFAULT 1)');
      Connection := ADOConnection1;
      Close;
      sql.Text := 'select * from ##tmpCustFile';
      Open;
    end;

    with qryMulti do
    begin
      ADOConnection1.Execute('if not exists(select 1 from sysobjects where name=''##tmpMulti'' and xtype=''U'') create table ##tmpMulti(selected bit DEFAULT 1,fileName varchar(500),isContainFirst bit DEFAULT 1)');
      Connection := ADOConnection1;
      Close;
      sql.Text := 'select * from ##tmpMulti';
      Open;
    end;
  end;
end;

procedure TfrmQRCode.qryCustFileAfterOpen(DataSet: TDataSet);
begin
  if DataSet.Active then
  begin
    DataSet.FieldByName('selected').DisplayLabel := '选择';
    DataSet.FieldByName('selected').DisplayWidth := 5;
    DataSet.FieldByName('fileName').DisplayLabel := '文件';
    DataSet.FieldByName('fileName').DisplayWidth := 60;
    DataSet.FieldByName('isContainFirst').DisplayLabel := '首行包含列名称';
    DataSet.FieldByName('isContainFirst').DisplayWidth := 15;
  end;
end;

procedure TfrmQRCode.qryMultiAfterOpen(DataSet: TDataSet);
begin
  if DataSet.Active then
  begin
    DataSet.FieldByName('selected').DisplayLabel := '选择';
    DataSet.FieldByName('selected').DisplayWidth := 5;
    DataSet.FieldByName('fileName').DisplayLabel := '文件';
    DataSet.FieldByName('fileName').DisplayWidth := 80;
    DataSet.FieldByName('isContainFirst').DisplayLabel := '首行包含列名称';
    DataSet.FieldByName('isContainFirst').DisplayWidth := 15;
  end;
end;

procedure TfrmQRCode.Searchfile(path: PChar; fileExt: string; fileList: TStringList);
var
  searchRec: TSearchRec;
  found: Integer;
  tmpStr: string;
  curDir: string;
  dirs: TQueue;
  pszDir: PChar;
begin
  dirs := TQueue.Create; //创建目录队列
  dirs.Push(path); //将起始搜索路径入队
  pszDir := dirs.Pop;
  curDir := StrPas(pszDir); //出队
   {开始遍历,直至队列为空(即没有目录需要遍历)}
  while (True) do
  begin
    //加上搜索后缀,得到类似'c:\*.*' 、'c:\windows\*.*'的搜索路径
    tmpStr := curDir + '\*.*';
    //在当前目录查找第一个文件、子目录
    found := FindFirst(tmpStr, faAnyFile, searchRec);
    while found = 0 do //找到了一个文件或目录后
    begin
      //如果找到的是个目录
      if (searchRec.Attr and faDirectory) <> 0 then
      begin
        {在搜索非根目录(C:\、D:\)下的子目录时会出现'.','..'的"虚拟目录"
         大概是表示上层目录和下层目录吧。。。要过滤掉才可以}
        if (searchRec.Name <> '.') and (searchRec.Name <> '..') then
        begin
           {由于查找到的子目录只有个目录名，所以要添上上层目录的路径
            searchRec.Name = 'Windows';
            tmpStr:='c:\Windows';
            加个断点就一清二楚了
           }
          tmpStr := curDir + '\' + searchRec.Name;
               {将搜索到的目录入队。让它先晾着。
                因为TQueue里面的数据只能是指针,所以要把string转换为PChar
                同时使用StrNew函数重新申请一个空间存入数据，否则会使已经进
                入队列的指针指向不存在或不正确的数据(tmpStr是局部变量)。}
          dirs.Push(StrNew(PChar(tmpStr)));
        end;
      end
      else //如果找到的是个文件
      begin
         {Result记录着搜索到的文件数。可是我是用CreateThread创建线程
          来调用函数的，不知道怎么得到这个返回值。。。我不想用全局变量}
        //把找到的文件加到Memo控件
        if fileExt = '.*' then
          fileList.Add(curDir + '\' + searchRec.Name)
        else
        begin
          if SameText(RightStr(curDir + '\' + searchRec.Name, Length(fileExt)), fileExt) then
            fileList.Add(curDir + '\' + searchRec.Name);
        end;
      end;
      //查找下一个文件或目录
      found := FindNext(searchRec);
    end;
    {当前目录找到后，如果队列中没有数据，则表示全部找到了；
      否则就是还有子目录未查找，取一个出来继续查找。}
    if dirs.Count > 0 then
    begin
      pszDir := dirs.Pop;
      curDir := StrPas(pszDir);
      StrDispose(pszDir);
    end
    else
      break;
  end;
  //释放资源
  dirs.Free;
  FindClose(searchRec);
end;

function TfrmQRCode.ShowAsk(MHandle: Thandle; tText: Pchar): Boolean;
begin
  if MessageBox(MHandle, tText, '询问信息', MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2) = IDYES then
    Result := True
  else
    Result := False;
end;

procedure TfrmQRCode.WmDropFiles(var Msg: TMessage);
var
  P: array[0..254] of Char;
  i: Word;
begin
  inherited;
  i := DragQueryFile(Msg.wParam, $FFFFFFFF, nil, 0);
  for i := 0 to i - 1 do
  begin
    DragQueryFile(Msg.wParam, i, P, 255);
    if (ExtractFileExt(StrPas(P)) = '.txt') or (ExtractFileExt(StrPas(P)) = '.xls') or (ExtractFileExt(StrPas(P)) = '.xlsx') then
    begin
      with qryCustFile do
      begin
        //如果存在相同的文件
        DisableControls;
        try
          First;
          while not Eof do
          begin
            if Trim(ChangeFileExt(ExtractFileName(FieldByName('fileName').AsString), '')) = Trim(ChangeFileExt(ExtractFileName(StrPas(P)), '')) then
            begin
              ShowMessage('已经存在文件【' + Trim(ChangeFileExt(ExtractFileName(FieldByName('fileName').AsString), '')) + '】!!!');
              Exit;
            end;
            Next;
          end;
        finally
          EnableControls;
        end;
        if not (State in [dsEdit, dsInsert]) then
          Append;
        FieldByName('selected').AsBoolean := True;
        FieldByName('fileName').AsString := StrPas(P);
        FieldByName('isContainFirst').AsBoolean := True;
        if not (State in [dsBrowse]) then
          Post;
      end;
    end;
  end;
end;

end.

