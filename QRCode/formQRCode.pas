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
    procedure WmDropFiles(var Msg: TMessage); message WM_DROPFILES; //����WM_DROPFILES��Ϣ�Ĺ���
    { Public declarations }
  end;

var
  frmQRCode: TfrmQRCode;
  Cs: TRTLCriticalSection;

implementation
{$R *.dfm}

//�ļ������ɹ��������ݺ�(�����š���,ӡƷ��š���,ӡƷ���ơ���,���͡���,��ά����������)

procedure TfrmQRCode.btnImportCheckAClick(Sender: TObject);
var
  iQuery: TADOQuery;
  tableName, isql: string;
  i: Integer;
begin
  if FCheckListBoxSelectCountA <= 0 then
  begin
    ShowMessage('��ѡ�����ļ��ĸ�ʽ��');
    Exit;
  end;

  OpenDialogA.FileName := '�ı��ļ�(���ļ�)';
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
        //�������ļ���ʽ��̬������
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
        mmoLog1.Lines.Add('���ļ��ܼ�¼[' + FieldByName('count').AsString + ']��');
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
    ShowMessage('��ѡ���ļ��ĸ�ʽ��');
    Exit;
  end;

  OpenDialogB.FileName := '�ı��ļ�';
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
        //����xls��xlsx
        if (ExtractFileExt(XLS.Filename) = '.xls') or (ExtractFileExt(XLS.Filename) = '.xlsx') then
        begin
          chk1.Visible := True;
          chk1.Checked := True;
          //  XLS.DirectRead := True; //ȡ���˲���
          XLS.Read;
          for sheetIndex := 0 to 99 do //���ɵ���100��Sheet
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
              Break;  //�����쳣
            end;
          end;
          fileName := ExtractFilePath(Trim(edtCheckB.Text)) + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '(��ʱ�ļ�).txt';
          //����txt�ļ�
          list.SaveToFile(fileName);
          tableName := Trim(ChangeFileExt(ExtractFileName(XLS.Filename), ''));
          //��txt�ļ����浽���ݿ�
          Connection := ADOConnection1;
          Close;
          isql := 'if Exists(select 1 from sysobjects where name=''tb_father'' and xtype=''U'') drop table tb_father;';
          //�������ļ���ʽ��̬������
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
              mmoCustFile.Lines.Add('�����ļ���' + fileName + '��ʧ�ܣ��쳣��Ϣ��' + e.message);
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
          //�������ļ���ʽ��̬������
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
//          SQL.Add('Update tb_father set ��������=�����κ�,�����κ�=''''');
          ExecSQL;
        end;
        Close;
        SQL.Text := 'select count(1) as count from tb_father';
        Open;
        mmoLog1.Lines.Add('���ļ��ܼ�¼[' + FieldByName('count').AsString + ']��');
      end;
    end;
  finally
    btnCheck.Enabled := True;
    iQuery.Free;
    list.Free;
  end;
end;

 //�������ַ������ַ������һ�γ��ֵ�λ�� alex add 2017-09-18
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
    ShowMessage('û��ѡ��Ҫ������ļ�');
    Exit;
  end;
  if qryCustFile.IsEmpty then
    Exit;


  //alex yu edit 2017 09-18
  {
    ��һ����Ϊ���ݵ����excel�����ĸ��ļ�(txt)����ʱ����_����_,Ʒ��_
  }
  mmoCustFile.Lines.Clear;
  mmoCustFile.Lines.Add('���������Ҫ�����ӣ����ڵ����У������ĵȴ�...');
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
//          XLS.DirectRead := True; //ȡ���˲���
          XLS.Read;
          for sheetIndex := 0 to 99 do //���ɵ���100��Sheet
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
                    //Ʒ���ʽ ��������,��֤��
                    //������ʽ �����κ� ��� �������� ��֤��  ��ע���ѿո������������������ո񣬴����ĸ���λ��
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
              Break;  //�����쳣
            end;
          end;
          //fileName := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '(������ļ�).txt';
          fileName := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + '��_' + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '.txt';
          //����txt�ļ�
          list.SaveToFile(fileName);
          fileName2 := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + '��_' + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '.txt';
          CopyFile(PChar(fileName), PChar(fileName2), False);
          fileName := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + 'Ʒ��_' + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '.txt';
          //����txt�ļ�
          list2.SaveToFile(fileName);
        end;

        //����xls��xlsx
        list.Clear;
        if (ExtractFileExt(XLS.Filename) = '.xls') or (ExtractFileExt(XLS.Filename) = '.xlsx') then
        begin
//          XLS.DirectRead := True; //ȡ���˲���
          XLS.Read;
          for sheetIndex := 0 to 99 do //���ɵ���100��Sheet
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
              Break;  //�����쳣
            end;
          end;
          fileName := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '(��ʱ�ļ�).txt';
          //����txt�ļ�
          list.SaveToFile(fileName);
          tableName := Trim(ChangeFileExt(ExtractFileName(XLS.Filename), ''));
          //��txt�ļ����浽���ݿ�

          isql := 'if not Exists(select 1 from sysobjects where name=' + QuotedStr(tableName) + ' and xtype=''U'') create table [' + tableName + '](�����κ� varchar(100),��� varchar(50),�������� varchar(100),��֤�� varchar(100)) else TRUNCATE TABLE [' + tableName + '];';
          isql := isql + 'BULK INSERT [' + tableName + '] FROM ' + QuotedStr(fileName) + 'WITH(FIELDTERMINATOR ='' '',ROWTERMINATOR =''\n'')';
          try
            ADOConnection1.Execute(isql);
          except
            on e: Exception do
            begin
              mmoCustFile.Lines.Add('�����ļ���' + fileName + '��ʧ�ܣ��쳣��Ϣ��' + e.message);
              exit;
            end;
          end;
        end
        else if (ExtractFileExt(XLS.Filename) = '.txt') then //����txt
        begin
          fileName := qryCustFile.FieldByName('fileName').AsString;
          tableName := Trim(ChangeFileExt(ExtractFileName(qryCustFile.FieldByName('fileName').AsString), ''));
          //��txt�ļ����浽���ݿ�
          isql := 'if not Exists(select 1 from sysobjects where name=' + QuotedStr(tableName) + ' and xtype=''U'') create table [' + tableName + '](�������� varchar(100),��֤�� varchar(100)) else TRUNCATE TABLE [' + tableName + '];';
          isql := isql + 'BULK INSERT [' + tableName + '] FROM ' + QuotedStr(fileName) + 'WITH(FIELDTERMINATOR ='' '',ROWTERMINATOR =''\n'')';
          try
            ADOConnection1.Execute(isql);
          except
            on e: Exception do
            begin
              mmoCustFile.Lines.Add('�����ļ���' + fileName + '��ʧ�ܣ��쳣��Ϣ��' + e.message);
              exit;
            end;
          end;
        end;
        mmoCustFile.Lines.Add('�����ļ���' + fileName + '���ɹ�!!!');
      end;
      qryCustFile.Next;
    end;
    mmoCustFile.Lines.Add('��ϲ���ļ���ȫ������ɹ�!!!');
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
    mmoCustFile.Lines.Add('���ڿ�ʼƥ���ļ�!!!');
  //ShowMessage(IntToStr(qryCustFile.RecordCount));
    qryCustFile.First;
    while not qryCustFile.Eof do
    begin
    //ShowMessage('aaaaaa');
    //ȡ���ұ�\��λ��
      index := LastPos('\', qryCustFilefileName.Value);
    //ȡ�ļ�·��
      spath := Copy(qryCustFilefileName.Value, 0, index);
    //ȡ�ļ���
      Sfilename := Copy(qryCustFilefileName.Value, index + 1, Length(qryCustFilefileName.Value));
    //ȡȥ��չ�����ļ��������ں�������txt�ļ�����ֻ�����չ�����Ѿ�����.��
      txtname := Copy(Sfilename, 0, LastPos('.', Sfilename));

      txtpath := spath + '��_' + txtname + 'txt';
      {
      ListArray[1].Clear;
      ListArray[1].LoadFromFile(txtpath);
      ListArray[2].AddStrings(ListArray[1]);
      }
      autoCheck(txtpath, qryCustFilefileName.Text, true, True);

      txtpath := spath + '��_' + txtname + 'txt';
      {
      ListArray[3].Clear;
      ListArray[3].LoadFromFile(txtpath);
      ListArray[4].AddStrings(ListArray[3]);
      }
      autoCheck(txtpath, qryCustFilefileName.Text, true, True);

      txtpath := spath + 'Ʒ��_' + txtname + 'txt';
      ListArray[5].Clear;
      ListArray[5].LoadFromFile(txtpath);
      ListArray[6].AddStrings(ListArray[5]);
      autoCheck(txtpath, qryCustFilefileName.Text, False, True);
      qryCustFile.Next;
    end;
    //ListArray[2].SaveToFile(spath+'��_�ϲ�.txt');
    //ListArray[4].SaveToFile(spath+'��_�ϲ�.txt');
    ListArray[6].SaveToFile(spath + 'Ʒ��_�ϲ�.txt');
    mmoCustFile.Lines.Add('��ϲ����ȫ������ͶԱ���ɣ�����ʱ��' + IntToStr(SecondsBetween(sDateTime, Now)) + '�룡');
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
    ShowMessage('û��ѡ��Ҫ������ļ�');
    Exit;
  end;
  if qryMulti.IsEmpty then
  begin
    ShowMessage('û��ѡ��Ҫ������ļ�');
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
        //����xls��xlsx
        if (ExtractFileExt(XLS.Filename) = '.xls') or (ExtractFileExt(XLS.Filename) = '.xlsx') then
        begin
        //  XLS.DirectRead := True; //ȡ���˲���
          {XLS.Read;
          for sheetIndex := 0 to 99 do //���ɵ���100��Sheet
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
              Break;  //�����쳣
            end;
          end;
          fileName := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '(��ʱ�ļ�).txt';
          //����txt�ļ�
          list.SaveToFile(fileName);
          tableName := Trim(ChangeFileExt(ExtractFileName(XLS.Filename), ''));
          //��txt�ļ����浽���ݿ�
          isql := 'if not Exists(select 1 from sysobjects where name=' + QuotedStr(tableName) + ' and xtype=''U'') create table [' + tableName + '](�����κ� varchar(100),��� varchar(50),�������� varchar(100),��֤�� varchar(100)) else TRUNCATE TABLE [' + tableName + '];';
          isql := isql + 'BULK INSERT [' + tableName + '] FROM ' + QuotedStr(fileName) + 'WITH(FIELDTERMINATOR ='' '',ROWTERMINATOR =''\n'')';
          try
            ADOConnection1.Execute(isql);
          except
            on e: Exception do
            begin
              mmoCustFile.Lines.Add('�����ļ���' + fileName + '��ʧ�ܣ��쳣��Ϣ��' + e.message);
              exit;
            end;
          end; }
        end
        else if (ExtractFileExt(XLS.Filename) = '.txt') then //����txt
        begin
          fileName := qryMulti.FieldByName('fileName').AsString;
          //��txt�ļ����浽���ݿ�
          isql := 'BULK INSERT ##tb_multi FROM ' + QuotedStr(fileName) + 'WITH(FIELDTERMINATOR ='' '',ROWTERMINATOR =''\n'')';
          try
            ADOConnection1.Execute(isql);
          except
            on e: Exception do
            begin
              mmoMulti.Lines.Add('�����ļ���' + fileName + '��ʧ�ܣ��쳣��Ϣ��' + e.message);
              exit;
            end;
          end;
        end;
        mmoMulti.Lines.Add('�����ļ���' + fileName + '���ɹ�!!!');
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
    mmoMulti.Lines.Add('�ļ���ȫ������ɹ��������롾' + iQuery.FieldByName('count').AsString + '���ʣ���ʱ��' + IntToStr(SecondsBetween(eDate, sDate)) + '��!!!');
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
    ShowMessage('û��ѡ��Ҫ������ļ�');
    Exit;
  end;
  if qryMulti.IsEmpty then
  begin
    ShowMessage('û��ѡ��Ҫ������ļ�');
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
        mmoMulti.Lines.Add('���س�������ɡ�'+qryMulti.FieldByName('fileName').AsString+'��');
      end;
      qryMulti.Next;
    end;
    eDate:=Now;
    mmoMulti.Lines.Add('���س�������ɣ�����ʱ��' + IntToStr(SecondsBetween(eDate, sDate)) + '��!!!');
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
  OpenDialog.FileName := '�ı��ļ�';
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
          //���������ͬ���ļ�
          DisableControls;
          try
            First;
            while not Eof do
            begin
              if Trim(FieldByName('fileName').AsString) = Trim(OpenDialog.Files[i]) then
              begin
                ShowMessage('�Ѿ������ļ���' + Trim(FieldByName('fileName').AsString) + '��!!!');
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
    ShowMessage('�����ݺŲ���Ϊ��!!!');
    Exit;
  end;
  if Trim(edtPath.Text) = '' then
  begin
    ShowMessage('�����ļ�·������Ϊ��!!!');
    Exit;
  end;
  Memo2.Lines.Clear;

  if not DirectoryExists(path) then
    CreateDirectory(PWideChar(path), nil);

  if Trim(edtBrandCode.Text) = '01' then
    ppMc := '˫ϲ�����乤����'
  else if Trim(edtBrandCode.Text) = '02' then
    ppMc := '˫ϲ�����ã�'
  else if Trim(edtBrandCode.Text) = '03' then
    ppMc := '˫ϲ��ϲ���꣩'
  else if Trim(edtBrandCode.Text) = '04' then
    ppMc := '˫ϲ�����꾭�䣩'
  else if Trim(edtBrandCode.Text) = '05' then
    ppMc := '˫ϲ����ϲ��'
  else if Trim(edtBrandCode.Text) = '06' then
    ppMc := '˫ϲ����ϲϸ֧��'
  else if Trim(edtBrandCode.Text) = '07' then
    ppMc := '˫ϲ����ϲ��˫ϲ�����ã�';
  path := Trim(edtPath.Text);
  filename := trim(iTable) + '(�����š�' + edtGdbh.Text + '��,ӡƷ��š�' + edtYpbh.Text + '��,ӡƷ���ơ�' + edtYpmc.Text + '��,���͡�' + edtType.Text + '��,��ά��������' + edtCount.Text + '��)' + '.txt';

  //�����������ݺ��µĶ�汾���ݱ�
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
        begin  //�ļ������ɹ��������ݺ�(�����š���,ӡƷ��š���,ӡƷ���ơ���,���͡���,��ά����������)
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
            Memo1.Lines.Add('�����ļ���' + path + '\' + iTable + filename + '��ʧ�ܣ��쳣��Ϣ��' + e.Message);
            Exit;
          end;
        end;
        eDate := Now;
        Memo2.Lines.Add('�����ļ���' + path + '\' + iTable + filename + '���ɹ�����ʱ��' + IntToStr(SecondsBetween(eDate, sDate)) + '��!!!');
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
  OpenDialogA.FileName := '�ı��ļ�';
  OpenDialogA.Filter := '*.txt|*.txt';
  if OpenDialogA.Execute then
  begin
    path := Trim(OpenDialogA.Files.Text);
    //path:=StringReplace(path,#13#10,'',[rfReplaceAll,rfIgnoreCase]);//�ļ������лس����ţ�����sql��䱨��
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
  OpenDialogA.FileName := 'Excel�ļ�';
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
    ErrorStr := Format('��%d���%d��δѡ��ƥ���������', [zu, yu]);
    mmoLog2.Lines.Add(ErrorStr);
    Exit;
  end;
  if (FileA = '') and (FileB <> '') then
  begin
    ErrorStr := Format('��%d���%d����ƥ���ļ��ļ�����Ϊ�գ�����', [zu, yu]);
    mmoLog2.Lines.Add(ErrorStr);
    Exit;
  end;
  if (FileA <> '') and (FileB = '') then
  begin
    ErrorStr := Format('��%d���%d��Դ����Ϊ�գ�����', [zu, yu]);
    mmoLog2.Lines.Add(ErrorStr);
    Exit;
  end;
  if DelimA then
    Delim := ' '
  else if DelimB then
    Delim := ' ';

  Application.ProcessMessages;
  mmoLog2.Lines.Add(Format('��ʼƥ���%d���%d������ȴ�...', [zu, yu]));

  fileName1 := FileA;
  tableName1 := '##' + Trim(ChangeFileExt(ExtractFileName(fileName1), ''));
    //��txt�ļ����浽���ݿ�
  isql := 'drop table [' + tableName1 + ']';
  try
    ADOConnection1.Execute(isql);
  except
  end;
  isql := 'create table [' + tableName1 + '](�������� varchar(200),��֤�� varchar(100));';
  isql := isql + 'BULK INSERT [' + tableName1 + '] FROM ' + QuotedStr(fileName1) + 'WITH(FIELDTERMINATOR =' + quotedstr(Delim) + ',ROWTERMINATOR =''\n'')';
  try
    ADOConnection1.Execute(isql);
  except
    on e: Exception do
    begin
      mmoLog2.Lines.Add(Format('�����%d���%d���ļ���%s��ʧ�ܣ��쳣��Ϣ%s', [zu, yu, FileA, e.Message]));
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
      for sheetIndex := 0 to 99 do //���ɵ���100��Sheet
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
          Break;  //�����쳣
        end;
      end;
          //fileName := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '(������ļ�).txt';
      fileName := ExtractFilePath(FileB) + tableName2 + '.txt';
      list.SaveToFile(fileName);
    end;
  finally
    FreeAndNil(List);
  end;

  //��txt�ļ����浽���ݿ�
  if FileExists(fileName) = False then
  begin
    mmoLog2.Lines.Add(Format('�����%d���%d���ļ���%s��ʧ��', [zu, yu, FileB]));
    exit;
  end;

  isql := 'drop table [' + tableName2 + ']';
  try
    ADOConnection1.Execute(isql);
  except
  end;
  try
    isql := 'create table [' + tableName2 + '](�������� varchar(200),��֤�� varchar(100));';
    isql := isql + 'BULK INSERT [' + tableName2 + '] FROM ' + QuotedStr(fileName) + 'WITH(FIELDTERMINATOR =' + quotedstr(Delim) + ',ROWTERMINATOR =''\n'')';
    try
      ADOConnection1.Execute(isql);
    except
      on e: Exception do
      begin
        mmoLog2.Lines.Add(Format('�����%d���%d���ļ���%s��ʧ�ܣ��쳣��Ϣ%s', [zu, yu, FileB, e.Message]));
        exit;
      end;
    end;
  finally
    if FileExists(fileName) then
      DeleteFile(fileName);
  end;

  fileNameout := ExtractFilePath(FileA) + Trim(ChangeFileExt(ExtractFileName(FileA), '')) + '(δƥ���¼).txt';
  isql := Format('SELECT A.* FROM [%s] AS A LEFT JOIN [%s] AS B ON A.��������=B.�������� AND A.��֤��=B.��֤�� WHERE B.�������� IS NULL AND B.��֤��  IS NULL', [tableName1, tableName2]);
  isql := 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr(isql) + ',' + QuotedStr(fileNameout) + ';';
  try
    ADOConnection1.Execute(isql);
  except
    on e: Exception do
    begin
      mmoLog2.Lines.Add(Format('�Աȵ�%d���%d���ļ�ʧ�ܣ��쳣��Ϣ%s', [zu, yu, e.Message]));
      exit;
    end;
  end;
  if FileSizeByName(fileNameout) > 0 then
    mmoLog2.Lines.Add(Format('��%d���%d���ļ����ڲ�ƥ��ļ�¼���뵽��%s���鿴', [zu, yu, fileNameout]))
  else
    DeleteFile(fileNameout);
  mmoLog2.Lines.Add(Format('ƥ���%d���%d���', [zu, yu]));  }
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
  mmoCustFile.Lines.Add('��ʼƥ��' + FileA + '��' + FileB + '��ȴ�...');

  fileName1 := FileA;
  tableName1 := '##' + Trim(ChangeFileExt(ExtractFileName(fileName1), ''));
    //��txt�ļ����浽���ݿ�
  isql := 'drop table [' + tableName1 + ']';
  try
    ADOConnection1.Execute(isql);
  except
  end;
  isql := 'create table [' + tableName1 + '](�������� varchar(200),��֤�� varchar(100));';
  isql := isql + 'BULK INSERT [' + tableName1 + '] FROM ' + QuotedStr(fileName1) + 'WITH(FIELDTERMINATOR =' + quotedstr(Delim) + ',ROWTERMINATOR =''\n'')';
  try
    ADOConnection1.Execute(isql);
  except
    on e: Exception do
    begin
      mmoCustFile.Lines.Add(Format('�����ļ���%s��ʧ�ܣ��쳣��Ϣ%s', [FileA, e.Message]));
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
      for sheetIndex := 0 to 99 do //���ɵ���100��Sheet
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
          Break;  //�����쳣
        end;
      end;
          //fileName := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '(������ļ�).txt';
      fileName := ExtractFilePath(FileB) + tableName2 + '.txt';
      list.SaveToFile(fileName);
    end;
  finally
    FreeAndNil(List);
  end;

  //��txt�ļ����浽���ݿ�
  if FileExists(fileName) = False then
  begin
    mmoCustFile.Lines.Add('����' + FileA + '��' + FileB + 'ʧ��');
    exit;
  end;

  isql := 'drop table [' + tableName2 + ']';
  try
    ADOConnection1.Execute(isql);
  except
  end;
  try
    isql := 'create table [' + tableName2 + '](�������� varchar(200),��֤�� varchar(100));';
    isql := isql + 'BULK INSERT [' + tableName2 + '] FROM ' + QuotedStr(fileName) + 'WITH(FIELDTERMINATOR =' + quotedstr(Delim) + ',ROWTERMINATOR =''\n'')';
    try
      ADOConnection1.Execute(isql);
    except
      on e: Exception do
      begin
        mmoCustFile.Lines.Add(Format('�����ļ���%s��ʧ�ܣ��쳣��Ϣ%s', [FileB, e.Message]));
        exit;
      end;
    end;
  finally
    if FileExists(fileName) then
      DeleteFile(fileName);
  end;

  fileNameout := ExtractFilePath(FileA) + Trim(ChangeFileExt(ExtractFileName(FileA), '')) + '(δƥ���¼).txt';
  isql := Format('SELECT A.* FROM [%s] AS A LEFT JOIN [%s] AS B ON A.��������=B.�������� AND A.��֤��=B.��֤�� WHERE B.�������� IS NULL AND B.��֤��  IS NULL', [tableName1, tableName2]);
  isql := 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr(isql) + ',' + QuotedStr(fileNameout) + ';';
  try
    ADOConnection1.Execute(isql);
  except
    on e: Exception do
    begin
      mmoCustFile.Lines.Add(Format('�Ա�' + FileA + '��' + FileB + '�ļ�ʧ�ܣ��쳣��Ϣ%s', [e.Message]));
      exit;
    end;
  end;
  if FileSizeByName(fileNameout) > 0 then
    mmoCustFile.Lines.Add(Format(FileA + '��' + FileB + '���ڲ�ƥ��ļ�¼���뵽��%s���鿴', [fileNameout]))
  else
    DeleteFile(fileNameout);
  mmoCustFile.Lines.Add('ƥ��' + FileA + '��' + FileB + '���');
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
    mmoLog2.Lines.Add('���ܺķѽ϶�ʱ�䣬����ƥ����...');
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
        mmoLog2.Lines.Add(Format('���㵽��%d���%d����������', [zu, yu]));
    end;
    mmoLog2.Lines.Add('��ϲ����ȫ��ƥ����ɣ�����ʱ��' + IntToStr(SecondsBetween(sDateTime, Now)) + '�룡');
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
    ShowMessage('δѡ����ʱ�������ƶ���');
    Exit;
  end
  else if FCheckListBoxSelectCountA > 1 then
  begin
    ShowMessage('ͬʱѡ�ж���ʱ�������ƶ���');
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
    ShowMessage('δѡ����ʱ�������ƶ���');
    Exit;
  end
  else if FCheckListBoxSelectCountB > 1 then
  begin
    ShowMessage('ͬʱѡ�ж���ʱ�������ƶ���');
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
  strCaption := '��ֱ��ѡ���ļ��С�';
  //�ò���������ļ��д��ڵ���ʾ˵������
  wstrRoot := Trim(edtPath.Text);
  //���������ʾ����ʾ������ļ��д����еĸ�Ŀ¼��Ĭ�ϻ�ձ�ʾ���ҵĵ��ԡ���
  SelectDirectory(strCaption, wstrRoot, strDirectory);
  lbl7.Caption := strDirectory;
  //���ݽ�������в���strDirectory��ʾ�����ķ���ֵ
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
    lblFileCount.Caption := '�ļ�������' + Inttostr(fileList.Count);
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
  OpenDialog.FileName := '�ͻ��ļ�';
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
          //���������ͬ���ļ�
          DisableControls;
          try
            First;
            while not Eof do
            begin
              if Trim(FieldByName('fileName').AsString) = Trim(OpenDialog.Files[i]) then//if Trim(ChangeFileExt(ExtractFileName(FieldByName('fileName').AsString), '')) = Trim(ChangeFileExt(ExtractFileName(OpenDialog.FileName), '')) then
              begin
                ShowMessage('�Ѿ������ļ���' + Trim(FieldByName('fileName').AsString) + '��!!!'); //ShowMessage('�Ѿ������ļ���' + Trim(ChangeFileExt(ExtractFileName(FieldByName('fileName').AsString), '')) + '��!!!');
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
  OpenDialogA.FileName := '�ı��ļ�';
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
        mmoLog.Lines.Add('A�ļ��ܼ�¼[' + FieldByName('count').AsString + ']��');
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
  OpenDialogB.FileName := '�ı��ļ�';
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
        mmoLog.Lines.Add('B�ļ��ܼ�¼[' + FieldByName('count').AsString + ']��');
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
    ShowMessage('�����ݺŲ���Ϊ��!!!');
    Exit;
  end;
  sDate := Now;
  ssDate := Now;
  Memo1.Lines.Clear;
  //��Ȿ�����ݺ��µĶ�汾���ݱ���ظ�����
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
      //�����������ݺ��µĶ�汾���ݱ�
      first;
      while not Eof do
      begin
        iTable := FieldByName('name').AsString;
        isql := 'select sjm from ' + iTable + ' where ppdm+yscdm+Ny+sjm in (select ppdm+yscdm+Ny+sjm from ' + Trim(iTable) + ' group by ppdm,yscdm,Ny,sjm having count(*) > 1) and id not in (select min(id) from ' + Trim(iTable) + ' group by ppdm,yscdm,Ny,sjm having count(ppdm+yscdm+Ny+sjm)>1)';

        //��Ȿ���ظ������
        with iQuery2 do
        begin
          CommandTimeout := 0;
          Connection := ADOConnection1;
          Close;
          SQL.Text := isql;
          Open;
          if not IsEmpty then
          begin
            Memo1.Lines.Add('������' + iTable + '����⵽�ظ�����룺');
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
                Memo1.Lines.Add('ɾ�������ظ�����ʧ�ܣ��쳣��Ϣ��' + e.Message);
                Exit;
              end;
            end;
            eDate := Now;
            Memo1.Lines.Add('ɾ�������ظ����ݳɹ�����ʱ��' + IntToStr(SecondsBetween(eDate, sDate)) + '��!!!');
          end
          else
            Memo1.Lines.Add('������' + iTable + '��δ��⵽�ظ������');
        end;
        iQuery.Next;
      end;
    end;
    //�����ʷ���ظ�����
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
          Memo1.Lines.Add('�����ʷ�⡾' + ADOQuery1.fieldbyname('name').AsString + '��...');
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
              Memo1.Lines.Add('��⵽��ʷ�⡾' + ADOQuery1.fieldbyname('name').AsString + '�����ظ����ݣ�');
              First;
              while not Eof do
              begin
                Memo1.Lines.Add('Ʒ�ƴ��롾' + FieldByName('ppdm').AsString + '����ӡˢ�����롾' + FieldByName('yscdm').AsString + '�������¡�' + FieldByName('ny').AsString + '��������롾' + FieldByName('sjm').AsString + '��');
                Next;
              end;
              Memo1.Lines.Add('ɾ����ʷ�⡾' + ADOQuery1.fieldbyname('name').AsString + '�����ظ�����......');
              isql := 'delete from ' + trim(iTable) + ' WHERE EXISTS(SELECT 1 from ' + ADOQuery1.fieldByName('name').AsString + ' WHERE ' + ADOQuery1.fieldByName('name').AsString + '.ppdm+' + ADOQuery1.fieldByName('name').AsString + '.yscdm+' + ADOQuery1.fieldByName('name').AsString + '.NY+' + ADOQuery1.fieldByName('name').AsString + '.sjm=' + trim(iTable) + '.ppdm+' + trim(iTable) + '.yscdm+' + trim(iTable) + '.NY+' + trim(iTable) + '.sjm)';
              try
                ADOConnection1.Execute(isql);
              except
                on e: Exception do
                begin
                  Memo1.Lines.Add('ɾ����ʷ�⡾' + ADOQuery1.fieldbyname('name').AsString + '�����ظ�����ʧ�ܣ��쳣��Ϣ��' + e.Message);
                  Exit;
                end;
              end;
              eDate := Now;
              Memo1.Lines.Add('ɾ����ʷ�⡾' + ADOQuery1.fieldbyname('name').AsString + '�����ظ����ݳɹ�����ʱ��' + IntToStr(SecondsBetween(eDate, sDate)) + '��!!!');
            end
            else
              Memo1.Lines.Add('�����ʷ�⡾' + ADOQuery1.fieldbyname('name').AsString + '�����ظ�����!');
          end;
          Next;
        end;
      end
      else
        Memo1.Lines.Add('û�м�⵽��ʷ��');
    end;
  finally
    iQuery.Free;
    iQuery2.Free;
    iQuery3.Free;
  end;
  eDate := Now;
  Memo1.Lines.Add('�����ɣ���ʱ��' + IntToStr(SecondsBetween(eDate, ssDate)) + '��!!!');
end;

procedure TfrmQRCode.btnNotRepeatMultiClick(Sender: TObject);
var
  fileName,aa: string;
  sDate, eDate: TDateTime;
begin
  if not FIsImported then
  begin
    mmoMulti.Lines.Add('û��Ҫ��������ݣ����ȵ�����ٴ���!!!');
    Exit;
  end;
  if Trim(lbl7.Caption) = '' then
    fileName := Trim(edtPath.Text) + '\' + FormatDateTime('yyyymmddhh', Now) + '�����ظ���.txt'
  else
    fileName := Trim(lbl7.Caption) + '\' + FormatDateTime('yyyymmddhh', Now) + '�����ظ���.txt';
  try
    sDate := now;
    aa:='exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('select DISTINCT code from ##tb_Multi') + ',' + QuotedStr(fileName) + ';';
    //ShowMessage(aa);
    ADOConnection1.Execute('exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('select DISTINCT code from ##tb_Multi') + ',' + QuotedStr(fileName) + ';');
  except
    on e: Exception do
    begin
      mmoMulti.Lines.Add('�����ظ�ʧ�ܣ��쳣��Ϣ��' + E.Message);
      Exit;
    end;
  end;
  eDate := Now;
  mmoMulti.Lines.Add('�����ظ��ɹ��������ļ���' + fileName + '��,��ʱ��' + IntToStr(SecondsBetween(eDate, sDate)) + '��!!!');
end;

procedure TfrmQRCode.btnPublicClick(Sender: TObject);
var
  datePointer: TDateTime;
  iQuery: TADOQuery;
  filename: string;
begin
  if Trim(edtA.Text) = '' then
    MessageBox(Self.Handle, 'A�ļ�����Ϊ�գ����飡', '��ʾ', MB_OK)
  else if Trim(edtB.Text) = '' then
    MessageBox(Self.Handle, 'B�ļ�����Ϊ�գ����飡', '��ʾ', MB_OK)
  else if ExtractFileExt(edtA.Text) <> '.txt' then
    MessageBox(Self.Handle, '�򿪵�A�ļ����Ͳ��ԣ����飡', '��ʾ', MB_OK)
  else if ExtractFileExt(edtB.Text) <> '.txt' then
    MessageBox(Self.Handle, '�򿪵�B�ļ����Ͳ��ԣ����飡', '��ʾ', MB_OK)
  else if not FileExists(edtA.Text) then
    MessageBox(Self.Handle, 'A�ļ������ڣ����飡', '��ʾ', MB_OK)
  else if not FileExists(edtB.Text) then
    MessageBox(Self.Handle, '���ļ������ڣ����飡', '��ʾ', MB_OK)
  else
  begin
    btnPublic.Enabled := not btnPublic.Enabled;
    mmoLog.Lines.Add('���ܺķѽ϶�ʱ�䣬����ƥ����...');
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
        mmoLog.Lines.Add('ƥ��ʧ�ܣ����Ժ����ԣ�');
        Exit;
      end;
      with iQuery do
      begin
        CommandTimeout := 0;
        Connection := ADOConnection1;
        Close;
        SQL.Text := 'select count(1) as count from tableC';
        Open;
        mmoLog.Lines.Add('ƥ����ɣ���ƥ�䣺' + FieldByName('count').AsString + '����¼������ʱ��' + IntToStr(SecondsBetween(Now, datePointer)) + '�룡');
        Close;
        filename := Trim(edtPath.Text) + '\' + Trim(btnPublic.Caption) + '.txt';
        sql.Text := 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('SELECT * from DB_QRCode..tableC') + ',' + QuotedStr(filename);
        try
          ExecSQL;
        except
          on e: Exception do
          begin
            mmoLog.Lines.Add('�ļ�����ʧ�ܣ��쳣��Ϣ��' + E.Message);
            Exit;
          end;
        end;
        mmoLog.Lines.Add('�ļ����ɳɹ����ļ�����' + filename);
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
    MessageBox(Self.Handle, 'A�ļ�����Ϊ�գ����飡', '��ʾ', MB_OK)
  else if Trim(edtB.Text) = '' then
    MessageBox(Self.Handle, 'B�ļ�����Ϊ�գ����飡', '��ʾ', MB_OK)
  else if ExtractFileExt(edtA.Text) <> '.txt' then
    MessageBox(Self.Handle, '�򿪵�A�ļ����Ͳ��ԣ����飡', '��ʾ', MB_OK)
  else if ExtractFileExt(edtB.Text) <> '.txt' then
    MessageBox(Self.Handle, '�򿪵�B�ļ����Ͳ��ԣ����飡', '��ʾ', MB_OK)
  else if not FileExists(edtA.Text) then
    MessageBox(Self.Handle, 'A�ļ������ڣ����飡', '��ʾ', MB_OK)
  else if not FileExists(edtB.Text) then
    MessageBox(Self.Handle, '���ļ������ڣ����飡', '��ʾ', MB_OK)
  else
  begin
    btnRepeat.Enabled := not btnRepeat.Enabled;
    mmoLog.Lines.Add('���ܺķѽ϶�ʱ�䣬����ƥ����...');
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
        mmoLog.Lines.Add('ƥ��ʧ�ܣ����Ժ����ԣ�');
        Exit;
      end;
      with iQuery do
      begin
        Connection := ADOConnection1;
        Close;
        SQL.Text := 'select count(1) as count from tableC';
        Open;
        mmoLog.Lines.Add('ƥ����ɣ���ƥ�䣺' + FieldByName('count').AsString + '����¼������ʱ��' + IntToStr(SecondsBetween(Now, datePointer)) + '�룡');
        Close;
        filename := Trim(edtPath.Text) + '\' + Trim(btnRepeat.Caption) + '.txt';
        sql.Text := 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('SELECT * from DB_QRCode..tableC') + ',' + QuotedStr(filename);
        try
          ExecSQL;
        except
          on e: Exception do
          begin
            mmoLog.Lines.Add('�ļ�����ʧ�ܣ��쳣��Ϣ��' + E.Message);
            Exit;
          end;
        end;
        mmoLog.Lines.Add('�ļ����ɳɹ����ļ�����' + filename);
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
    mmoMulti.Lines.Add('û��Ҫ��������ݣ����ȵ�����ٴ���!!!');
    Exit;
  end;
  if Trim(lbl7.Caption) = '' then
    fileName := Trim(edtPath.Text) + '\' + FormatDateTime('yyyymmddhh', Now) + '���ظ���.txt'
  else
    fileName := Trim(lbl7.Caption) + '\' + FormatDateTime('yyyymmddhh', Now) + '���ظ���.txt';
  try
    sDate := Now;
    ADOConnection1.Execute('exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('select code from ##tb_Multi GROUP BY code HAVING COUNT(code)>1') + ',' + QuotedStr(fileName) + ';');
  except
    on e: Exception do
    begin
      mmoMulti.Lines.Add('�����ظ�ʧ�ܣ��쳣��Ϣ��' + E.Message);
      Exit;
    end;
  end;
  eDate := Now;
  mmoMulti.Lines.Add('�����ظ��ɹ��������ļ���' + fileName + '��,��ʱ��' + IntToStr(SecondsBetween(eDate, sDate)) + '��!!!');
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
    ShowMessage('�����ݺŲ���Ϊ��!!!');
    Exit;
  end;
  if Trim(Copy(edtNo.Text, 1, 2)) <> 'QR' then
  begin
    ShowMessage('�����ݺű�����QR��ͷ!!!');
    Exit;
  end;
  if Trim(edtType.Text) = '' then
  begin
    ShowMessage('����/С�в���Ϊ��!!!');
    Exit;
  end;
  if Trim(edtMode.Text) = '' then
  begin
    ShowMessage('�������Ͳ���Ϊ��!!!');
    Exit;
  end;
  if Trim(edtWebAddr.Text) = '' then
  begin
    ShowMessage('��ַ����Ϊ��!!!');
    Exit;
  end;
  if Trim(edtBrandCode.Text) = '' then
  begin
    ShowMessage('Ʒ�ƴ��벻��Ϊ��!!!');
    Exit;
  end;
  if Trim(edtPrinterCode.Text) = '' then
  begin
    ShowMessage('ӡˢ�����벻��Ϊ��!!!');
    Exit;
  end;
  if Trim(edtNy.Text) = '' then
  begin
    ShowMessage('���²���Ϊ��!!!');
    Exit;
  end;
  if Length(edtNy.Text) <> 4 then
  begin
    ShowMessage('����λ�����ԣ����Ҹ�ʽ������YYMM!!!');
    Exit;
  end;
  if (Trim(edtCount.Text) = '') or (Trim(edtCount.Text) = '0') then
  begin
    ShowMessage('��ά����������Ϊ�ջ�0!!!');
    Exit;
  end;
  if (Trim(edtThreadCount.Text) = '') or (Trim(edtThreadCount.Text) = '0') then
  begin
    ShowMessage('�����߳�������Ϊ�ջ�0!!!');
    Exit;
  end;
  if (Trim(edtMax.Text) = '') or (Trim(edtMax.Text) = '0') then
  begin
    ShowMessage('�������ݰ����޲���Ϊ�ջ�0!!!');
    Exit;
  end;

  if not ShowAsk(0, PChar('��ȷ�����������Ϣ����ȷ��ȷ��Ҫ���ɶ�ά����')) then
    Exit;

  with ADOQuery1 do
  begin
    Close;
    SQL.Text := 'select 1 from sysobjects where name like' + QuotedStr(Trim(edtNo.Text) + '%') + ' and xtype=''U''';
    Open;
    if not IsEmpty then
    begin
      if not ShowAsk(0, PChar('�˵��Ķ�ά���Ѿ����ɹ�����ȷ����Ҫ����������')) then
        Exit;
    end;
  end;

  lblResult.Caption := '�����У����Ժ�......';
  //����
//  EnterCriticalSection(Cs);
  //��UIȡֵ
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
  //�ͷ���
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
          //�����µ�ADOConnection�����߳����ò�������ִ��
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
              frmQRCode.lblResult.Caption := '��ά�����ɳɹ�����ʱ��' + IntToStr(SecondsBetween(eDate, sDate)) + '�롣Ϊ�˱����ά�����ʷ���е��ظ��������ƥ����ʷ���ά��!!!'
            else
              frmQRCode.lblResult.Caption := '��ά������ʧ�ܣ��쳣��Ϣ��' + err;
            frmQRCode.lblResult.Refresh;
          end);
      end).Start;
  except
    on E: Exception do
    begin
      lblResult.Caption := '��ά������ʧ�ܣ��쳣��Ϣ:' + e.Message;
      Exit;
    end;
  end;
end;

procedure TfrmQRCode.btnTxtClick(Sender: TObject);
var
  path: string;
begin
  OpenDialogA.FileName := '�ı��ļ�';
  OpenDialogA.Filter := '*.txt|*.txt';
  if OpenDialogA.Execute then
  begin
    path := Trim(OpenDialogA.Files.Text);
    //path:=StringReplace(path,#13#10,'',[rfReplaceAll,rfIgnoreCase]);//�ļ������лس����ţ�����sql��䱨��
    edtTxt.Text := path;
  end;
end;

procedure TfrmQRCode.btnUpAClick(Sender: TObject);
var
  i: Integer;
begin
  if FCheckListBoxSelectCountA = 0 then
  begin
    ShowMessage('δѡ����ʱ�������ƶ���');
    Exit;
  end
  else if FCheckListBoxSelectCountA > 1 then
  begin
    ShowMessage('ͬʱѡ�ж���ʱ�������ƶ���');
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
    ShowMessage('δѡ����ʱ�������ƶ���');
    Exit;
  end
  else if FCheckListBoxSelectCountB > 1 then
  begin
    ShowMessage('ͬʱѡ�ж���ʱ�������ƶ���');
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
        Result := Format('��������%s��ʧ�ܣ�ʧ��ԭ��%s', [fileNameout, e.Message]);
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
    XLS2.Sheets[0].AsString[0, 0] := '���';
    for i := 0 to adoCheck3.FieldCount - 1 do
    begin
      XLS2.Sheets[0].AsString[i + 1, 0] := adoCheck3.Fields[i].DisplayName;
    end;
    MaxY := adoCheck3.RecordCount;
    MsgStr := Format('����ת���ļ�����%s��', [FileNameout]);
    //mmoCustFile.Lines.Add(MsgStr);

    y := 1;
    adoCheck3.First;
    for I := 1 to MaxY do
    begin
      XLS2.Sheets[0].AsString[0, i] := IntToStr(i);
      XLS2.Sheets[0].AsString[1, i] := adoCheck3.FieldByName('��������').AsString;
      XLS2.Sheets[0].AsString[2, i] := adoCheck3.FieldByName('��֤��').AsString;

      adoCheck3.Next;
    end;
    for i := 0 to adoCheck3.FieldCount - 1 do
    begin
      XLS2.Sheets[0].AutoWidthCol(i + 1);
    end;
    XLS2.SaveToFile(fileNameout);
    MsgStr := Format('ת���ļ�����%s�����', [FileNameout]);
    //mmoCustFile.Lines.Add(MsgStr);
  except
    on E: Exception do
    begin
      MsgStr := Format('��������%s��ʧ�ܣ�ʧ��ԭ��%s', [fileNameout, e.Message]);
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
  mmoCustFile.Lines.Add('��ʼƥ��' + FileA + '��' + FileB + '��ȴ�...');

  fileName1 := FileA;
  tableName1 := '##' + Trim(ChangeFileExt(ExtractFileName(fileName1), ''));
    //��txt�ļ����浽���ݿ�
  isql := 'drop table [' + tableName1 + ']';
  try
    ADOConnection1.Execute(isql);
  except
  end;
  isql := 'create table [' + tableName1 + '](�������� varchar(200),��֤�� varchar(100));';
  isql := isql + 'BULK INSERT [' + tableName1 + '] FROM ' + QuotedStr(fileName1) + 'WITH(FIELDTERMINATOR =' + quotedstr(Delim) + ',ROWTERMINATOR =''\n'')';
  try
    ADOConnection1.Execute(isql);
  except
    on e: Exception do
    begin
      mmoCustFile.Lines.Add(Format('�����ļ���%s��ʧ�ܣ��쳣��Ϣ%s', [FileA, e.Message]));
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

      for sheetIndex := 0 to 99 do //���ɵ���100��Sheet
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
          Break;  //�����쳣
        end;
      end;
          //fileName := ExtractFilePath(qryCustFile.FieldByName('fileName').AsString) + ChangeFileExt(ExtractFileName(XLS.Filename), '') + '(������ļ�).txt';
      fileName := ExtractFilePath(FileB) + tableName2 + '.txt';
      list.SaveToFile(fileName);
    end;
  finally
    FreeAndNil(List);
  end;

  //��txt�ļ����浽���ݿ�
  if FileExists(fileName) = False then
  begin
    mmoCustFile.Lines.Add('����' + FileA + '��' + FileB + 'ʧ��');
    exit;
  end;

  isql := 'drop table [' + tableName2 + ']';
  try
    ADOConnection1.Execute(isql);
  except
  end;
  try
    isql := 'create table [' + tableName2 + '](�������� varchar(200),��֤�� varchar(100));';
    isql := isql + 'BULK INSERT [' + tableName2 + '] FROM ' + QuotedStr(fileName) + 'WITH(FIELDTERMINATOR =' + quotedstr(Delim) + ',ROWTERMINATOR =''\n'')';
    try
      ADOConnection1.Execute(isql);
    except
      on e: Exception do
      begin
        mmoCustFile.Lines.Add(Format('�����ļ���%s��ʧ�ܣ��쳣��Ϣ%s', [FileB, e.Message]));
        exit;
      end;
    end;
  finally
    if FileExists(fileName) then
      DeleteFile(fileName);
  end;

  isql := Format('UPDATE [%s] SET ��֤��=B.��֤�� FROM [%s] A   INNER JOIN [%s] AS B ON A.��������=B.�������� WHERE ISNULL(A.��֤��,'''')='''' ', [tableName1, tableName1, tableName2]);
  try
    ADOConnection1.Execute(isql);
  except
    on e: Exception do
    begin
      mmoCustFile.Lines.Add(Format(FileA + '��' + FileB + '������֤��ʧ�ܣ��쳣��Ϣ%s', [e.Message]));
      exit;
    end;
  end;

  fileNamein := ExtractFilePath(FileB) + Trim(ChangeFileExt(ExtractFileName(FileB), '')) + '(δƥ���¼).txt';
  fileNameout := ExtractFilePath(FileB) + Trim(ChangeFileExt(ExtractFileName(FileB), '')) + '(δƥ���¼).xlsx';
  isql := Format('SELECT * FROM [%s] AS A LEFT JOIN [%s] AS B ON A.��������=B.�������� AND A.��֤��=B.��֤�� WHERE B.�������� IS NULL AND B.��֤��  IS NULL', [tableName2, tableName1]);
  //isql := 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr(isql) + ',' + QuotedStr(fileNamein) + ';';
  try
    adoCheck3.SQL.Text := isql;
    adoCheck3.Open;
    DataToExcel(fileNamein, fileNameout);
    //DbGridEhToExcel(fileNameout);

  except
    on e: Exception do
    begin
      mmoCustFile.Lines.Add(Format('�Ա�' + FileA + '��' + FileB + '�ļ�ʧ�ܣ��쳣��Ϣ%s', [e.Message]));
      if FileExists(fileNameout) then
        DeleteFile(fileNameout);
      exit;
    end;
  end;
  if FileSizeByName(fileNameout) > 0 then
  begin
    mmoCustFile.Lines.Add(Format(FileA + '��' + FileB + '���ڲ�ƥ��ļ�¼���뵽��%s���鿴', [fileNameout]));
  end
  else
    DeleteFile(fileNameout);

  fileNamein := ExtractFilePath(FileB) + Trim(ChangeFileExt(ExtractFileName(FileB), '')) + '(ƥ��ɹ���¼).txt';
  fileNameout := ExtractFilePath(FileB) + Trim(ChangeFileExt(ExtractFileName(FileB), '')) + '(ƥ��ɹ���¼).xlsx';
  isql := Format('SELECT A.* FROM [%s] AS A INNER JOIN [%s] AS B ON A.��������=B.�������� AND A.��֤��=B.��֤��', [tableName1, tableName2]);
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
      mmoCustFile.Lines.Add(Format('�Ա�' + FileA + '��' + FileB + '�ļ�ʧ�ܣ��쳣��Ϣ%s', [e.Message]));
      if FileExists(fileNameout) then
        DeleteFile(fileNameout);
      exit;
    end;
  end;
  if FileSizeByName(fileNameout) > 0 then
    mmoCustFile.Lines.Add(Format(FileA + '��' + FileB + '����ƥ��ɹ��ļ�¼���뵽��%s���鿴', [fileNameout]))
  else
    DeleteFile(fileNameout);

  mmoCustFile.Lines.Add('ƥ��' + FileA + '��' + FileB + '���');
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
    ShowMessage('û��ѡ��Ʒ���ļ�');
    Exit;
  end;
  if not qryCustFile.Active then
  begin
    ShowMessage('û��ѡ��Ҫ������ļ�');
    Exit;
  end;
  if qryCustFile.IsEmpty then
  begin
    ShowMessage('û��ѡ��Ҫ������ļ�');
    Exit;
  end;
  mmoCustFile.Lines.Clear;
  mmoCustFile.Lines.Add('ƥ�������Ҫ�����ӣ������ĵȴ�...');
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
    mmoCustFile.Lines.Add('��ϲ����ȫ��ƥ����ɣ�����ʱ��' + IntToStr(SecondsBetween(sDateTime, Now)) + '�룡');
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
    MessageBox(Self.Handle, '���ļ�����Ϊ�գ����飡', '��ʾ', MB_OK);
    Exit;
  end
  else if not FileExists(edtCheckA.Text) then
  begin
    MessageBox(Self.Handle, '���ļ������ڣ����飡', '��ʾ', MB_OK);
    Exit;
  end
  else if not chkAuto.Checked then
  begin
    if Trim(edtCheckB.Text) = '' then
    begin
      MessageBox(Self.Handle, '���Զ�ƥ�丸�ļ�ģʽ�£����ļ�����Ϊ�գ����飡', '��ʾ', MB_OK);
      Exit;
    end
    else if not FileExists(edtCheckB.Text) then
    begin
      MessageBox(Self.Handle, '���Զ�ƥ�丸�ļ�ģʽ�£����ļ������ڣ����飡', '��ʾ', MB_OK);
      Exit;
    end
  end
  else if chkAuto.Checked then
  begin
    if Trim(edtCustNy.Text) = '' then
    begin
      MessageBox(Self.Handle, '�Զ�ƥ�丸�ļ�ģʽ�£�����ͻ��ļ����²���Ϊ�գ����飡', '��ʾ', MB_OK);
      Exit;
    end;
  end;
  if not ShowAsk(0, PChar('��ȷ�϶Աȵ��ļ������ɵ��ļ���ʽ��ȷ������')) then
    Exit;
  btnCheck.Enabled := not btnCheck.Enabled;
  mmoLog1.Lines.Add('���ܺķѽ϶�ʱ�䣬����ƥ����...');
  datePointer := Now;

  try
    iQuery := TADOQuery.Create(Self);
    try
      //���Զ�ƥ�丸�ļ�ģʽ��ֻ��һ������
      if not chkAuto.Checked then
      begin
        with iQuery do
        begin
          CommandTimeout := 0;
          Connection := ADOConnection1;
          Close;
          SQL.Clear;
          SQL.Add('TRUNCATE TABLE tb_middle;' + 'INSERT INTO tb_middle(code,number) SELECT ��������,��֤�� FROM tb_father WHERE exists(SELECT �������� FROM tb_child WHERE tb_child.��������=tb_father.��������) order by tb_father.��������;');
          SQL.Add('TRUNCATE TABLE tb_middle_Not;' + 'INSERT INTO tb_middle_Not(code) SELECT �������� FROM tb_child WHERE not exists(SELECT �������� FROM tb_father WHERE tb_child.��������=tb_father.��������) order by tb_child.��������;');
          ExecSQL;
        end;
        if chk2.Checked then//����txt�ļ�,�ļ��������ļ�������
        begin
          fileName := ExtractFilePath(edtCheckB.Text) + Trim(ChangeFileExt(ExtractFileName(edtCheckB.Text), '')) + '��ƥ��ɹ����ļ���.txt';
          NotFileName := ExtractFilePath(edtCheckB.Text) + Trim(ChangeFileExt(ExtractFileName(edtCheckB.Text), '')) + '��δƥ��ɹ����ļ���.txt';
          isql := 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('SELECT code+'' ''+number FROM DB_QRCode..tb_middle') + ',' + QuotedStr(fileName) + ';';
          isql := isql + 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('SELECT code FROM DB_QRCode..tb_middle_Not') + ',' + QuotedStr(NotFileName) + ';';
          ADOConnection1.Execute(isql);
        end
        else if chk3.Checked then//����xlsx�ļ�,�ļ��������ļ�������
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
                mmoLog1.Lines.Add('û��ƥ��ɹ������ݣ�δ�����ļ�!!!');
                btnCheck.Enabled := True;
                Exit;
              end
              else
              begin
                fileName := ExtractFilePath(edtCheckB.Text) + Trim(ChangeFileExt(ExtractFileName(edtCheckB.Text), '')) + '��ƥ��ɹ����ļ���.xlsx';
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
                XLS.Sheets[0].AsString[0, Row] := '���';
                XLS.Sheets[0].AsString[1, Row] := '��ά��';
                XLS.Sheets[0].AsString[2, Row] := '��֤��';
                XLS.Sheets[0].AsString[3, Row] := '�����κ�';
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
            mmoLog1.Lines.Add('ƥ����ɣ������ļ���' + fileName + '������ƥ�䣺' + ADOQuery1.FieldByName('count').AsString + '����¼������ʱ��' + IntToStr(SecondsBetween(Now, datePointer)) + '�룡');
        end;

        if chk2.Checked then  //�����ı�-����δƥ��ɹ����ļ�
        begin
          with ADOQuery1 do
          begin
            Close;
            SQL.Text := 'select count(1) as count from tb_middle_Not';
            Open;
            if FieldByName('count').AsInteger > 0 then
              mmoLog1.Lines.Add('��δƥ��ɹ��Ķ�ά�룬�����ļ���' + notFileName + '����δƥ��ɹ��Ķ�ά�빲��' + ADOQuery1.FieldByName('count').AsString + '����¼������ʱ��' + IntToStr(SecondsBetween(Now, datePointer)) + '�룡');
          end;
        end;
      end
      //�Զ�ƥ�丸�ļ�ģʽ��ֻ�ж����������������
      else
      begin
        //�ҵ�ָ������
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
              SQL.Text := 'TRUNCATE TABLE tb_middle;' + 'INSERT INTO tb_middle(code,number) SELECT ��������,��֤�� FROM dbo.[' + tableName + '] WHERE exists(SELECT �������� FROM tb_child WHERE tb_child.��������=dbo.[' + tableName + '].��������) order by ��������;';
              ExecSQL;

              Close;
              SQL.Text := 'select * from tb_middle';
              Open;

              if not IsEmpty then
              begin
                mmoLog1.Lines.Add('ƥ�䵽������' + tableName + '������ƥ�䵽��' + inttostr(iQuery.RecordCount) + '������¼!');
                totalcon := totalcon + iQuery.RecordCount;
                if chk2.Checked then //����txt�ļ�,�ļ����Ը��ļ�������
                begin
                  NotFileName := ExtractFilePath(edtCheckA.Text) + Trim(ChangeFileExt(ExtractFileName(tableName), '')) + '��δƥ��ɹ����ļ���.txt';
                  fileName := ExtractFilePath(edtCheckA.Text) + Trim(ChangeFileExt(ExtractFileName(tableName), '')) + '��ƥ��ɹ����ļ���.txt';
                  isql := 'TRUNCATE TABLE tb_middle_not;' + 'INSERT INTO tb_middle_not(code) SELECT �������� FROM dbo.[' + tableName + '] WHERE not exists(SELECT �������� FROM tb_child WHERE tb_child.��������=dbo.[' + tableName + '].��������) order by ��������;';
                  isql := isql + 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('SELECT code+'' ''+number FROM DB_QRCode..tb_middle') + ',' + QuotedStr(fileName) + ';';
                  isql := isql + 'exec sp_sql_query_to_file ' + QuotedStr(hostIP.Text) + ',' + QuotedStr(DBUser.Text) + ',' + QuotedStr(DBPass.Text) + ',' + QuotedStr('SELECT code FROM DB_QRCode..tb_middle_Not') + ',' + QuotedStr(NotFileName) + ';';
                  ADOConnection1.Execute(isql);
                end
                else if chk3.Checked then  //����xlsx�ļ�,�ļ����Ը��ļ�������
                begin
                  //�����ݿ������д��xlsx
                  fileName := ExtractFilePath(edtCheckA.Text) + Trim(ChangeFileExt(ExtractFileName(tableName), '')) + '��ƥ��ɹ����ļ���.xlsx';
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
                  XLS.Sheets[0].AsString[0, Row] := '���';
                  XLS.Sheets[0].AsString[1, Row] := '��ά��';
                  XLS.Sheets[0].AsString[2, Row] := '��֤��';
                  XLS.Sheets[0].AsString[3, Row] := '�����κ�';
                  Row := 1;
                  num := 0;
                  DisableControls;
                  try
                    First;
                    while not Eof do
                    begin                                 //�������������
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
                mmoLog1.Lines.Add('�����ļ���' + fileName + '���ɹ�!!!');
                if chk2.Checked then  //�����ı�-����δƥ��ɹ����ļ�
                begin
                  with iQuery do
                  begin
                    Connection := ADOConnection1;
                    Close;
                    SQL.Text := 'select count(1) as count from tb_middle_Not';
                    Open;
                    if FieldByName('count').AsInteger > 0 then
                      mmoLog1.Lines.Add('��δƥ��ɹ��Ķ�ά�룬�����ļ���' + notFileName + '����δƥ��ɹ��Ķ�ά�빲��' + FieldByName('count').AsString + '����¼������ʱ��' + IntToStr(SecondsBetween(Now, datePointer)) + '�룡');
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
          mmoLog1.Lines.Add('δƥ�䵽�κ����ݣ�')
        else
        begin
          mmoLog1.Lines.Add('�Ӱ���ƥ�䵽��' + floattostr(totalcon) + '������¼!');
          mmoLog1.Lines.Add('��ϲ����ȫ��ƥ����ɣ�����ʱ��' + IntToStr(SecondsBetween(Now, datePointer)) + '�룡');
        end;
      end;
    finally
      iQuery.Free;
    end;
  except
    on e: Exception do
    begin
      mmoLog1.Lines.Add('ƥ��ʧ�ܣ��쳣��Ϣ:' + E.Message + '��');
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
    ShowMessage('�����ݺű�����QR��ͷ!!!');
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
        ShowMessage('�����ݺ������������飡');
        edtNo.SetFocus;
      end
      else
      begin
        if Trim(FieldByName('hexing').AsString) = '����' then
          edtType.ItemIndex := 0
        else if Trim(FieldByName('hexing').AsString) = 'С��' then
          edtType.ItemIndex := 1;
        if FieldByName('gcbh').AsString = '�㶫����' then
          edtMode.Text := '0';
        edtWebAddr.Text := Trim(FieldByName('wz').AsString);
        edtBrandCode.Text := Copy(Trim(FieldByName('ppdm').AsString), 1, Pos('��', FieldByName('ppdm').AsString) - 1);
        edtPrinterCode.Text := Copy(Trim(FieldByName('yscdm').AsString), 1, Pos('��', FieldByName('yscdm').AsString) - 1);
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
    ShowMessage('���������ѵ�������ϵ�ͷ���');
    Application.Terminate;
  end;

  InitializeCriticalSection(Cs);
  DragAcceptFiles(self.Handle, True); //����windows API
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
    DataSet.FieldByName('selected').DisplayLabel := 'ѡ��';
    DataSet.FieldByName('selected').DisplayWidth := 5;
    DataSet.FieldByName('fileName').DisplayLabel := '�ļ�';
    DataSet.FieldByName('fileName').DisplayWidth := 60;
    DataSet.FieldByName('isContainFirst').DisplayLabel := '���а���������';
    DataSet.FieldByName('isContainFirst').DisplayWidth := 15;
  end;
end;

procedure TfrmQRCode.qryMultiAfterOpen(DataSet: TDataSet);
begin
  if DataSet.Active then
  begin
    DataSet.FieldByName('selected').DisplayLabel := 'ѡ��';
    DataSet.FieldByName('selected').DisplayWidth := 5;
    DataSet.FieldByName('fileName').DisplayLabel := '�ļ�';
    DataSet.FieldByName('fileName').DisplayWidth := 80;
    DataSet.FieldByName('isContainFirst').DisplayLabel := '���а���������';
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
  dirs := TQueue.Create; //����Ŀ¼����
  dirs.Push(path); //����ʼ����·�����
  pszDir := dirs.Pop;
  curDir := StrPas(pszDir); //����
   {��ʼ����,ֱ������Ϊ��(��û��Ŀ¼��Ҫ����)}
  while (True) do
  begin
    //����������׺,�õ�����'c:\*.*' ��'c:\windows\*.*'������·��
    tmpStr := curDir + '\*.*';
    //�ڵ�ǰĿ¼���ҵ�һ���ļ�����Ŀ¼
    found := FindFirst(tmpStr, faAnyFile, searchRec);
    while found = 0 do //�ҵ���һ���ļ���Ŀ¼��
    begin
      //����ҵ����Ǹ�Ŀ¼
      if (searchRec.Attr and faDirectory) <> 0 then
      begin
        {�������Ǹ�Ŀ¼(C:\��D:\)�µ���Ŀ¼ʱ�����'.','..'��"����Ŀ¼"
         ����Ǳ�ʾ�ϲ�Ŀ¼���²�Ŀ¼�ɡ�����Ҫ���˵��ſ���}
        if (searchRec.Name <> '.') and (searchRec.Name <> '..') then
        begin
           {���ڲ��ҵ�����Ŀ¼ֻ�и�Ŀ¼��������Ҫ�����ϲ�Ŀ¼��·��
            searchRec.Name = 'Windows';
            tmpStr:='c:\Windows';
            �Ӹ��ϵ��һ�������
           }
          tmpStr := curDir + '\' + searchRec.Name;
               {����������Ŀ¼��ӡ����������š�
                ��ΪTQueue���������ֻ����ָ��,����Ҫ��stringת��ΪPChar
                ͬʱʹ��StrNew������������һ���ռ�������ݣ������ʹ�Ѿ���
                ����е�ָ��ָ�򲻴��ڻ���ȷ������(tmpStr�Ǿֲ�����)��}
          dirs.Push(StrNew(PChar(tmpStr)));
        end;
      end
      else //����ҵ����Ǹ��ļ�
      begin
         {Result��¼�����������ļ���������������CreateThread�����߳�
          �����ú����ģ���֪����ô�õ��������ֵ�������Ҳ�����ȫ�ֱ���}
        //���ҵ����ļ��ӵ�Memo�ؼ�
        if fileExt = '.*' then
          fileList.Add(curDir + '\' + searchRec.Name)
        else
        begin
          if SameText(RightStr(curDir + '\' + searchRec.Name, Length(fileExt)), fileExt) then
            fileList.Add(curDir + '\' + searchRec.Name);
        end;
      end;
      //������һ���ļ���Ŀ¼
      found := FindNext(searchRec);
    end;
    {��ǰĿ¼�ҵ������������û�����ݣ����ʾȫ���ҵ��ˣ�
      ������ǻ�����Ŀ¼δ���ң�ȡһ�������������ҡ�}
    if dirs.Count > 0 then
    begin
      pszDir := dirs.Pop;
      curDir := StrPas(pszDir);
      StrDispose(pszDir);
    end
    else
      break;
  end;
  //�ͷ���Դ
  dirs.Free;
  FindClose(searchRec);
end;

function TfrmQRCode.ShowAsk(MHandle: Thandle; tText: Pchar): Boolean;
begin
  if MessageBox(MHandle, tText, 'ѯ����Ϣ', MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2) = IDYES then
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
        //���������ͬ���ļ�
        DisableControls;
        try
          First;
          while not Eof do
          begin
            if Trim(ChangeFileExt(ExtractFileName(FieldByName('fileName').AsString), '')) = Trim(ChangeFileExt(ExtractFileName(StrPas(P)), '')) then
            begin
              ShowMessage('�Ѿ������ļ���' + Trim(ChangeFileExt(ExtractFileName(FieldByName('fileName').AsString), '')) + '��!!!');
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

