unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, jpeg, ExtCtrls, ShellApi,

  // XLSReadWriteII5 used units.
  XLSUtils5, Xc12DataStylesheet5, XLSDrawing5, XLSSheetData5, XLSReadWriteII5,
  XLSCmdFormat5, XLSExport5, XLSExportHTML5;

type
  TForm1 = class(TForm)
    Image: TImage;
    btnClose: TButton;
    btnCreateSheet: TButton;
    XLS: TXLSReadWriteII5;
    btnExportHtml: TButton;
    XLSHTML: TXLSExportHTML5;
    btnSave: TButton;
    btnProtect: TButton;
    edPassword: TEdit;
    Label1: TLabel;
    dlgOpen: TOpenDialog;
    btnOpenExcel: TButton;
    Label2: TLabel;
    edFilename: TEdit;
    Button1: TButton;
    Label3: TLabel;
    edHTMLFilename: TEdit;
    Button2: TButton;
    btnFindReplace: TButton;
    Label4: TLabel;
    procedure btnCreateSheetClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnProtectClick(Sender: TObject);
    procedure btnExportHtmlClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnOpenExcelClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure btnFindReplaceClick(Sender: TObject);
  private
    procedure AddTextBox(const C1,R1,C2,R2: integer; const C2Offs: double; const FontSize: integer; const AText: AxUCString);
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.AddTextBox(const C1, R1, C2, R2: integer; const C2Offs: double; const FontSize: integer; const AText: AxUCString);
var
  TB: TXLSDrawingTextBox;
  Editor: TXLSDrawingEditorTextBox;
  EditorShape: TXLSDrawingEditorShape;
begin
  // Add a text box.
  TB := XLS[0].Drawing.TextBoxes.Add;

  // Left column
  TB.Col1 := C1;
  // Top row.
  TB.Row1 := R1;
  // Offset within the row, in percent of the row height. Default is zero.
  TB.Row1Offs := 0;
  // Right column
  TB.Col2 := C2;
  // Offset within the column, in percent of the column width. Default is zero.
  TB.Col2Offs := C2Offs;
  TB.Row2 := R2;

  // Create an editor for editing the text. This is not required if you just
  // want to set the text, you can then use: TB.PlainText := 'My text';
  Editor := XLS[0].Drawing.EditTextBox(TB);
  try
    // Set the default font size.
    // Font properies in text runs are inherited from the previous text run,
    // and the first text run uses the properties of the DefaultFont.
    Editor.Body.DefaultFont.Size := FontSize;

    // Set the text.
    Editor.Body.PlainText := AText;
  finally
    Editor.Free;
  end;

  // Create an editor for the drawing shape properties.
  // A drawing shape is the graphic object that the text is displayed in.
  EditorShape := XLS[0].Drawing.EditShape(TB);
  try
    // Make sure that there is a Line property. The line is the border line
    // around the text box.
    EditorShape.ShapeProperies.HasLine := True;
    // Set the line width, in points.
    EditorShape.ShapeProperies.Line.Width := 1.5;

    // Set the fill of the text box.
    EditorShape.ShapeProperies.Fill.FillType := xdftSolid;
    EditorShape.ShapeProperies.Fill.AsSolid.Color.AsRGB := $FFFFA0;
  finally
    EditorShape.Free;
  end;
end;

procedure TForm1.btnCreateSheetClick(Sender: TObject);
var
  C,R: integer;
begin
  XLS[0].Options := XLS[0].Options - [soGridlines];

  XLS[0].Columns[0].CharWidth := 1;
  XLS[0].Columns[1].CharWidth := 13;
  XLS[0].Columns[2].CharWidth := 30;
  XLS[0].Columns[3].CharWidth := 16;
  XLS[0].Columns[4].CharWidth := 14;
  XLS[0].Columns[5].CharWidth := 14;

  XLS[0].Rows[1].PointHeight := 107;
  XLS[0].Rows[15].PointHeight := 27;

  for R := 15 to 25 do
    XLS[0].Rows[R].PointHeight := 19.5;

  XLS.CmdFormat.BeginEdit(XLS[0]);

  XLS.CmdFormat.Clear;
  XLS.CmdFormat.Font.Color.ExcelSwatch(3,0);
  XLS.CmdFormat.Font.Size := 26;
  XLS.CmdFormat.Alignment.Horizontal := chaRight;
  XLS.CmdFormat.Alignment.Vertical := cvaCenter;
  XLS.CmdFormat.Apply(4,1,5,1);
  XLS[0].MergeCells(4,1,5,1);
  XLS[0].AsString[4,1] := 'INVOICE';

  AddTextBox(1,3,2,7,0.5,9,
             'Street address' + #13 +
             'City, ST 00000' + #13 +
             'Phone: (206) 555-1163' + #13 +
             'Fax: (206) 555-1164' + #13 +
             'someone@example.com');

  XLS.CmdFormat.Clear;
  XLS.CmdFormat.Font.Size := 9;
  XLS.CmdFormat.Font.Style := [xfsBold];
  XLS.CmdFormat.Alignment.Horizontal := chaRight;
  XLS.CmdFormat.Alignment.Indent := 1;
  XLS.CmdFormat.Apply(4,3,4,6);
  XLS[0].AsString[4,3] := 'Date';
  XLS[0].AsString[4,4] := 'Invoice #';
  XLS[0].AsString[4,5] := 'For:';

  XLS.CmdFormat.Clear;
  XLS.CmdFormat.Border.Style := cbsThin;
  XLS.CmdFormat.Border.Color.ExcelSwatch(3,2);
  XLS.CmdFormat.Border.Preset(cbspOutlineAndInside);
  XLS.CmdFormat.Apply(5,3,5,6);
  XLS[0].MergeCells(5,5,5,6);
  XLS[0].AsString[5,3] := '[Date]';
  XLS[0].AsString[5,4] := '[InvoiceNo]';

  XLS.CmdFormat.Clear;
  XLS.CmdFormat.Fill.BackgroundColor.ExcelSwatch(3,0);
  XLS.CmdFormat.Font.Color.RGB := $FFFFFF;
  XLS.CmdFormat.Font.Size := 10;
  XLS.CmdFormat.Font.Style := [xfsBold];
  XLS.CmdFormat.Apply(1,8,5,8);
  for R := 8 to 13 do
    XLS[0].MergeCells(1,R,5,R);
  XLS[0].AsString[1,8] := 'Bill to:';
  XLS[0].AsString[1,9] := '[Contact]';
  XLS[0].AsString[1,10] := '[Company]';
  XLS[0].AsString[1,11] := '[Street]';
  XLS[0].AsString[1,12] := '[City]';
  XLS[0].AsString[1,13] := '[Phone]';

  XLS.CmdFormat.Clear;
  XLS.CmdFormat.Border.Style := cbsThin;
  XLS.CmdFormat.Border.Color.ExcelSwatch(3,2);
  XLS.CmdFormat.Border.Preset(cbspOutline);
  XLS.CmdFormat.Apply(1,8,5,13);

  XLS.CmdFormat.Clear;
  XLS.CmdFormat.Alignment.Horizontal := chaCenter;
  XLS.CmdFormat.Fill.BackgroundColor.ExcelSwatch(3,0);
  XLS.CmdFormat.Font.Color.RGB := $FFFFFF;
  XLS.CmdFormat.Font.Size := 10;
  XLS.CmdFormat.Font.Style := [xfsBold];
  XLS.CmdFormat.Apply(1,15,5,15);
  XLS[0].AsString[1,15] := 'Quantity';
  XLS[0].AsString[2,15] := 'Description';
  XLS[0].AsString[3,15] := 'Unit price';
  XLS[0].AsString[4,15] := 'Amount';
  XLS[0].AsString[5,15] := 'Discount';

  for R := 16 to 24 do begin
    for C := 1 to 5 do begin
      XLS.CmdFormat.Clear;
      if not Odd(R) then
        XLS.CmdFormat.Fill.BackgroundColor.ExcelSwatch(3,1);
      XLS.CmdFormat.Border.Style := cbsThin;
      XLS.CmdFormat.Border.Color.ExcelSwatch(3,2);
      XLS.CmdFormat.Border.Side[cbsBottom] := True;
      case C of
        1,5: begin
          if C = 1 then begin
            XLS.CmdFormat.Border.Side[cbsLeft] := True;
            XLS[0].AsFloat[C,R] := 1;
          end
          else begin
            XLS.CmdFormat.Border.Side[cbsRight] := True;
            XLS.CmdFormat.Number.Format := '0%';
            XLS[0].AsFloat[C,R] := 0;
          end;
          XLS.CmdFormat.Alignment.Horizontal := chaCenter;
          XLS.CmdFormat.Alignment.Vertical := cvaCenter;
          XLS.CmdFormat.Font.Size := 10;
          XLS.CmdFormat.Apply(C,R,C,R);
        end;
        2: begin
          XLS.CmdFormat.Alignment.Horizontal := chaLeft;
          XLS.CmdFormat.Alignment.Vertical := cvaCenter;
          XLS.CmdFormat.Alignment.Indent := 1;
          XLS.CmdFormat.Font.Size := 10;
          XLS[0].AsString[C,R] := 'Item number ' + IntToStr(R - 15);
        end;
        3: begin
          XLS.CmdFormat.Alignment.Horizontal := chaLeft;
          XLS.CmdFormat.Alignment.Vertical := cvaCenter;
          XLS.CmdFormat.Font.Size := 9;
          XLS.CmdFormat.Number.Format := '_("$"* # ##0.00_);_("$"* (# ##0.00);_("$"* "-"??_);_(@_)';
          XLS[0].AsFloat[C,R] := 125;
        end;
        4: begin
          XLS.CmdFormat.Alignment.Horizontal := chaLeft;
          XLS.CmdFormat.Alignment.Vertical := cvaCenter;
          XLS.CmdFormat.Font.Size := 9;
          XLS.CmdFormat.Number.Format := '_("$"* # ##0.00_);_("$"* (# ##0.00);_("$"* "-"??_);_(@_)';
          XLS[0].AsFormula[C,R] := 'B' + IntToStr(R + 1) + '*D' + IntToStr(R + 1) + '*(1-F' + IntToStr(R + 1) + ')';
        end;
      end;
      XLS.CmdFormat.Apply(C,R,C,R);
    end;
  end;

  for C := 1 to 5 do begin
    XLS.CmdFormat.Clear;
    XLS.CmdFormat.Alignment.Vertical := cvaCenter;
    XLS.CmdFormat.Border.Style := cbsMedium;
    XLS.CmdFormat.Border.Side[cbsTop] := True;
    XLS.CmdFormat.Border.Style := cbsThin;
    XLS.CmdFormat.Border.Color.ExcelSwatch(3,2);
    XLS.CmdFormat.Border.Side[cbsBottom] := True;
    XLS.CmdFormat.Font.Size := 9;
    XLS.CmdFormat.Font.Style := [xfsBold];
    case C of
      1,5: begin
        if C = 1 then
          XLS.CmdFormat.Border.Side[cbsLeft] := True
        else
          XLS.CmdFormat.Border.Side[cbsRight] := True;
      end;
      3: XLS[0].AsString[C,25] := 'Subtotal';
      4: begin
        XLS.CmdFormat.Alignment.Horizontal := chaLeft;
        XLS.CmdFormat.Number.Format := '_("$"* # ##0.00_);_("$"* (# ##0.00);_("$"* "-"??_);_(@_)';
        XLS[0].AsFormula[C,25] := 'SUM(E17:E25)';
      end;
    end;
    XLS.CmdFormat.Apply(C,25,C,25);
  end;

  XLS.CmdFormat.Clear;
  XLS.CmdFormat.Alignment.Horizontal := chaRight;
  XLS.CmdFormat.Alignment.Indent := 1;
  XLS.CmdFormat.Font.Size := 9;
  XLS.CmdFormat.Font.Style := [xfsBold];
  XLS.CmdFormat.Apply(3,27,3,27);
  XLS[0].AsString[3,27] := 'Credit';

  XLS.CmdFormat.Clear;
  XLS.CmdFormat.Fill.BackgroundColor.ExcelSwatch(3,1);
  XLS.CmdFormat.Number.Format := '_("$"* # ##0.00_);_("$"* (# ##0.00);_("$"* "-"??_);_(@_)';
  XLS.CmdFormat.Apply(4,27,4,27);

  XLS.CmdFormat.Clear;
  XLS.CmdFormat.Alignment.Horizontal := chaRight;
  XLS.CmdFormat.Alignment.Indent := 1;
  XLS.CmdFormat.Font.Size := 9;
  XLS.CmdFormat.Font.Style := [xfsBold];
  XLS.CmdFormat.Apply(3,28,3,28);
  XLS[0].AsString[3,28] := 'Additional discount';

  XLS.CmdFormat.Clear;
  XLS.CmdFormat.Fill.BackgroundColor.ExcelSwatch(7,2);
  XLS.CmdFormat.Number.Format := '0%';
  XLS.CmdFormat.Apply(4,28,4,28);

  XLS.CmdFormat.Clear;
  XLS.CmdFormat.Alignment.Horizontal := chaRight;
  XLS.CmdFormat.Alignment.Indent := 1;
  XLS.CmdFormat.Font.Size := 11;
  XLS.CmdFormat.Font.Style := [xfsBold];
  XLS.CmdFormat.Apply(3,29,3,29);
  XLS[0].AsString[3,29] := 'Balance due';

  XLS.CmdFormat.Clear;
  XLS.CmdFormat.Fill.BackgroundColor.ExcelSwatch(7,3);
  XLS.CmdFormat.Number.Format := '_("$"* # ##0.00_);_("$"* (# ##0.00);_("$"* "-"??_);_(@_)';
  XLS.CmdFormat.Font.Style := [xfsBold];
  XLS.CmdFormat.Apply(4,29,4,29);
  XLS[0].AsFormula[4,29] := 'E26-E28-IF(E29>0,E29*E26,0)';

  XLS.CmdFormat.Clear;
  XLS.CmdFormat.Border.Style := cbsThin;
  XLS.CmdFormat.Border.Color.ExcelSwatch(3,1);
  XLS.CmdFormat.Border.Preset(cbspOutline);
  XLS.CmdFormat.Apply(4,27,4,29);

  AddTextBox(1,27,3,30,0,8,
             'Make all checks payable to [CompanyName]. If you have any questions concerning this invoice, contact [Name] at (206) 555-1163, someone@example.com.' + #13 +
             'Thank you for your business!');


  XLS[0].Drawing.InsertImage('sample.jpeg',1,1,0,0,0.8);

  XLS.Calculate;
end;

procedure TForm1.btnCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TForm1.btnProtectClick(Sender: TObject);
begin
  // Protect the cells in order to prevent users from changing cell values.
  // Column A (Quantity) is editable. This is done by removing the default
  // Locked option.
  XLS.CmdFormat.BeginEdit(XLS[0]);
  XLS.CmdFormat.Protection.Locked := False;
  XLS.CmdFormat.Apply(1,16,1,24);

  // Assign a password. Without a password the cells won't de protected.
  XLS[0].Protection.Password := edPassword.Text;
end;

procedure TForm1.btnExportHtmlClick(Sender: TObject);
begin
  XLSHTML.Filename := edHTMLFilename.Text;
  XLSHTML.Write;
end;

procedure TForm1.btnSaveClick(Sender: TObject);
begin
  XLS.Filename := edFilename.Text;
  XLS.Write;
end;

procedure TForm1.btnOpenExcelClick(Sender: TObject);
begin
  // If you receive a compilation error here, change PAnsiChar to PWideChar.
  ShellExecute(Handle,'open', 'excel.exe',PAnsiChar(edFilename.Text), nil, SW_SHOWNORMAL);
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  dlgOpen.Filter := 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*';
  dlgOpen.FileName := edFilename.Text;
  if dlgOpen.Execute then
    edFilename.Text := dlgOpen.FileName;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  dlgOpen.Filter := 'HTML files (*.htm)|*.htm|All files (*.*)|*.*';
  dlgOpen.FileName := edHTMLFilename.Text;
  if dlgOpen.Execute then
    edHTMLFilename.Text := dlgOpen.FileName;
end;

procedure TForm1.btnFindReplaceClick(Sender: TObject);
var
  C,R: integer;
  S: string;
begin
  XLS[0].BeginIterate;

  while XLS[0].IterateNext do begin
    XLS[0].IteratePos(C,R);
    S := Uppercase(XLS[0].AsString[C,R]);

    if S = '[DATE]' then
      XLS[0].AsString[C,R] := DateToStr(Date)
    else if S = '[INVOICENO]' then
      XLS[0].AsFloat[C,R] := 723578934
    else if S = '[CONTACT]' then
      XLS[0].AsString[C,R] := 'John Smith'
    else if S = '[COMPANY]' then
      XLS[0].AsString[C,R] := 'Big Mountain Export and Import'
    else if S = '[STREET]' then
      XLS[0].AsString[C,R] := 'Main street'
    else if S = '[CITY]' then
      XLS[0].AsString[C,R] := 'Medium Size City'
    else if S = '[PHONE]' then
      XLS[0].AsString[C,R] := '524 772 902';
  end;
end;

end.
