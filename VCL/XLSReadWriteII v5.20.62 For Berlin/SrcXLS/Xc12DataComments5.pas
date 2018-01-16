unit Xc12DataComments5;

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}

interface

uses Classes, SysUtils, Contnrs,
     xpgPSimpleDOM,
     xc12Utils5, Xc12Common5, Xc12DataStyleSheet5,
     XLSUtils5;

type TXc12Author = class(TXLSStyleObject)
protected
     FName: AxUCString;

     function  GetName: AxUCString;
     procedure SetName(const Value: AxUCString);
     function  GetIndex: integer;
     procedure CalcHash; override;
public
     function  Equal(AItem: TXLSStyleObject): boolean; override;
     procedure Assign(AItem: TXc12Author);
     function  Copy: TXc12Author;

     property Name: AxUCString read GetName write SetName;
     end;

type TXc12Authors = class(TXLSStyleObjectList)
private
     function GetItems(Index: integer): TXc12Author;
protected
public
     constructor Create;

     procedure SetIsDefault; override;

     function  Find(const AName: AxUCString): TXc12Author;
     function  Add: TXc12Author; overload;
     function  Add(AName: AxUCString): TXc12Author; overload;

     property Items[Index: integer]: TXc12Author read GetItems; default;
     end;

type TXc12Comments = class;

     TXc12Comment = class(TXc12Data)
private
     procedure SetAuthor(const Value: TXc12Author);
     function  GetTempAuthorId: integer;
     procedure SetTempAuthorId(const Value: integer);
     function  GetItems(Index: integer): TXc12FontRun;
     function  GetRef: TXLSCellArea;
     procedure SetCol(const Value: integer);
     procedure SetRow(const Value: integer);
protected
     FOwner       : TXc12Comments;

     FCol         : integer;
     FRow         : integer;
     FOrigCol     : integer;
     FOrigRow     : integer;
     FText        : AxUCString;
     FFontRuns    : TXc12DynFontRunArray;
     FPhoneticRuns: TXc12DynPhoneticRunArray;
     FAuthor      : TXc12Author;
     FGUID        : AxUCString;
     FVML         : TXpgDOMNode;

     FAuthorId    : integer;

     // Helper vars for XSS.
     FCol1        : integer;
     FCol1Offs    : integer;
     FRow1        : integer;
     FRow1Offs    : integer;
     FCol2        : integer;
     FCol2Offs    : integer;
     FRow2        : integer;
     FRow2Offs    : integer;
     FColor       : longword;
public
     constructor Create(AOwner: TXc12Comments);
     destructor Destroy; override;

     function  Count: integer;
     procedure ClearFontRuns;
     function  AddFontRun: PXc12FontRun;
     procedure AssignFontRuns(const ARuns: TXc12DynFontRunArray);
     procedure SetDefault;
     procedure UpdateVML;

     property Col: integer read FCol write SetCol;
     property Row: integer read FRow write SetRow;
     property Ref: TXLSCellArea read GetRef;
     property Text: AxUCString read FText write FText;
     property FontRuns: TXc12DynFontRunArray read FFontRuns write FFontRuns;
     property PhoneticRuns: TXc12DynPhoneticRunArray read FPhoneticRuns write FPhoneticRuns;
     property Author: TXc12Author read FAuthor write SetAuthor;
     property AuthorId: integer read FAuthorId write FAuthorId;
     property TempAuthorId: integer read GetTempAuthorId write SetTempAuthorId;
     property GUID: AxUCString read FGUID write FGUID;
     property VML: TXpgDOMNode read FVML write FVML;

     property Col1        : integer read FCol1 write FCol1;
     property Col1Offs    : integer read FCol1Offs write FCol1Offs;
     property Row1        : integer read FRow1 write FRow1;
     property Row1Offs    : integer read FRow1Offs write FRow1Offs;
     property Col2        : integer read FCol2 write FCol2;
     property Col2Offs    : integer read FCol2Offs write FCol2Offs;
     property Row2        : integer read FRow2 write FRow2;
     property Row2Offs    : integer read FRow2Offs write FRow2Offs;
     property Color       : longword read FColor write FColor;

     property Items[Index: integer]: TXc12FontRun read GetItems; default;
     end;

     TXc12Comments = class(TObjectList)
private
     function GetItems(Index: integer): TXc12Comment;
protected
     FStyles: TXc12DataStyleSheet;

     FAuthors: TXc12Authors;
public
     constructor Create(AStyleSheet: TXc12DataStyleSheet);
     destructor Destroy; override;

     procedure Clear; override;
     function  Find(const ACol,ARow: integer): TXc12Comment;

     function Add: TXc12Comment;

     property Authors: TXc12Authors read FAuthors;
     property Items[Index: integer]: TXc12Comment read GetItems; default;
     end;

implementation

var
  L_CommentIdCounter: integer;

{ TXc12DataComment }

function TXc12Comment.AddFontRun: PXc12FontRun;
begin
  SetLength(FFontRuns,Length(FFontRuns) + 1);
  Result := @FFontRuns[High(FFontRuns)];
end;

procedure TXc12Comment.AssignFontRuns(const ARuns: TXc12DynFontRunArray);
var
  i: integer;
begin
  SetLength(FFontRuns,Length(ARuns));

  for i := 0 to High(ARuns) do begin
    FFontRuns[i] := ARuns[i];
  end;
end;

procedure TXc12Comment.ClearFontRuns;
begin
  SetLength(FFontRuns,0);
  SetLength(FPhoneticRuns,0);
end;

function TXc12Comment.Count: integer;
begin
  Result := Length(FFontRuns);
end;

constructor TXc12Comment.Create(AOwner: TXc12Comments);
begin
  FOwner := AOwner;

  FOrigCol := -1;
  FOrigRow := -1;

  FColor := XLS_DEFAULT_COLOR_COMMENT;
end;

destructor TXc12Comment.Destroy;
begin
  FOwner.FStyles.XFEditor.FreeFonts(FFontRuns);

  FAuthor := Nil;

  if FVML <> Nil then
    FVML.Free;
  inherited;
end;

function TXc12Comment.GetItems(Index: integer): TXc12FontRun;
begin
  Result := FFontRuns[Index];
end;

function TXc12Comment.GetRef: TXLSCellArea;
begin
  Result.Col1 := FCol;
  Result.Row1 := FRow;
  Result.Col2 := FCol;
  Result.Row2 := FRow;
end;

function TXc12Comment.GetTempAuthorId: integer;
begin
  Result := Integer(FAuthor);
end;

procedure TXc12Comment.SetAuthor(const Value: TXc12Author);
begin
  FAuthor := Value.Copy;
end;

procedure TXc12Comment.SetCol(const Value: integer);
begin
  FCol := Value;
  if FOrigCol < 0 then
    FOrigCol := Value;
end;

procedure TXc12Comment.SetDefault;
const DefaultVML =
'<v:shape id="vodka_%d" type="#_x0000_t202" style="position:absolute;width:120pt;height:59.25pt;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">' +
'  <v:fill color2="#ffffe1"/>' +
'  <v:shadow on="t" color="black" obscured="t"/>' +
'  <v:path o:connecttype="none"/>' +
'  <v:textbox style="mso-direction-alt:auto">' +
'    <div style="text-align:left"/>' +
'  </v:textbox>' +
'  <x:ClientData ObjectType="Note">' +
'    <x:MoveWithCells/>' +
'    <x:SizeWithCells/>' +
'    <x:AutoFill>False</x:AutoFill>' +
'    <x:Row>1</x:Row>' +
'    <x:Column>1</x:Column>' +
'  </x:ClientData>' +
'</v:shape>';
var
  S: AxUCString;
  DOM: TXpgSimpleDOM;
begin
  if FVML <> Nil then
    FVML.Free;

  Inc(L_CommentIdCounter);

  DOM := TXpgSimpleDOM.Create;
  try
    S := Format(DefaultVML,[L_CommentIdCounter]);
    DOM.LoadFromString(S);
    VML := DOM.Root.Detach(0);
  finally
    DOM.Free;
  end;
end;

procedure TXc12Comment.SetRow(const Value: integer);
begin
  FRow := Value;
  if FOrigRow < 0 then
    FOrigRow := Value;
end;

procedure TXc12Comment.SetTempAuthorId(const Value: integer);
begin
  FAuthor := TXc12Author(Value);
end;

procedure TXc12Comment.UpdateVML;
var
  Node,N: TXpgDOMNode;
  Lines: TStringList;
  DCol,DRow: integer;
begin
  if FVML <> Nil then begin
    Node := FVML.Find('x:ClientData',False);
    if Node <> Nil then begin
      N := Node.Find('x:Column',False);
      if N <> Nil then
        N.Content := IntToStr(FCol);
      N := Node.Find('x:Row',False);
      if N <> Nil then
        N.Content := IntToStr(FRow);

      if FOrigCol >= 0 then
        DCol := FCol - FOrigCol
      else
        DCol := 0;
      if FOrigRow >= 0 then
        DRow := FRow - FOrigRow
      else
        DRow := 0;

      if (DCol <> 0) or (DRow <> 0) then begin
        N := Node.Find('x:Anchor',False);
        if N <> Nil then begin
          Lines := TStringList.Create;
          try
{$ifdef DELPHI_5}
            DelimitedToStrings(N.Content,Lines,',');
{$else}
            Lines.Delimiter := ',';
            Lines.DelimitedText := N.Content;
{$endif}
            if Lines.Count = 8 then begin
              Lines[0] := IntToStr(StrToIntDef(Lines[0],0) + DCol);
              Lines[2] := IntToStr(StrToIntDef(Lines[2],0) + DRow);
              Lines[4] := IntToStr(StrToIntDef(Lines[4],0) + DCol);
              Lines[6] := IntToStr(StrToIntDef(Lines[6],0) + DRow);
{$ifdef DELPHI_5}
              N.Content := StringsToDelimited(Lines,',');
{$else}
              N.Content := Lines.DelimitedText;
{$endif}
            end;
          finally
            Lines.Free;
          end;
        end;
      end;
    end;
  end;
end;

{ TXc12DataComments }

function TXc12Comments.Add: TXc12Comment;
begin
  Result := TXc12Comment.Create(Self);
  inherited Add(Result);
end;

procedure TXc12Comments.Clear;
begin
  inherited Clear;
  FAuthors.Clear;
end;

constructor TXc12Comments.Create(AStyleSheet: TXc12DataStyleSheet);
begin
  inherited Create;
  FStyles := AStyleSheet;

  FAuthors := TXc12Authors.Create;
end;

destructor TXc12Comments.Destroy;
begin
  inherited;

  FAuthors.Free;
end;

function TXc12Comments.Find(const ACol, ARow: integer): TXc12Comment;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if (Items[i].Col = ACol) and (Items[i].Row = ARow) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXc12Comments.GetItems(Index: integer): TXc12Comment;
begin
  Result := TXc12Comment(inherited Items[Index]);
end;

{ TXc12Authors }

function TXc12Authors.Add: TXc12Author;
begin
  Result := TXc12Author.Create(Self);
  inherited Add(Result);
end;

function TXc12Authors.Add(AName: AxUCString): TXc12Author;
begin
  Result := Add;
  Result.Name := AName;
end;

constructor TXc12Authors.Create;
begin
  inherited Create;
end;

function TXc12Authors.Find(const AName: AxUCString): TXc12Author;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if SameText(AName,Items[i].Name) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXc12Authors.GetItems(Index: integer): TXc12Author;
begin
  Result := TXc12Author(inherited Items[Index]);
end;

procedure TXc12Authors.SetIsDefault;
begin
  inherited;

end;

{ TXc12Author }

procedure TXc12Author.Assign(AItem: TXc12Author);
begin
  FName := AItem.Name;
end;

procedure TXc12Author.CalcHash;
begin
  FHash := 0;
end;

function TXc12Author.Copy: TXc12Author;
begin
  Result := Self;
end;

function TXc12Author.Equal(AItem: TXLSStyleObject): boolean;
begin
  Result := SameText(FName,TXc12Author(AItem).Name);
end;

function TXc12Author.GetIndex: integer;
begin
  Result := FIndex;
end;

function TXc12Author.GetName: AxUCString;
begin
  Result := FName;
end;

procedure TXc12Author.SetName(const Value: AxUCString);
begin
  FName := Value;
end;

initialization
  L_CommentIdCounter := 0;

end.
