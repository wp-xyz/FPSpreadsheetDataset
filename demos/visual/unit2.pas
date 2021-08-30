unit unit2;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ExtCtrls,
  DB, DBCtrls, DBGrids,
  fpsDataset, xlsxOOXML;

type

  { TForm1 }

  TForm1 = class(TForm)
    Bevel1: TBevel;
    btnFind: TButton;
    btnLookup: TButton;
    btnSetBookmark: TButton;
    btnGoToBookmark: TButton;
    Button2: TButton;
    cbFilter: TCheckBox;
    CheckBox1: TCheckBox;
    cmbFields: TComboBox;
    cmbFilterFields: TComboBox;
    cmbFilterOp: TComboBox;
    DataSource1: TDataSource;
    DBCheckBox1: TDBCheckBox;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    DBEdit4: TDBEdit;
    DBGrid1: TDBGrid;
    DBMemo1: TDBMemo;
    DBNavigator1: TDBNavigator;
    edFilterText: TEdit;
    edKeyValue: TEdit;
    edFilterValue: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    procedure btnFindClick(Sender: TObject);
    procedure btnGoToBookmarkClick(Sender: TObject);
    procedure btnLookupClick(Sender: TObject);
    procedure btnSetBookmarkClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure cbFilterChange(Sender: TObject);
    procedure CheckBox1Change(Sender: TObject);
    procedure cmbFilterFieldsChange(Sender: TObject);
    procedure cmbFilterOpChange(Sender: TObject);
    procedure edFilterValueChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    FDataset: TsWorksheetDataset;
    FBookmark: TBookmark;
    procedure AfterScrollHandler(Dataset: TDataset);
    procedure ExecFilter;
    procedure FilterRecord_String(Dataset: TDataset; var Accept: Boolean);
    procedure FilterRecord_Float(Dataset: TDataset; var Accept: Boolean);
    procedure FilterRecord_Integer(Dataset: TDataset; var Accept: Boolean);

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

uses
  Variants, Math;

const
  DATA_FILE = '../TestData1.xlsx';

{ TForm1 }

procedure TForm1.FormCreate(Sender: TObject);
var
  i: Integer;
begin
  FDataset := TsWorksheetDataset.Create(self);
  FDataset.Filename := DATA_FILE;
  FDataset.SheetName := 'Sheet';
  FDataset.AfterScroll := @AfterScrollHandler;

  FDataset.AutoFieldDefs := false;
//  FDataset.FieldDefs.Add('MemoCol', ftMemo);
  FDataset.FieldDefs.Add('AutoIncCol', ftAutoInc);
  FDataset.FieldDefs.Add('IntCol', ftInteger);
  FDataset.FieldDefs.Add('SmallIntCol', ftSmallInt);
  FDataset.FieldDefs.Add('WordCol', ftWord);
  FDataset.FieldDefs.Add('StringCol3', ftString, 3);
  FDataset.FieldDefs.Add('StringCol5', ftString, 5);
  FDataset.FieldDefs.Add('WideStringCol', ftWideString, 12);
  FDataset.FieldDefs.Add('FloatCol', ftFloat);
  FDataset.FieldDefs.Add('DateCol', ftDate);
  FDataset.FieldDefs.Add('BoolCol', ftBoolean);
  FDataset.FieldDefs.Add('CurrencyCol', ftCurrency);

  for i := 0 to FDataset.FieldDefs.Count-1 do
    TsFieldDef(FDataset.FieldDefs[i]).Column := i;

  FDataset.Open;
  DataSource1.Dataset := FDataset;

  DBEdit1.Datafield := 'IntCol';
  DBEdit2.DataField := 'StringCol3';
  DBEdit3.Datafield := 'StringCol5';
  DBEdit4.DataField := 'DateCol';
  DBCheckbox1.DataField := 'BoolCol';
//  DBMemo1.Datafield := 'MemoCol';
  (FDataset.FieldByName('FloatCol') as TFloatField).DisplayFormat := '0.000';

  FDataset.GetFieldNames(cmbFields.Items);
  FDataset.GetFieldNames(cmbFilterFields.Items);
  cmbFields.ItemIndex := 0;
  cmbFilterFields.ItemIndex := 0;

end;

procedure TForm1.AfterScrollHandler(Dataset: TDataset);
begin
  Label1.Caption := 'Record number: ' + IntToStr(Dataset.RecNo);
end;

procedure TForm1.FilterRecord_String(Dataset: TDataset; var Accept: Boolean);
var
  field: TField;
  fieldname: string;
  op: String;
  value: String;
begin
  fieldname := cmbFilterFields.Items[cmbFilterFields.ItemIndex];
  op := cmbFilterOp.Items[cmbFilterOp.ItemIndex];
  value := edFilterValue.Text;

  field := Dataset.FieldByName(fieldname);
  case op of
    '=': Accept := field.AsString = value;
    '<': Accept := field.AsString < value;
    '>': Accept := field.AsString > value;
  end;
end;

procedure TForm1.FilterRecord_Integer(Dataset: TDataset; var Accept: Boolean);
var
  field: TField;
  fieldname: string;
  op: String;
  value: Integer;
begin
  fieldname := cmbFilterFields.Items[cmbFilterFields.ItemIndex];
  op := cmbFilterOp.Items[cmbFilterOp.ItemIndex];
  value := StrToInt(edFilterValue.Text);

  field := Dataset.FieldByName(fieldname);
  case op of
    '=': Accept := field.AsInteger = value;
    '<': Accept := field.AsInteger < value;
    '>': Accept := field.AsInteger > value;
  end;
end;

procedure TForm1.FilterRecord_Float(Dataset: TDataset; var Accept: Boolean);
var
  field: TField;
  fieldname: string;
  op: String;
  value: Double;
begin
  fieldname := cmbFilterFields.Items[cmbFilterFields.ItemIndex];
  op := cmbFilterOp.Items[cmbFilterOp.ItemIndex];
  value := StrToFloat(edFilterValue.Text);

  field := Dataset.FieldByName(fieldname);
  case op of
    '=': Accept := SameValue(field.AsFloat, value);
    '<': Accept := (field.AsFloat < value) and not SameValue(field.AsFloat, value);
    '>': Accept := (field.AsFloat > value) and not SameValue(field.AsFloat, value);
  end;
end;

procedure TForm1.cbFilterChange(Sender: TObject);
begin
  ExecFilter;
end;

procedure TForm1.CheckBox1Change(Sender: TObject);
begin
  FDataset.Filtered := false;
  if Checkbox1.Checked then
  begin
    FDataset.Filter := edFilterText.Text;
    FDataset.Filtered := true;
  end;
end;

procedure TForm1.ExecFilter;
var
  field: TField;
  fieldName: String;
begin
  if cbFilter.Checked then
  begin
    fieldName := cmbFilterFields.Items[cmbFilterFields.ItemIndex];
    field := FDataset.FieldByName(fieldName);
    if field is TStringField then
      FDataset.OnFilterRecord := @FilterRecord_String
    else
    if field.DataType = ftInteger then
      FDataset.OnFilterRecord := @FilterRecord_Integer
    else
    if field is TFloatField then
      FDataset.OnFilterRecord := @FilterRecord_Float;
  end
  else
    FDataset.OnFilterRecord := nil;
  FDataset.Filtered := cbFilter.Checked;
end;

procedure TForm1.cmbFilterFieldsChange(Sender: TObject);
begin
  ExecFilter;
end;

procedure TForm1.cmbFilterOpChange(Sender: TObject);
begin
  ExecFilter;
end;

procedure TForm1.edFilterValueChange(Sender: TObject);
begin
  ExecFilter;
end;

procedure TForm1.btnFindClick(Sender: TObject);
begin
  if FDataset.Locate(cmbFields.Items[cmbFields.ItemIndex], edKeyValue.Text, []) then
    ShowMessage('Found')
  else
    ShowMessage('Not found');
end;

procedure TForm1.btnGoToBookmarkClick(Sender: TObject);
begin
  if FDataset.BookmarkValid(FBookmark) then
    FDataset.GotoBookmark(FBookmark);
end;

procedure TForm1.btnLookupClick(Sender: TObject);
var
  v: Variant;
  s: String;
  d: Double;
begin
  v := FDataset.Lookup(cmbFields.Items[cmbFields.ItemIndex], edKeyValue.Text, 'DateCol;FloatCol');
  if VarIsNull(v) then
    ShowMessage('Not found')
  else
  begin
    s := VarToStr(v[0]);
    d := v[1];
    ShowMessage('DateCol = ' + VarToStr(v[0]) + LineEnding + 'FloatCol = ' + FormatFloat('0.00', d));
  end;
end;

procedure TForm1.btnSetBookmarkClick(Sender: TObject);
begin
  FBookmark := FDataset.GetBookmark;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  ShowMessage(FDataset.Fields[0].AsString);
end;

end.

