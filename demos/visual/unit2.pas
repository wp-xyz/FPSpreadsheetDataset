unit unit2;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, DB, Forms, Controls, Graphics, Dialogs, DBCtrls, StdCtrls,
  DBGrids, fpsDataset, xlsxOOXML;

type

  { TForm1 }

  TForm1 = class(TForm)
    btnFind: TButton;
    Button1: TButton;
    CheckBox1: TCheckBox;
    cmbFields: TComboBox;
    DataSource1: TDataSource;
    DBCheckBox1: TDBCheckBox;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    DBEdit4: TDBEdit;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    edKeyValue: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    procedure btnFindClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure CheckBox1Change(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    FDataset: TsWorksheetDataset;
    procedure AfterScrollHandler(Dataset: TDataset);
    procedure FilterRecord(Dataset: TDataset; var Accept: Boolean);

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

uses
  Variants;

const
  DATA_FILE = '../TestData.xlsx';

{ TForm1 }

procedure TForm1.AfterScrollHandler(Dataset: TDataset);
begin
  Label1.Caption := 'Record number: ' + IntToStr(Dataset.RecNo);
end;

procedure TForm1.FilterRecord(Dataset: TDataset; var Accept: Boolean);
begin
  Accept := Dataset.FieldByName('IntCol').AsInteger > 3;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  FDataset := TsWorksheetDataset.Create(self);
  FDataset.Filename := DATA_FILE;
  FDataset.SheetName := 'Sheet';
  FDataset.AfterScroll := @AfterScrollHandler;
  FDataset.Open;
  DataSource1.Dataset := FDataset;
  DBEdit1.Datafield := 'IntCol';
  DBEdit2.DataField := 'StringCol3';
  DBEdit3.Datafield := 'StringCol5';
  DBEdit4.DataField := 'DateCol';
  DBCheckbox1.DataField := 'BoolCol';
  (FDataset.FieldByName('FloatCol') as TFloatField).DisplayFormat := '0.000';

  FDataset.GetFieldNames(cmbFields.Items);
  cmbFields.ItemIndex := 0;
end;

procedure TForm1.CheckBox1Change(Sender: TObject);
begin
  if Checkbox1.Checked then
    FDataset.OnFilterRecord := @FilterRecord
  else
    FDataset.OnFilterRecord := nil;
  FDataset.Filtered := Checkbox1.Checked;
end;

procedure TForm1.btnFindClick(Sender: TObject);
begin
  if FDataset.Locate(cmbFields.Items[cmbFields.ItemIndex], edKeyValue.Text, []) then
    ShowMessage('Found')
  else
    ShowMessage('Not found');
end;

procedure TForm1.Button1Click(Sender: TObject);
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

end.

