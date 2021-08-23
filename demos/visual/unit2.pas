unit unit2;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, DB, Forms, Controls, Graphics, Dialogs, DBCtrls, StdCtrls,
  DBGrids, fpsDataset, xlsxOOXML;

type

  { TForm1 }

  TForm1 = class(TForm)
    CheckBox1: TCheckBox;
    DataSource1: TDataSource;
    DBCheckBox1: TDBCheckBox;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    DBEdit4: TDBEdit;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    Label1: TLabel;
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
end;

procedure TForm1.CheckBox1Change(Sender: TObject);
begin
  if Checkbox1.Checked then
    FDataset.OnFilterRecord := @FilterRecord
  else
    FDataset.OnFilterRecord := nil;
  FDataset.Filtered := Checkbox1.Checked;
end;

end.

