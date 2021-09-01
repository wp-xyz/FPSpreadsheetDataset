unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, DBGrids, fpsDataset,
  DB;

type

  { TForm1 }

  TForm1 = class(TForm)
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    sWorksheetDataset1: TsWorksheetDataset;
    sWorksheetDataset1AutoIncCol: TLongintField;
    sWorksheetDataset1BoolCol: TBooleanField;
    sWorksheetDataset1calculated: TFloatField;
    sWorksheetDataset1CurrencyCol: TCurrencyField;
    sWorksheetDataset1DateCol: TDateTimeField;
    sWorksheetDataset1FloatCol: TFloatField;
    sWorksheetDataset1IntCol: TLongintField;
    sWorksheetDataset1MemoCol: TMemoField;
    sWorksheetDataset1SmallIntCol: TLongintField;
    sWorksheetDataset1StringCol3: TStringField;
    sWorksheetDataset1StringCol5: TStringField;
    sWorksheetDataset1WideStringCol: TStringField;
    sWorksheetDataset1Wordcol: TLongintField;
    procedure FormCreate(Sender: TObject);
    procedure sWorksheetDataset1CalcFields(DataSet: TDataSet);
  private

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.FormCreate(Sender: TObject);
begin
  sWorksheetDataset1.FileName := 'D:\Prog_Lazarus\wp-git\FPSpreadsheetDataset\demos\TestData.xlsx';
  sWorksheetDataset1.Open;
end;

procedure TForm1.sWorksheetDataset1CalcFields(DataSet: TDataSet);
begin
  sWorksheetDataset1calculated.AsFloat := sWorksheetDataset1FloatCol.AsFloat + 1.0;
end;

end.

