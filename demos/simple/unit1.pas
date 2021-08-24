unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, DB, Forms, Controls, Graphics, Dialogs, StdCtrls,
  fpsDataset, fpsAllFormats;

type

  { TForm1 }

  TForm1 = class(TForm)
    btnCloseOpen: TButton;
    btnListValues: TButton;
    btnFirst: TButton;
    btnPrior: TButton;
    btnNext: TButton;
    btnLast: TButton;
    btnFind: TButton;
    cmbFields: TComboBox;
    edKeyValue: TEdit;
    Label1: TLabel;
    Memo1: TMemo;
    procedure btnCloseOpenClick(Sender: TObject);
    procedure btnFindClick(Sender: TObject);
    procedure btnListValuesClick(Sender: TObject);
    procedure btnFirstClick(Sender: TObject);
    procedure btnPriorClick(Sender: TObject);
    procedure btnNextClick(Sender: TObject);
    procedure btnLastClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    WorksheetDataset: TsWorksheetDataset;
    procedure AfterOpen(Dataset: TDataset);
    procedure AfterClose(Dataset: TDataset);
    procedure AfterScroll(Dataset: TDataset);
    procedure DataInfo(Dataset: TDataset; ACaption: String);
    procedure ListBuffers;
    procedure ListFieldValues;
  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

uses
  TypInfo, Variants;

const
  DATA_FILE = '../TestData.xlsx';

{ TForm1 }

procedure TForm1.AfterOpen(Dataset: TDataset);
begin
  DataInfo(Dataset, 'After Open');
end;

procedure TForm1.AfterClose(Dataset:TDataset);
begin
  DataInfo(Dataset, 'After Close');
end;

procedure TForm1.AfterScroll(Dataset: TDataset);
begin
  ListFieldValues;
  ListBuffers;
end;

procedure TForm1.DataInfo(Dataset: TDataset; ACaption: String);
const
  CLOSE_OPEN: array [boolean] of String = ('CLOSED', 'OPEN');
var
  i: Integer;
  s: String;
begin
  Memo1.Lines.Add(ACaption + ':');
  Memo1.Lines.Add('  Dataset is ' + CLOSE_OPEN[Dataset.Active]);
  Memo1.Lines.Add('  FieldDefs.Count: ' + IntToStr(Dataset.FieldDefs.count));
  for i := 0 to Dataset.FieldDefs.Count-1 do
  begin
    Memo1.Lines.Add(Format('  FieldDefs[%d]', [i]));
    Memo1.Lines.Add(Format('    Name: %s', [Dataset.FieldDefs[i].Name]));
    Memo1.Lines.Add(Format('    DataType: %s', [GetEnumName(TypeInfo(TFieldType), integer(Dataset.FieldDefs[i].DataType))]));
    Memo1.Lines.Add(Format('    Size: %d', [Dataset.FieldDefs[i].Size]));
    if (Dataset is TsWorksheetDataset) then
      Memo1.Lines.Add(Format('    Column: %d', [TsFieldDef(Dataset.FieldDefs[i]).Column]));
  end;

  Memo1.Lines.Add('  Fields.Count: ' + IntToStr(Dataset.Fields.count));
  for i := 0 to Dataset.Fields.Count-1 do
  begin
    Memo1.Lines.Add(Format('  Fields[%d]', [i]));
    Memo1.Lines.Add(Format('    FieldName: %s', [Dataset.Fields[i].FieldName]));
    Memo1.Lines.Add(Format('    DataType: %s', [GetEnumName(TypeInfo(TFieldType), integer(Dataset.Fields[i].DataType))]));
    Memo1.Lines.Add(Format('    DataSize: %d', [Dataset.Fields[i].DataSize]));
  end;

  ListFieldValues;
end;

procedure TForm1.ListFieldValues;
var
  i: Integer;
begin
  Memo1.Lines.Add('');
  Memo1.Lines.Add('LIST OF FIELD VALUES:');
  Memo1.Lines.Add('  Current record number: ' + IntToStr(WorksheetDataset.RecNo) + ' of ' + IntToStr(WorksheetDataset.RecordCount));
  for i := 0 to WorksheetDataset.Fields.Count-1 do
    Memo1.Lines.Add('  Fields[' + IntToStr(i) + '].AsString: ' + WorksheetDataset.Fields[i].AsString);
end;

procedure TForm1.ListBuffers;
var
  i: Integer;
  buf: TRecordBuffer;
  bm: Integer;
  flag: TBookmarkFlag;
begin
  (*
  Memo1.Lines.Add('');
  Memo1.Lines.Add('LIST OF BUFFERS:');
  Memo1.Lines.Add('  BufferCount: ' + IntToStr(WorksheetDataset.BufferCount));
  for i := 0 to WorksheetDataset.BufferCount-1 do
  begin
    buf := WorksheetDataset.Buffers[i];
    bm := WorksheetDataset.GetRecordInfoptr(buf)^.Bookmark;
    flag := WorksheetDataset.GetRecordInfoPtr(buf)^.BookmarkFlag;
    Memo1.Lines.Add(Format('  Buffer #%d: Bookmark=%d Flag=%s', [
      i, bm, GetEnumName(TypeInfo(TBookmarkFlag), integer(flag))]));
  end;
  *)
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  DeleteFile('a');
  WorksheetDataset := TsWorksheetDataset.Create(self);
  WorksheetDataset.AfterOpen := @AfterOpen;
  WorksheetDataset.AfterClose := @AfterClose;
  WorksheetDataset.AfterScroll := @AfterScroll;
  WorksheetDataset.FileName := DATA_FILE;
  WorksheetDataset.SheetName := 'Sheet';
  WorksheetDataset.Open;

  WorksheetDataset.GetFieldNames(cmbFields.Items);
  cmbFields.ItemIndex := 0;
end;

procedure TForm1.btnCloseOpenClick(Sender: TObject);
begin
  if btnCloseOpen.Caption = 'Close' then
  begin
    Worksheetdataset.Close;
    btnCloseOpen.Caption := 'Open';
  end else
  if btnCloseOpen.Caption = 'Open' then
  begin
    WorksheetDataset.Open;
    btnCloseOpen.Caption := 'Close';
  end;
end;

procedure TForm1.btnFindClick(Sender: TObject);
begin
  if WorksheetDataset.Locate(cmbFields.Items[cmbFields.ItemIndex], edKeyValue.Text, []) then
    ShowMessage(edKeyValue.Text + ' found.')
  else
    ShowMessage('Not found.');
end;

procedure TForm1.btnListValuesClick(Sender: TObject);
begin
  ListFieldValues;
  ListBuffers;
end;

procedure TForm1.btnFirstClick(Sender: TObject);
begin
  WorksheetDataset.First;
end;

procedure TForm1.btnPriorClick(Sender: TObject);
begin
  WorksheetDataset.Prior;
end;

procedure TForm1.btnNextClick(Sender: TObject);
begin
  WorksheetDataset.Next;
end;

procedure TForm1.btnLastClick(Sender: TObject);
begin
  WorksheetDataset.Last;
end;

end.

