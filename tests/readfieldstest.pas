unit ReadFieldsTest;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpcunit, testutils, testregistry,
  DB,
  fpspreadsheet, fpstypes, fpsdataset;

type

  TReadFieldsTest= class(TTestCase)
  private
    FWorkbook: TsWorkbook;
    FWorksheet: TsWorksheet;
    function CreateAndOpenDataset: TsWorksheetDataset;
    procedure ReadFieldTest(Col: Integer; FieldName: String);
  protected
    procedure SetUp; override;
    procedure TearDown; override;
  published
    procedure ReadIntegerField;
    procedure ReadByteField;
    procedure ReadWordField;
    procedure ReadFloatField;
    procedure ReadStringField;
    procedure ReadMemoField;
    procedure ReadBoolField;
    procedure ReadDateField;
    procedure ReadTimeField;
    procedure ReadDateTimeField;
  end;

implementation

const
  FILE_NAME = 'testfile.xlsx';
  SHEET_NAME = 'Sheet';

  INT_COL = 0;
  BYTE_COL = 1;
  WORD_COL = 2;
  FLOAT_COL = 3;
  STRING_COL = 4;
  BOOL_COL = 5;
  DATE_COL = 6;
  TIME_COL = 7;
  DATETIME_COL = 8;
  MEMO_COL = 9;

  LoremIpsum: array[0..1] of string = (
    'Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua',
    'At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet'
  );


var
  DataFileName: String;

function TReadFieldsTest.CreateAndOpenDataset: TsWorksheetDataset;
begin
  Result := TsWorksheetDataset.Create(nil);
  Result.AutoFieldDefs:= true;
  Result.FileName := DataFileName;
  Result.SheetName := SHEET_NAME;
  Result.Open;
end;

procedure TReadFieldsTest.SetUp;
const
  NumRows = 10;
var
  r: Integer;
begin
  // Create test spreadsheet file
  FWorkbook := TsWorkbook.Create;
  // Create worksheet
  FWorksheet := FWorkbook.AddWorkSheet(SHEET_NAME);
  // Write headers (= field names)
  FWorksheet.WriteText(0, INT_COL, 'IntCol');
  FWorksheet.WriteText(0, BYTE_COL, 'ByteCol');
  FWorksheet.WriteText(0, WORD_COL, 'WordCol');
  FWorksheet.WriteText(0, FLOAT_COL, 'FloatCol');
  FWorksheet.WriteText(0, STRING_COL, 'StringCol');
  FWorksheet.WriteText(0, BOOL_COL, 'BoolCol');
  FWorksheet.WriteText(0, DATE_COL, 'DateCol');
  FWorksheet.WriteText(0, TIME_COL, 'TimeCol');
  FWorksheet.Writetext(0, DATETIME_COL, 'DateTimeCol');
  FWorksheet.Writetext(0, MEMO_COL, 'MemoCol');
  for r := 1 to NumRows do begin
    // Write values to IntCol
    FWorksheet.WriteNumber(r, INT_COL, r*120- 50, nfFixed, 0);
    // Write values to ByteCol
    FWorksheet.WriteNumber(r, BYTE_COL, r*2, nfFixed, 0);
    //Write values to WordCol
    FWorksheet.WriteNumber(r, WORD_COL, r*3, nfFixed, 0);
    // Write values to FloatCol
    FWorksheet.WriteNumber(r, FLOAT_COL, r*1.1-5.1, nfFixed, 2);
    // Write values to StringCol
    FWorksheet.WriteText(r, STRING_COL, char(ord('A') + r-1) + char(ord('b') + r-1) + char(ord('c') + r-1));
    // Write values to BoolCol
    FWorksheet.WriteBoolValue(r, BOOL_COL, odd(r));
    // Write values to DateCol
    FWorksheet.WriteDateTime(r, DATE_COL, EncodeDate(2021, 8, 1) + r-1, nfShortDate);
    // Write values to TimeCol
    FWorksheet.WriteDateTime(r, TIME_COL, EncodeTime(8, 0, 0, 0) + (r-1) / (24*60), nfShortTime);
    // Write value to DateTimeCol
    FWorksheet.WriteDateTime(r, DATETIME_COL, EncodeDate(2021, 8, 1) + EncodeTime(8, 0, 0, 0) + (r-1) + (r-1)/24, nfShortDateTime);
    // Write value to MemoCol
    FWorksheet.WriteText(r, MEMO_COL, LoremIpsum[r mod Length(LoremIpsum)]);
  end;
  // Save
  DataFileName := GetTempDir + FILE_NAME;
  FWorkbook.WriteToFile(DataFileName, true);
end;

procedure TReadFieldsTest.TearDown;
begin
  FreeAndNil(FWorkbook);
  if FileExists(DataFileName) then DeleteFile(DataFileName);
end;

procedure TReadFieldsTest.ReadFieldTest(Col: Integer; FieldName: String);
const
  FLOAT_EPS = 1E-9;
var
  dataset: TDataset;
  row: Integer;
  f: TField;
  dt: TDateTime;
begin
  dataset := CreateAndOpenDataset;
  try
    f := dataset.FieldByName(FieldName);

    CheckEquals(
      FWorksheet.ReadAsText(0, col),
      f.FieldName,
      'Column header / FieldName mismatch'
    );

    CheckEquals(
      col,
      f.FieldNo-1,
      'Field number mismatch'
    );

    CheckEquals(
      FWorksheet.GetLastRowIndex(true),
      dataset.RecordCount,
      'Row count / record count mismatch'
    );

    dataset.First;
    row := 1;
    while not dataset.EoF do
    begin
      if (f.DataType in [ftString, ftWideString, ftMemo]) then
        CheckEquals(
          FWorksheet.ReadAsText(row, col),
          f.AsString,
          'Text mismatch in row ' + IntToStr(row)
        )
      else if (f.DataType in [ftInteger, ftByte, ftWord, ftSmallInt, ftLargeInt]) then
        CheckEquals(
          round(FWorksheet.ReadAsNumber(row, col)),
          f.AsInteger,
          'Integer value mismatch in row ' + IntToStr(row)
        )
      else if (f.DataType in [ftFloat]) then
        CheckEquals(
          FWorksheet.ReadAsNumber(row, col),
          f.AsFloat,
          FLOAT_EPS,
          'Float value mismatch in row ' + IntToStr(row)
        )
      else if (f.DataType = ftDate) then
      begin
        CheckEquals(
          true,
          FWorksheet.ReadAsDateTime(row, col, dt),
          'Invalid date in row ' + IntToStr(row)
        );
        CheckEquals(
          dt,
          f.AsDateTime,
          FLOAT_EPS,
          'Date value mismatch in row ' + IntToStr(row)
        )
      end
      else if (f.DataType = ftTime) then
      begin
        CheckEquals(
          true,
          FWorksheet.ReadAsDateTime(row, col, dt),
          'Invalid time in row ' + IntToStr(row)
        );
        CheckEquals(
          dt,
          f.AsDateTime,
          FLOAT_EPS,
          'Time value mismatch in row ' + IntToStr(row)
        )
      end
      else if (f.DataType = ftDateTime) then
      begin
        CheckEquals(
          true,
          FWorksheet.ReadAsDateTime(row, col, dt),
          'Invalid date/time in row ' + IntToStr(row)
        );
        CheckEquals(
          dt,
          f.AsDateTime,
          FLOAT_EPS,
          'Date/time value mismatch in row ' + IntToStr(row)
        );
      end;
      inc(row);
      dataset.Next;
    end;

  finally
    dataset.Free;
  end;
end;


procedure TReadFieldsTest.ReadIntegerField;
begin
  ReadFieldTest(INT_COL, 'IntCol');
end;

procedure TReadFieldsTest.ReadByteField;
begin
  ReadFieldTest(BYTE_COL, 'ByteCol');
end;

procedure TReadFieldsTest.ReadWordField;
begin
  ReadFieldTest(WORD_COL, 'WordCol');
end;

procedure TReadFieldsTest.ReadFloatField;
begin
  ReadFieldTest(FLOAT_COL, 'FloatCol');
end;

procedure TReadFieldsTest.ReadStringField;
begin
  ReadFieldTest(STRING_COL, 'StringCol');
end;

procedure TReadFieldsTest.ReadMemoField;
begin
  ReadFieldTest(MEMO_COL, 'MemoCol');
end;

procedure TReadFieldsTest.ReadBoolField;
begin
  ReadFieldTest(BOOL_COL, 'BoolCol');
end;

procedure TReadFieldsTest.ReadDateField;
begin
  ReadFieldTest(DATE_COL, 'DateCol');
end;

procedure TReadFieldsTest.ReadTimeField;
begin
  ReadFieldTest(TIME_COL, 'TimeCol');
end;

procedure TReadFieldsTest.ReadDateTimeField;
begin
  ReadFieldTest(DATETIME_COL, 'DateTimeCol');
end;



initialization

  RegisterTest(TReadFieldsTest);
end.

