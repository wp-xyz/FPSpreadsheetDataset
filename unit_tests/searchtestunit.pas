unit SearchTestUnit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpcunit, testutils, testregistry,
  DB,
  fpspreadsheet, fpsTypes, fpsDataset;

type

  TSearchTest = class(TTestCase)
  private
    function CreateAndOpenDataset: TsWorksheetDataset;
    procedure LocateTest(SearchInField: String; SearchValue: Variant;
      ExpectedRecNo: Integer; Options: TLocateOptions = []);

  protected
    procedure SetUp; override;
    procedure TearDown; override;

  published
    procedure LocateTest_Int_Found;
    procedure LocateTest_Int_NotFound;
    procedure LocateTest_String_Found;
    procedure LocateTest_String_Found_CaseInsensitive;
    procedure LocateTest_String_NotFound;
    procedure LocateTest_NonASCIIString_Found;
    procedure LocateTest_NonASCIIString_Found_CaseInsensitive;
    procedure LocateTest_NonASCIIString_NotFound;
    procedure LocateTest_WideString_Found;
    procedure LocateTest_WideString_Found_CaseInsensitive;
    procedure LocateTest_WideString_NotFound;
    procedure LocateTest_NonASCIIWideString_Found;
    procedure LocateTest_NonASCIIWideString_Found_CaseInsensitive;
    procedure LocateTest_NonASCIIWideString_NotFound;
  end;

implementation

uses
  LazUTF8;

const
  FILE_NAME = 'testfile.xlsx';
  SHEET_NAME = 'Sheet';
  INT_COL = 0;
  STRING_COL = 1;
  WIDESTRING_COL = 2;
  INT_FIELD = 'IntCol';
  STRING_FIELD = 'StringCol';
  WIDESTRING_FIELD = 'WideStringCol';

var
  DataFileName: String;

const
  NUM_ROWS = 5;
var
  INT_VALUES: array[1..NUM_ROWS] of Integer = (
    12, 20, -10, 83, 3
  );
  STRING_VALUES: array[1..NUM_ROWS] of String = (
    'abc', 'a', 'Hallo', 'ijk', 'äöü'
  );
  WIDESTRING_VALUES: array[1..NUM_ROWS] of String = (  // Strings are converted to wide at runtime
    'ABC', 'A', 'Test', 'Äöü', 'xyz'
  );

function TSearchTest.CreateAndOpenDataset: TsWorksheetDataset;
begin
  Result := TsWorksheetDataset.Create(nil);
  Result.FileName := DataFileName;
  Result.SheetName := SHEET_NAME;
  Result.AutoFieldDefs := false;
  Result.AddFieldDef(INT_FIELD, ftInteger);
  Result.AddFieldDef(STRING_FIELD, ftString, 20);
  Result.AddFieldDef(WIDESTRING_FIELD, ftWideString, 20);
  Result.Open;
end;

procedure TSearchTest.LocateTest(SearchInField: String; SearchValue: Variant;
  ExpectedRecNo: Integer; Options: TLocateOptions = []);
var
  dataset: TsWorksheetDataset;
  actualRecNo: Integer;
  found: Boolean;
  f: TField;
begin
  dataset := CreateAndOpenDataset;
  try
    found := dataset.Locate(SearchInField, SearchValue, options);

    if ExpectedRecNo = -1 then
      CheckEquals(
        false,
        found,
        'Record found unexpectedly.'
      )
    else
      CheckEquals(
        true,
        found,
        'Existing record not found.'
      );

    if found then
    begin
      actualRecNo := dataset.RecNo;
      CheckEquals(
        ExpectedRecNo,
        actualRecNo,
        'Mismatch of found RecNo.'
      );

      for f in dataset.Fields do
        case f.FieldName of
          INT_FIELD:
            CheckEquals(
              INT_VALUES[actualRecNo],
              f.AsInteger,
              'Value mismatch in integer field'
            );
          STRING_FIELD:
            CheckEquals(
              STRING_VALUES[actualRecNo],
              f.AsString,
              'Value mismatch in string field'
            );
          WIDESTRING_FIELD:
            CheckEquals(
              UTF8ToUTF16(WIDESTRING_VALUES[actualRecNo]),
              f.AsWideString,
              'Value mismatch in widestring field'
            );
        end;
    end;
  finally
    dataset.Free;
  end;
end;

procedure TSearchTest.LocateTest_Int_Found;
begin
  LocateTest(INT_FIELD, -10, 3);
end;

procedure TSearchTest.LocateTest_Int_NotFound;
begin
  LocateTest(INT_FIELD, 1000, -1);
end;

procedure TSearchTest.LocateTest_String_Found;
begin
  LocateTest(STRING_FIELD, 'a', 2);
end;

procedure TSearchTest.LocateTest_String_Found_CaseInsensitive;
begin
  LocateTest(STRING_FIELD, 'ABC', 1, [loCaseInsensitive]);
end;

procedure TSearchTest.LocateTest_String_NotFound;
begin
  LocateTest(STRING_FIELD, 'ttt', -1);
end;

procedure TSearchTest.LocateTest_NonASCIIString_Found;
begin
  LocateTest(STRING_FIELD, 'äöü', 5);
end;

procedure TSearchTest.LocateTest_NonASCIIString_Found_CaseInsensitive;
begin
  LocateTest(STRING_FIELD, 'ÄöÜ', 5, [loCaseInsensitive]);
end;

procedure TSearchTest.LocateTest_NonASCIIString_NotFound;
begin
  LocateTest(STRING_FIELD, 'ä', -1);
end;

procedure TSearchTest.LocateTest_WideString_Found;
begin
  LocateTest(WIDESTRING_FIELD, WideString('ABC'), 1);
end;

procedure TSearchTest.LocateTest_WideString_Found_CaseInsensitive;
begin
  LocateTest(WIDESTRING_FIELD, WideString('Abc'), 1, [loCaseInsensitive]);
end;

procedure TSearchTest.LocateTest_WideString_NotFound;
begin
  LocateTest(WIDESTRING_FIELD, WideString('abc'), -1);
end;

procedure TSearchTest.LocateTest_NonASCIIWideString_Found;
var
  ws: WideString;
begin
  ws := UTF8ToUTF16('Äöü');
  LocateTest(WIDESTRING_FIELD, ws, 4);
end;

procedure TSearchTest.LocateTest_NonASCIIWideString_Found_CaseInsensitive;
var
  ws: Widestring;
begin
  ws := UTF8ToUTF16('Äöü');
  LocateTest(WIDESTRING_FIELD, ws, 4, [loCaseInsensitive]);
end;

procedure TSearchTest.LocateTest_NonASCIIWideString_NotFound;
var
  ws: WideString;
begin
  ws := UTF8ToUTF16('Würde');
  LocateTest(WIDESTRING_FIELD, ws, -1);
end;

procedure TSearchTest.SetUp;
var
  r: Integer;
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
begin
  // Create test spreadsheet file
  workbook := TsWorkbook.Create;
  try
    // Create worksheet
    worksheet := workbook.AddWorkSheet(SHEET_NAME);

    // Write headers (= field names)
    worksheet.WriteText(0, INT_COL, INT_FIELD);
    worksheet.WriteText(0, STRING_COL, STRING_FIELD);
    worksheet.WriteText(0, WIDESTRING_COL, WIDESTRING_FIELD);

    // Write values
    for r := 1 to NUM_ROWS do
    begin
      worksheet.WriteNumber(r, INT_COL, INT_VALUES[r], nfFixed, 0);
      worksheet.WriteText(r, STRING_COL, STRING_VALUES[r]);
      worksheet.WriteText(r, WIDESTRING_COL, WIDESTRING_VALUES[r]);
    end;

    // Save
    DataFileName := GetTempDir + FILE_NAME;
    workbook.WriteToFile(DataFileName, true);
  finally
    workbook.Free;
  end;
end;

procedure TSearchTest.TearDown;
begin
  if FileExists(DataFileName) then DeleteFile(DataFileName);
end;

initialization
  RegisterTest(TSearchTest);

end.

