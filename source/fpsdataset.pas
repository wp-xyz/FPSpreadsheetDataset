{@@ ----------------------------------------------------------------------------
  Unit **fpsDataset** implements a TDataset based on spreadsheet data.
  This way spreadsheets can be accessed in a database-like manner.
  Of course, it is required that all cells in a column have the same type.

  AUTHORS: Werner Pamler

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
           distribution, for details about the license.

  References:
  * https://www.delphipower.xyz/guide_8/building_custom_datasets.html
  * http://etutorials.org/Programming/mastering+delphi+7/Part+III+Delphi+Database-Oriented+Architectures/Chapter+17+Writing+Database+Components/Building+Custom+Datasets/
  * http://216.19.73.24/articles/customds.asp
  * https://delphi.cjcsoft.net/viewthread.php?tid=44220

  Much of the code is adapted from TMemDataset.

  Current status (Sept 01, 2021):

  Working
  * Field defs: determined automatically from file
  * Field defs defined by user: working (requires AutoFieldDefs = false)
  * Fields: working
  * Field types: ftFloat, ftInteger, ftAutoInc, ftByte, ftSmallInt, ftWord, ftLargeInt,
    ftCurrency, ftDateTime, ftDate, ftTime, ftString, ftFixedChar, ftBoolean,
    ftWideString, ftFixedWideString, ftMemo
  * Locate: working
  * Lookup: working
  * Edit, Delete, Insert, Append, Post, Cancel: working
  * NULL fields: working
  * GetBookmark, GotoBookmark: working
  * Filtering by OnFilter event and by Filter property: working.
  * Persistent and calculated fields working

  Planned but not yet working
  ' Field defs: Required, Unique etc possibly not supported ATM - to be tested
  * Indexes: not implemented
  * Sorting: not implemented
  * Auto-Format detection of string fields: use ftWideString rather than ftString
    (see issues below).

  Issues
  * TStringField and TMemoField by default store strings using code page CP_ACP.
    Because FPSpreadsheet works this way these fields should be created with
    CP_UTF8. However, such fields are created at max width because at worst a
    UTF8 code-point need 4 bytes. This means that the max text width of a UTF8
    cannot be controlled any more: when a field def is setup with Size=5 then
    the user can enter 5*4=20 ASCII characters!

    This does not happen for auto-detected text cells because they are created
    as TWideStringField.

  * Manually deleting a fielddef removes it from the object tree, but not from
    the lfm file.

  * Opening the dataset crashes when there is as empty cell at the end of
    an autoinc column.
-------------------------------------------------------------------------------}

unit fpsDataset;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, DB, BufDataset_Parser,
  fpSpreadsheet, fpsTypes, fpsUtils, fpsAllFormats;

type
  TRowIndex = Int64;
  TColIndex = Int64;
  PPCell = ^PCell;

  { TsFieldDef }
  TsFieldDef = class(TFieldDef)
  private
    FColIndex: TColIndex;
  public
    constructor Create(ACollection: TCollection); override;
    constructor Create(AOwner: TFieldDefs; const AName: string;
      ADataType: TFieldType; ASize: Integer; ARequired: Boolean; AFieldNo: Longint;
      AColIndex: TColIndex; ACodePage: TSystemCodePage = CP_ACP); overload;
    procedure Assign(ASource: TPersistent); override;
  published
    property ColIndex: TColIndex read FColIndex write FColIndex default -1;
  end;

  { TsFieldDefs }
  TsFieldDefs = class(TFieldDefs)
  protected
    class function FieldDefClass : TFieldDefClass; override;
  end;

  { TsRecordInfo }
  TsRecordInfo = record
    Bookmark: PCell;  // Pointer to a cell in the bookmarked row.
    BookmarkFlag: TBookmarkFlag;
  end;
  PsRecordInfo = ^TsRecordInfo;

  { TsWorksheetDataset }
  TsWorksheetDataset = class(TDataset)
  private
    FFileName: TFileName;
    FSheetName: String;
    FWorkbook: TsWorkbook;
    FWorksheet: TsWorksheet;
    FRecNo: Integer;                 // Current record number
    FFirstRow: TRowIndex;            // WorksheetIndex of the first record
    FLastRow: TRowIndex;             // Worksheet index of the last record
    FRecordCount: Integer;           // Number of records between first and last data rows
    FRecordBufferSize: Integer;      // Size of the record buffer
    FTotalFieldSize: Integer;        // Total size of the field data
    FFieldOffsets: array of Integer; // Offset to field start in buffer
    FModified: Boolean;              // Flag to show that workbook needs saving
    FFilterBuffer: TRecordBuffer;
    FTableCreated: boolean;
    FAutoFieldDefs: Boolean;
    FIsOpen: boolean;
    FParser: TBufDatasetParser;
    FAutoIncValue: Integer;
    FAutoIncField: TAutoIncField;
  private
    function FixFieldName(const AText: String): String;
    function GetActiveBuffer(out Buffer: TRecordBuffer): Boolean;
    function GetBookmarkCellFromRecNo(ARecNo: Integer): PCell;
    function GetCurrentRowIndex: TRowIndex;
    function GetFirstDataRowIndex: TRowIndex;
    function GetLastDataRowIndex: TRowIndex;
    function GetNullMaskPtr(Buffer: TRecordBuffer): Pointer;
    function GetNullMaskSize: Integer;
    function GetRecordInfoPtr(Buffer: TRecordBuffer): PsRecordInfo;
    function GetRowIndexFromRecNo(ARecNo: Integer): TRowIndex;
    procedure SetCurrentRow(ARow: TRowIndex);
  protected
    // methods inherited from TDataset
    function AllocRecordBuffer: TRecordBuffer; override;
    procedure ClearCalcFields(Buffer: TRecordBuffer); override;
    procedure DoBeforeOpen; override;
    class function FieldDefsClass : TFieldDefsClass; override;
    procedure FreeRecordBuffer(var Buffer: TRecordBuffer); override;
    procedure GetBookmarkData(Buffer: TRecordBuffer; Data: Pointer); override;
    function GetBookmarkFlag(Buffer: TRecordBuffer): TBookmarkFlag; override;
    function GetRecNo: LongInt; override;
    function GetRecord(Buffer: TRecordBuffer; GetMode: TGetMode;
      DoCheck: Boolean): TGetResult; override;
    function GetRecordCount: LongInt; override;
    function GetRecordSize: Word; override;
    procedure InternalAddRecord(Buffer: Pointer; DoAppend: Boolean); override;
    procedure InternalClose; override;
    procedure InternalDelete; override;
    procedure InternalFirst; override;
    procedure InternalGotoBookmark(ABookmark: Pointer); override;
    procedure InternalInitFieldDefs; override;
    procedure InternalInitRecord(Buffer: TRecordBuffer); override;
    procedure InternalLast; override;
    procedure InternalOpen; override;
    procedure InternalPost; override;
    procedure InternalSetToRecord(Buffer: TRecordBuffer); override;
    function IsCursorOpen: Boolean; override;
    procedure SetBookmarkData(Buffer: TRecordBuffer; Data: Pointer); override;
    procedure SetBookmarkFlag(Buffer: TRecordBuffer; Value: TBookmarkFlag); override;
    procedure SetFiltered(Value: Boolean); override;
    procedure SetFilterText(const Value: String); override;
    procedure SetRecNo(Value: Integer); override;
    // new methods
    procedure AllocBlobPointers(Buffer: TRecordBuffer);
    procedure CalcFieldOffsets;
    function ColIndexFromField(AField: TField): TColIndex;
    procedure DetectFieldDefs;
    function FilterRecord(Buffer: TRecordBuffer): Boolean;
    procedure FreeBlobPointers(Buffer: TRecordBuffer);
    procedure FreeWorkbook;
    function GetTotalFieldSize: Integer;
    procedure LoadWorksheetToBuffer(Buffer: TRecordBuffer; ARecNo: Integer);
    function LocateRecord(const KeyFields: string; const KeyValues: Variant;
      Options: TLocateOptions; out ARecNo: integer): Boolean;
    procedure ParseFilter(const AFilter: STring);
    procedure SetupAutoInc;
    procedure WriteBufferToWorksheet(Buffer: TRecordBuffer);

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function BookmarkValid(ABookmark: TBookmark): Boolean; override;
    procedure Clear;
    procedure Clear(ClearDefs: Boolean);
    function CompareBookmarks(Bookmark1, Bookmark2: TBookmark): Longint; override;
    function CreateBlobStream(Field: TField; Mode: TBlobStreamMode): TStream; override;
    procedure CreateTable;
    function GetFieldData(Field: TField; Buffer: Pointer): Boolean; override;
    function Locate(const KeyFields: String; const KeyValues: Variant;
      Options: TLocateOptions): boolean; override;
    function Lookup(const Keyfields: String; const KeyValues: Variant;
      const ResultFields: String): Variant; override;
    procedure SetFieldData(Field: TField; Buffer: Pointer); override;

    property Modified: boolean read FModified;

  // This section is to be removed after debugging.
  public
    // to be removed
    property Buffers;
    property BufferCount;

  published
    property AutoFieldDefs: Boolean read FAutoFieldDefs write FAutoFieldDefs default true;
    property FileName: TFileName read FFileName write FFileName;
    property SheetName: String read FSheetName write FSheetName;

    // inherited properties
    property Active;
    property AutoCalcFields;
    property FieldDefs;
    property Filter;
    property Filtered;
    property FilterOptions default [];

    // inherited events
    property AfterCancel;
    property AfterClose;
    property AfterDelete;
    property AfterEdit;
    property AfterInsert;
    property AfterOpen;
    property AfterPost;
    property AfterRefresh;
    property AfterScroll;
    property BeforeCancel;
    property BeforeClose;
    property BeforeDelete;
    property BeforeEdit;
    property BeforeInsert;
    property BeforeOpen;
    property BeforePost;
    property BeforeRefresh;
    property BeforeScroll;
    property OnCalcFields;
    property OnDeleteError;
    property OnEditError;
    property OnFilterRecord;
    property OnNewRecord;
    property OnPostError;
  end;

procedure Register;


implementation

uses
  LazUTF8, LazUTF16, Math, TypInfo, Variants, fpsNumFormat;

procedure Register;
begin
  RegisterComponents('Data Access', [
    TsWorksheetDataset
  ]);
end;


{ Null mask handling

  The null mask is a part of the record buffer which stores in its bits the
  information which fields are NULL. Since all bytes in a new record all bits
  are cleared and a new record has NULL fields the logic must "inverted", i.e.
  a 0-bit means: "field is NULL", and a 1-bit means "field is not NULL". }

{ Clears the information that the field is null by setting the corresponding
  bit in the null mask. }
procedure ClearFieldIsNull(NullMask: PByte; FieldNo: Integer);
var
  n: Integer = 0;
  m: Integer = 0;
begin
  DivMod(FieldNo - 1, 8, n, m);
  inc(NullMask, n);
  // Set the bit to indicate that the field is not NULL.
  NullMask^ := NullMask^ or (1 shl m);
end;

{ Returns true when the field is null, i.e. when its bit in the null mask is not set. }
function GetFieldIsNull(NullMask: PByte; FieldNo: Integer): Boolean;
var
  n: Integer = 0;
  m: Integer = 0;
begin
  DivMod(FieldNo - 1, 8, n, m);
  inc(NullMask, n);
  Result := NullMask^ and (1 shl m) = 0;
end;

{ Clears in the null mask the bit corresponding to FieldNo to indicate that the
  associated field is NULL. }
procedure SetFieldIsNull(NullMask: PByte; FieldNo: Integer);
var
  n: Integer = 0;
  m: Integer = 0;
begin
  DivMod(FieldNo - 1, 8, n, m);
  inc(NullMask, n);
  NullMask^ := nullMask^ and not (1 shl m);
end;


{ TsFieldDef }

constructor TsFieldDef.Create(ACollection: TCollection);
begin
  inherited;
  FColIndex := -1;
end;

constructor TsFieldDef.Create(AOwner: TFieldDefs; const AName: string;
  ADataType: TFieldType; ASize: Integer; ARequired: Boolean; AFieldNo: Longint;
  AColIndex: TColIndex; ACodePage: TSystemCodePage = CP_ACP); overload;
begin
  inherited Create(AOwner, AName, ADataType, ASize, ARequired, AFieldNo, ACodePage);
  FColIndex := AColIndex;
end;

procedure TsFieldDef.Assign(ASource: TPersistent);
begin
  if ASource is TsFieldDef then
    FColIndex := TsFieldDef(ASource).FColIndex;
  inherited Assign(ASource);
end;


{ TsFieldDefs }

class function TsFieldDefs.FieldDefClass: TFieldDefClass;
begin
  Result := TsFieldDef;
end;


{ TsBlobData }

type
  TsBlobData = record
    Data: TBytes;
//    dummy: Int64;
  end;
  PsBlobData = ^TsBlobData;


{ TsBlobStream }

type
  TsBlobStream = class(TMemoryStream)
  private
    FField: TBlobField;
    FDataSet: TsWorksheetDataSet;
    FMode: TBlobStreamMode;
    FModified: Boolean;
    procedure LoadBlobData;
    procedure SaveBlobData;
  public
    constructor Create(Field: TBlobField; Mode: TBlobStreamMode);
    destructor Destroy; override;
    function Read(var Buffer; Count: Longint): Longint; override;
    function Write(const Buffer; Count: Longint): Longint; override;
  end;

constructor TsBlobStream.Create(Field: TBlobField; Mode: TBlobStreamMode);
begin
  inherited Create;
  FField := Field;
  FMode := Mode;
  FDataset := FField.Dataset as TsWorksheetDataset;
  if Mode <> bmWrite then
    LoadBlobData;
end;

destructor TsBlobStream.Destroy;
begin
  if FModified then
    SaveBlobData;
  inherited Destroy;
end;

// Copies the BLOB field data from the active buffer into the stream
procedure TsBlobStream.LoadBlobData;
var
  buffer: TRecordBuffer;
  nullMask: Pointer;
begin
  Self.Size := 0;
  if FDataset.GetActiveBuffer(buffer) then
  begin
    nullMask := FDataset.GetNullMaskPtr(buffer);
    inc(buffer, FDataset.FFieldOffsets[FField.FieldNo-1]);
    Size := 0;
    if not GetFieldIsNull(nullMask, FField.FieldNo) then
      with PsBlobData(buffer)^ do
        Write(Data[0], Length(Data));   // Writes the data into the stream
    Position := 0;
    SaveToFile('test.txt');
  end;
  Position := 0;
end;

function TsBlobStream.Read(var Buffer; Count: LongInt): LongInt;
begin
  Result := inherited Read(Buffer, Count);
end;

// Writes the stream data to the buffer of the BLOB field.
// Take care of the null mask!
procedure TsBlobStream.SaveBlobData;
var
  buffer: TRecordBuffer;
  nullMask: Pointer;
begin
  if FDataset.GetActiveBuffer(buffer) then
  begin
    nullMask := FDataset.GetNullMaskPtr(buffer);
    inc(buffer, FDataset.FFieldOffsets[FField.FieldNo-1]);
    Position := 0;
    if Size = 0 then
      SetFieldIsNull(nullMask, FField.FieldNo)
    else
      with PsBlobData(buffer)^ do
      begin
        SetLength(Data, Size);
        Read(Data[0], Size);  // Reads the stream data to put them into the buffer
        ClearFieldIsNull(nullMask, FField.FieldNo);
      end;
    Position := 0;
  end;
  FModified := false;
end;

function TsBlobStream.Write(const Buffer; Count: LongInt): LongInt;
begin
  Result := inherited Write(Buffer, Count);
  FModified := true;
end;


{ TsWorksheetDataset }

constructor TsWorksheetDataset.Create(AOwner: TComponent);
begin
  inherited;
  FAutoFieldDefs := true;
  FRecordCount := -1;
  FTotalFieldSize := -1;
  FRecordBufferSize := -1;
  FRecNo := -1;
  FAutoIncValue := -1;
  BookmarkSize := SizeOf(TRowIndex);
end;

destructor TsWorksheetDataset.Destroy;
begin
  Close;
  inherited;
end;

procedure TsWorksheetDataset.AllocBlobPointers(Buffer: TRecordBuffer);
var
  i: Integer;
  f: TField;
  offset: Integer;
begin
  for i := 0 to FieldCount-1 do
  begin
    f := Fields[i];
    if f.DataType in [ftMemo{, ftGraphic}] then
    begin
      offset := FFieldOffsets[f.FieldNo-1];
  //    FillChar(PsBlobData(Buffer + offset)^, SizeOf(TsBlobData), 0);
      PsBlobData(Buffer + offset)^.Data := nil;
//      SetLength(PsBlobData(Buffer + offset)^.Data, 0);
    end;
  end;
end;


{ Allocates a buffer for the dataset

  Structure of the TsWorksheetDataset buffer
  +---------------------------------------------------+-----------------------+
  |        field data       | null mask | record info |   calculated fields   |
  +---------------------------------------------------+-----------------------+

  <-------------------- GetRecordSize ----------------> <-- CalcFieldsSize --->
}
function TsWorksheetDataset.AllocRecordBuffer: TRecordBuffer;
var
  n: Integer;
begin
  n := GetRecordSize + CalcFieldsSize;
  GetMem(Result, n);
  FillChar(Result^, n, 0);
  AllocBlobPointers(Result);
end;

{ Returns whether the specified bookmark is valid, i.e. the worksheet row index
  associated with the bookmark cell is between first and last data rows. }
function TsWorksheetDataset.BookmarkValid(ABookmark: TBookmark): Boolean;
var
  bookmarkCell: PCell;
begin
  Result := False;
  if ABookMark = nil then exit;
  bookmarkCell := PPCell(ABookmark)^;
  Result := (bookmarkCell^.Row >= GetFirstDataRowIndex) and 
            (bookmarkCell^.Row <= GetLastDataRowIndex);
end;

procedure TsWorksheetDataset.CalcFieldOffsets;
var
  i: Integer;
  fs: Integer;  // field size
begin
  SetLength(FFieldOffsets, FieldDefs.Count);
  FFieldOffsets[0] := 0;
  for i := 0 to FieldDefs.Count-2 do
  begin
    case FieldDefs[i].DataType of
      ftString, ftFixedChar:
        fs := FieldDefs[i].Size + 1;  // +1 for zero termination
      ftWideString, ftFixedWideChar:
        fs := (FieldDefs[i].Size + 1) * 2;
      ftInteger, ftAutoInc:
        fs := SizeOf(Integer);
      {$IF FPC_FullVersion >= 30202}
      ftByte:
        fs := SizeOf(Byte);
      {$IFEND}
      ftSmallInt:
        fs := SizeOf(SmallInt);
      ftWord:
        fs := SizeOf(Word);
      ftLargeInt:
        fs := Sizeof(LargeInt);
      ftFloat, ftCurrency:
        fs := SizeOf(Double);
      ftDateTime, ftDate, ftTime:
        fs := SizeOf(TDateTime);  // date/time values are TDateTime in the buffer
      ftBoolean:
        fs := SizeOf(WordBool);  // boolean is expected by TBooleanField to be WordBool
      ftMemo:
        fs := SizeOf(TsBlobData);
      else
        DatabaseError(Format('Field data type %s not supported.', [
          GetEnumName(TypeInfo(TFieldType), integer(FieldDefs[i].DataType))
        ]));
    end;
    FFieldOffsets[i+1] := FFieldOffsets[i] + fs;
  end;
end;

procedure TsWorksheetDataset.Clear;
begin
  Clear(true);
end;

procedure TsWorksheetDataset.Clear(ClearDefs: Boolean);
begin
  FRecNo := -1;
  FRecordCount := -1;
  FTotalFieldSize := -1;
  FRecordBufferSize := -1;
  if Active then
    Resync([]);
  if ClearDefs then
  begin
    Close;
    FieldDefs.Clear;
    FTableCreated := false;
  end;
end;

procedure TsWorksheetDataset.ClearCalcFields(Buffer: TRecordBuffer);
begin
  FillChar(Buffer[RecordSize], CalcFieldsSize, 0);
end;

{ Determines the worksheet column index for a specific field }
function TsWorksheetDataset.ColIndexFromField(AField: TField): TColIndex;
var
  fieldDef: TsFieldDef;
begin
  fieldDef := AField.FieldDef as TsFieldDef;
  if fieldDef <> nil then
    Result := fieldDef.ColIndex
  else
    Result := -1;
end;

// Compares two bookmarks (row indices). This tricky handling of nil is
// "borrowed" from TMemDataset
function TsWorksheetDataset.CompareBookmarks(Bookmark1, Bookmark2: TBookmark): Longint;
const
  r: array[Boolean, Boolean] of ShortInt = ((2,-1),(1,0));
var
  cell1, cell2: PCell;
begin
  Result := r[Bookmark1 = nil, Bookmark2 = nil];
  if Result = 2 then
  begin
    cell1 := PPCell(Bookmark1)^;
    cell2 := PPCell(Bookmark2)^;
    Result := Int64(cell1^.Row) - Int64(cell2^.Row);
  end;
end;

function TsWorksheetDataSet.CreateBlobStream(Field: TField;
  Mode: TBlobStreamMode): TStream;
begin
  Result := TsBlobStream.Create(Field as TBlobField, Mode);
end;

{ Creates a new table, i.e. a new empty worksheet based on the given FieldDefs
  The field names are written to the first row of the worksheet. }
procedure TsWorksheetDataset.CreateTable;
var
  i: Integer;
  fd: TsFieldDef;
  noWorkbook: Boolean;
begin
  CheckInactive;
  Clear(false);    // false = do not clear FieldDefs

  if FAutoIncValue < 0 then
    FAutoIncValue := 1;

  noWorkbook := (FWorkbook = nil);
  if noWorkbook then
  begin
    FWorkbook := TsWorkbook.Create;
    FWorkSheet := FWorkbook.AddWorksheet(FSheetName);
  end;

  for i := 0 to FieldDefs.Count-1 do
  begin
    fd := FieldDefs[i] as TsFieldDef;
    FWorksheet.WriteText(0, fd.ColIndex, fd.Name);
  end;
  FWorkbook.WriteToFile(FFileName, true);

  if noWorkbook then
  begin
    FreeAndNil(FWorkbook);
    FWorksheet := nil;
  end;

  FTableCreated := true;
end;

{ Automatic detection of field types and field sizes, as well as the offsets
  for each field in the buffers to be used when accessing records.
  Is called in case of auto-detection from a spreadsheet file (i.e. when
  AutoFieldDefs is true and no other field defs have been defined. }
procedure TsWorksheetDataset.DetectFieldDefs;
var
  r, c: Integer;
  cLast: cardinal;
  cell: PCell;
  fd: TFieldDef;
  fn: String;
  ft: TFieldType;
  fs: Integer;
  isDate, isTime: Boolean;
  fmt: TsCellFormat;
  numFmt: TsNumFormatParams;
begin
  FieldDefs.Clear;

  // Iterate through all columns and collect field defs.
  cLast := FWorksheet.GetLastOccupiedColIndex;
  for c := 0 to cLast do
  begin
    cell := FWorksheet.FindCell(FFirstRow, c);
    if cell = nil then
      Continue;

    // Store field name from cell in FFirstRow
    fn := FWorksheet.ReadAsText(cell);

    // Determine field type: Iterate over rows until first data value is found.
    // The cell content type determines the field type. Iteration stops then.
    for r := GetFirstDataRowIndex to GetLastDataRowIndex do
    begin
      cell := FWorksheet.FindCell(r, c);
      if (cell = nil) then
        continue;
      fmt := FWorkbook.GetCellFormat(FWorksheet.GetEffectiveCellFormatIndex(cell));
      numFmt := FWorkbook.GetNumberFormat(fmt.NumberFormatIndex);
      case cell^.ContentType of
        cctNumber:
          if IsCurrencyFormat(numfmt) then
            ft := ftCurrency
          else
          if (numfmt <> nil) and (CountDecs(numfmt.NumFormatStr) > 0) then
            ft := ftFloat
          else
            ft := ftInteger;    // float will be checked further below
        cctUTF8String:
          ft := ftWideString;
          // Handle text cells as widestring although the worksheet provides then
          // as UTF8. The reason is that a UTF8 field has a datasize of 4*size+1
          // to allow at worst 4-byte code-points. This makes it impossible to
          // control the max text length of a field.
        cctDateTime:
          ft := ftDateTime; // ftDate, ftTime will be checked below
        cctBool:
          ft := ftBoolean;
        else
          continue;
      end;
      break;
    end;

    // Determine field size and distinguish between similar field types
    fs := 0;
    case ft of
      ftWideString:
        begin
          // Find longest text in column...
          for r := GetFirstDataRowIndex to GetLastDataRowIndex do
            fs := Max(fs, Length(FWorksheet.ReadAsText(r, c)));
          if fs > 255 then  // Switch to memo when the strings are "very" long
          begin
            ft := ftMemo;
            fs := 0;
          end else
          if fs > 128 then
            fs := 255
          else
          if fs > 64 then
            fs := 128
          else
          if fs > 32 then
            fs := 64
          else
          if fs > 16 then
            fs := 32
          else
          if fs > 8 then
            fs := 16
          else
          if fs <> 1 then
            fs := 8;
        end;
      ftInteger:    // Distinguish between integer and float
        for r := GetFirstDataRowIndex to GetLastDataRowIndex do
        begin
          cell := FWorksheet.FindCell(r, c);
          if cell = nil then
            continue;
          if (cell^.ContentType = cctNumber) and (frac(cell^.NumberValue) <> 0) then
          begin
            ft := ftFloat;
            break;
          end;
        end;
      ftDateTime:
        begin
          // Determine whether the date/time can be simplified to a pure date or pure time.
          isDate := true;
          isTime := true;
          for r := GetFirstDataRowIndex to GetLastDataRowIndex do
          begin
            cell := FWorksheet.FindCell(r, c);
            if cell = nil then
              continue;
            if frac(cell^.DateTimeValue) <> 0 then isDate := false;  // Non-integer date/time is date
            if (cell^.DateTimeValue > 0) then isTime := false;       // We assume that time is only between 0:00 and 23:59:59.999
            if (not isDate) and (not isTime) then break;
          end;
          if isDate then ft := ftDate;
          if isTime then ft := ftTime;
        end;
      else
        ;
    end;

    // Add FieldDef and set its properties
    TsFieldDef.Create(TsFieldDefs(FieldDefs), FixFieldName(fn), ft, fs,
      false, FieldDefs.Count + 1, c, CP_UTF8);
  end;

  // Determine the offsets at which the field data will begin in the buffer.
  CalcFieldOffsets;
end;

{ Is called before the workbook is opened: checks for filename and sheet name
  as well as file existence. }
procedure TsWorksheetDataset.DoBeforeOpen;
begin
  if (FFileName = '') then
    DatabaseError('Filename not specified.');

  if (FieldDefs.Count = 0) then begin
    if not FileExists(FFileName) then
      DatabaseError('File not found.');
  end;

  if (FSheetName = '') then
    DatabaseError('Worksheet name not specified.');

  inherited;
end;

// Returns the class to be used for FieldDefs. Is overridden to get access
// to the worksheet column index of a field.
class function TsWorksheetDataset.FieldDefsClass : TFieldDefsClass;
begin
  Result := TsFieldDefs;
end;

{ Is called during filtering and returns true when the record who's buffer is
  specified as parameter passes the filter criterions.
  These are determined by the OnFilterRecord event and/or by the Filter property.

  Based on TMemDataset and TBufDataset. }
function TsWorksheetDataset.FilterRecord(Buffer: TRecordBuffer): Boolean;
var
  SaveState: TDatasetState;
begin
  Result := True;

  SaveState := SetTempState(dsFilter);
  try
    FFilterBuffer := Buffer;

    // Check user filter
    if Assigned(OnFilterRecord) then
      OnFilterRecord(Self, Result);

    // Check filter text
    if Result and (Length(Filter) > 0) then
      Result := Boolean(FParser.ExtractFromBuffer(FFilterBuffer)^);
  finally
    RestoreState(SaveState);
  end;
end;

// Removes characters from AText which would make it an invalid fieldname.
function TsWorksheetDataset.FixFieldName(const AText: String): String;
var
  ch: char;
begin
  Result := '';
  for ch in AText do
    if (ch in ['A'..'Z', 'a'..'z', '0'..'9']) then
      Result := Result + ch;
end;

procedure TsWorksheetDataset.FreeBlobPointers(Buffer: TRecordBuffer);
var
  i: Integer;
  f: TField;
  offset: Integer;
begin
  for i := 0 to FieldCount-1 do
  begin
    f := Fields[i];
    if f is TBlobField then
//    if f.DataType in [ftMemo{, ftGraphic}] then
    begin
      offset := FFieldOffsets[f.FieldNo-1];
      PsBlobData(Buffer + offset)^.Data := nil;
//      SetLength(PsBlobData(Buffer + offset)^.Data,0);
//      FillChar(PsBlobData(Buffer + offset)^, SizeOf(TsBlobData), 0);
    end;
  end;
end;

// Frees a record buffer.
procedure TsWorksheetDataset.FreeRecordBuffer(var Buffer: TRecordBuffer);
begin
  FreeBlobPointers(Buffer);
  FreeMem(Buffer);
end;

procedure TsWorksheetDataset.FreeWorkbook;
begin
  FreeAndNil(FWorkbook);
  FWorksheet := nil;
end;

// Returns the active buffer, depending on dataset's state.
// Borrowed from TMemDataset.
function TsWorksheetDataset.GetActiveBuffer(out Buffer: TRecordBuffer): Boolean;
begin
  case State of
    dsEdit,
    dsInsert:
      Buffer := ActiveBuffer;
    dsFilter:
      Buffer := FFilterBuffer;
    dsCalcFields:
      Buffer := CalcBuffer;
    else
      if IsEmpty then
        Buffer := nil
      else
        Buffer := ActiveBuffer;
  end;
  Result := (Buffer <> nil);
end;

{ Returns the pointer to the first cell in the row corresponding to the RecNo
  to be used as a bookmark. }
function TsWorksheetDataset.GetBookmarkCellFromRecNo(ARecNo: Integer): PCell;
var
  row: TRowIndex;
  col: TColIndex;
begin
  row := GetRowIndexFromRecNo(ARecNo);
  col := FWorksheet.GetFirstColIndex;
  Result := FWorksheet.GetCell(row, col);
  // Do not use FindCell here because the returned cell is referenced by the
  // bookmark system.
end;

// Extracts the bookmark from the specified buffer.
procedure TsWorksheetDataset.GetBookmarkData(Buffer: TRecordBuffer; Data: Pointer);
var
  bookmarkCell: PCell;
begin
  if Data <> nil then
  begin
    bookmarkCell := GetRecordInfoPtr(Buffer)^.Bookmark;
    PPCell(Data)^ := bookmarkcell;
  end;
end;

// Extracts the bookmark flag from the specified buffer.
function TsWorksheetDataset.GetBookmarkFlag(Buffer: TRecordBuffer): TBookmarkFlag;
begin
  Result := GetRecordInfoPtr(Buffer)^.BookmarkFlag;
end;

// Determines worksheet row index for the current record.
function TsWorksheetDataset.GetCurrentRowIndex: TRowIndex;
begin
  Result := GetFirstDataRowIndex + FRecNo;
end;

{ Extracts the data value of a specific field from the active buffer and copies
  it to the memory to which Buffer points.
  Returns false when nothing is copied.
  Adapted from TMemDataset. }
function TsWorksheetDataset.GetFieldData(Field: TField; Buffer: Pointer): Boolean;
var
  srcBuffer: TRecordBuffer;
  idx: Integer;
  dt: TDateTime = 0;
  {%H-}dtr: TDateTimeRec;
begin
  Result := GetActiveBuffer(srcBuffer);
  if not Result then
    exit;

  idx := Field.FieldNo - 1;
  if idx >= 0 then
  begin
    Result := not GetFieldIsNull(GetNullMaskPtr(srcBuffer), Field.FieldNo);
    if not Result then
    begin
      if Field = FAutoIncField then
        Move(FAutoIncValue, Buffer^, Field.DataSize);
      exit;
    end;
    if Assigned(Buffer) then
    begin
      inc(srcBuffer, FFieldOffsets[idx]);
      if (Field.DataType in [ftDate, ftTime, ftDateTime]) then
      begin
        // The srcBuffer contains date/time values as TDateTime, but the
        // field expects them to be TDateTimeRec --> convert to TDateTimeRec
        Move(srcBuffer^, dt, SizeOf(TDateTime));
        dtr := DateTimeToDateTimeRec(Field.DataType, dt);
        Move(dtr, Buffer^, SizeOf(TDateTimeRec));
      end else
        // No need to handle BLOB fields here because they always have Buffer=nil
        Move(srcBuffer^, Buffer^, Field.DataSize);
    end;
  end else
  begin  // Calculated, Lookup
    inc(srcBuffer, RecordSize + Field.Offset);
    Result := Boolean(SrcBuffer[0]);
    if Result and Assigned(Buffer) then
      Move(srcBuffer[1], Buffer^, Field.DataSize);
  end;
end;

// Returns the worksheet row index of the record. This is the row
// following the first worksheet row because that is reserved for the column
// titles (field names).
function TsWorksheetDataset.GetFirstDataRowIndex: TRowIndex;
begin
  Result := FFirstRow + 1;   // +1 because the first row contains the column titles.
end;

// Returns the worksheet row index of the record.
function TsWorksheetDataset.GetLastDataRowIndex: TRowIndex;
begin
  Result := FLastRow;
end;

{ Calculates the pointer to the position of the null mask in the buffer.
  The null mask is after the data block. }
function TsWorksheetDataset.GetNullMaskPtr(Buffer: TRecordBuffer): Pointer;
begin
  Result := Buffer;
  inc(Result, GetTotalFieldSize);
end;

// The information whether a field is NULL is stored in the bits of the
// "Null mask". Each bit corresponds to a field.
// Calculates the size of the null mask.
function TsWorksheetDataset.GetNullMaskSize: Integer;
var
  n: Integer;
begin
  n := FieldDefs.Count;
  Result := n div 8 + 1;
end;

// Returns the number of the current record.
function TsWorksheetDataset.GetRecNo: LongInt;
begin
  UpdateCursorPos;
  if (FRecNo < 0) or (RecordCount = 0) or (State = dsInsert) then
    Result := 0
  else
    Result := FRecNo + 1;
end;

function TsWorksheetDataset.GetRecord(Buffer: TRecordBuffer;
  GetMode: TGetMode; DoCheck: Boolean): TGetResult;
var
  accepted: Boolean;
begin
  Result := grOK;
  accepted := false;

  if RecordCount < 1 then
  begin
    Result := grEOF;
    exit;
  end;

  repeat
    case GetMode of
      gmCurrent:
        if (FRecNo >= RecordCount) or (FRecNo < 0) then
          Result := grError;
      gmNext:
        if (FRecNo < RecordCount - 1) then
          inc(FRecNo)
        else
          Result := grEOF;
      gmPrior:
        if (FRecNo > 0) then
          dec(FRecNo)
        else
          Result := grBOF;
    end;

    // Load the data
    if Result = grOK then
    begin
      LoadWorksheetToBuffer(Buffer, FRecNo);
      with GetRecordInfoPtr(Buffer)^ do
      begin
        Bookmark := GetBookmarkCellFromRecNo(FRecNo);
        BookmarkFlag := bfCurrent;
      end;
      GetCalcFields(Buffer);
      if Filtered then
        accepted := FilterRecord(Buffer)    // Filtering
      else
        accepted := true;
      if (GetMode = gmCurrent) and not accepted then
        Result := grError;
    end;
  until (Result <> grOK) or accepted;

  if (Result = grError) and DoCheck then
    DatabaseError('[GetRecord] Invalid record.');
end;

function TsWorksheetDataset.GetRecordCount: LongInt;
begin
  //CheckActive;
  if FRecordCount = -1 then
    FRecordCount := GetLastDataRowIndex - GetFirstDataRowIndex + 1;
  Result := FRecordCount;
end;

// Returns a pointer to the bookmark block inside the given buffer.
function TsWorksheetDataset.GetRecordInfoPtr(Buffer: TRecordBuffer): PsRecordInfo;
begin
  Result := PsRecordInfo(Buffer + GetTotalFieldSize + GetNullMaskSize);
end;

{ Determines the size of the full record buffer:
  - data block: a contiguous field of bytes consisting of the field values
  - null mask: a bit mask storing the information that a field is null
  - Record Info: the bookmark part of the record }
function TsWorksheetDataset.GetRecordSize: Word;
begin
  if FRecordBufferSize = -1 then
    FRecordBufferSize := GetTotalFieldSize + GetNullMaskSize + SizeOf(TsRecordInfo);
  Result := FRecordBufferSize;
end;

function TsWorksheetDataset.GetRowIndexFromRecNo(ARecNo: Integer): TRowIndex;
begin
  Result := GetFirstDataRowIndex + ARecNo;
end;

// Returns the size of the data part in a buffer. This is the sume of all
// field sizes.
function TsWorksheetDataset.GetTotalFieldSize: Integer;
var
  f: TField;
begin
  if FTotalFieldSize = -1 then
  begin
    FTotalFieldSize := 0;
    for f in Fields do
      if f is TBlobField then
        // Blob fields have zero DataSize, but they occupy space in the record buffer.
        FTotalFieldSize := FTotalFieldSize + SizeOf(TsBlobData)
      else
        FTotalFieldSize := FTotalFieldSize + f.DataSize;
  end;
  Result := FTotalFieldSize;
end;

{ Called internally when a record is added. }
procedure TsWorksheetDataset.InternalAddRecord(Buffer: Pointer; DoAppend: Boolean);
var
  row: TRowIndex;
begin
  inc(FLastRow);
  inc(FRecordCount);
  if DoAppend then
  begin
    row := FLastRow;
    SetCurrentRow(row);
  end;
  WriteBufferToWorksheet(Buffer);
  FModified := true;
end;

{ Closes the dataset }
procedure TsWorksheetDataset.InternalClose;
begin
  FIsOpen := false;

  if FModified then begin
    FWorkbook.WriteToFile(FFileName, true);
    FModified := false;
  end;
  FreeWorkbook;
  if FAutoIncValue > -1 then FAUtoIncValue := 1;
  FreeAndNil(FParser);

  if DefaultFields then
    DestroyFields;
  FTotalFieldSize := -1;
  FRecordBufferSize := -1;
  FRecNo := -1;
end;

{ Called internally when a record is deleted.
  Must delete the row from the worksheet. }
procedure TsWorksheetDataset.InternalDelete;
var
  row: TRowIndex;
begin
  if (FRecNo <0) or (FRecNo >= GetRecordCount) then
    exit;

  row := GetRowIndexFromRecNo(FRecNo);
  FWorksheet.DeleteRow(row);
  dec(FRecordCount);
  if FRecordCount = 0 then
    FRecNo := -1
  else
  if FRecNo >= FRecordCount then FRecNo := FRecordCount - 1;
  FModified := true;
end;

{ Moves the cursor to the first record, i.e. the first data row in the worksheet.}
procedure TsWorksheetDataset.InternalFirst;
begin
  FRecNo := -1;
end;

{ Internally, a bookmark is a cell in a worksheet row. }
procedure TsWorksheetDataset.InternalGotoBookmark(ABookmark: Pointer);
var
  bookmarkCell: PCell;
begin
  bookmarkCell := PPCell(ABookmark)^;
  if (bookmarkCell <> nil) and (bookmarkCell^.Row >= GetFirstDataRowIndex) and
    (bookmarkCell^.Row <= GetLastDataRowIndex)
  then
    SetCurrentRow(bookmarkCell^.Row)
  else
    DatabaseError('Bookmark not found.');
end;

{ Initializes the field defs. }
procedure TsWorksheetDataset.InternalInitFieldDefs;
begin
  if FAutoFieldDefs and (FieldDefs.Count = 0) then
    DetectFieldDefs;
  CalcFieldOffsets;
end;

{ Moves the cursor to the last record, the last data row of the worksheet }
procedure TsWorksheetDataset.InternalLast;
begin
  FRecNo := RecordCount;
end;

{ Opens the dataset: Opens the workbook, initializes field defs, creates fields }
procedure TsWorksheetDataset.InternalOpen;
begin
  FWorkbook := TsWorkbook.Create;
  try
    if not FWorkbook.ValidWorksheetName(FSheetName) then
      DatabaseError('"' + FSheetName + '" is not a valid worksheet name.');

    if not FileExists(FFileName) and (not FAutoFieldDefs) and (not FTableCreated) then
    begin
      FWorkSheet := FWorkbook.AddWorksheet(FSheetName);
      CreateTable;
    end else
    begin
      FWorkbook.ReadFromFile(FFileName);
      FWorksheet := FWorkbook.GetWorksheetByName(FSheetName);
      if FWorksheet = nil then
        DatabaseError('Worksheet not found.');
    end;

    FFirstRow := FWorksheet.GetFirstRowIndex(true);
    FLastRow := FWorksheet.GetLastOccupiedRowIndex;
    FRecordCount := -1;
    FTotalFieldSize := -1;
    FRecordBufferSize := -1;

    InternalInitFieldDefs;
    if DefaultFields then
      CreateFields;
    BindFields(True);  // Computes CalcFieldsSize
    GetTotalFieldSize;
    GetRecordSize;
    FRecNo := -1;

    SetupAutoInc;
    FModified := false;

    FIsOpen := true;
  except
    on E: Exception do
    begin
      FreeWorkbook;
      DatabaseError('Error opening workbook: ' + E.Message);
    end;
  end;
end;

{ Called inernally when a record is posted. }
procedure TsWorksheetDataset.InternalPost;
begin
  CheckActive;
  if not (State in [dsEdit, dsInsert]) then
    Exit;
  inherited InternalPost;
  if (State=dsEdit) then
    WriteBufferToWorksheet(ActiveBuffer)
  else
  begin
    if Assigned(FAutoIncField) then
    begin
      FAutoIncField.AsInteger := FAutoIncValue;
      inc(FAutoIncValue);
    end;
    InternalAddRecord(ActiveBuffer, True);
  end;
end;

{ Reinitializes a buffer which has been allocated previously
  -> zero out everything
  In this step the NullMask is erased, and this means that all fields are null.

  We cannot just fill the buffer with 0s since that would overwrite our BLOB
  pointers. Therefore we free the blob pointers first, then fill the buffer
  with zeros, then reallocate the blob pointers }
procedure TsWorksheetDataset.InternalInitRecord(Buffer: TRecordBuffer);
begin
  FreeBlobPointers(Buffer);
  FillChar(Buffer^, FRecordBufferSize, 0);
  AllocBlobPointers(Buffer);
end;

{ Sets the database cursor to the record specified by the given buffer. We
  extract here the bookmark associated with the buffer and go to this bookmark. }
procedure TsWorksheetDataset.InternalSetToRecord(Buffer: TRecordBuffer);
var
  bookmarkCell: PCell;
begin
  bookmarkCell := GetRecordInfoPtr(Buffer)^.Bookmark;
  InternalGotoBookmark(@bookmarkCell);
end;

function TsWorksheetDataset.IsCursorOpen: boolean;
begin
  Result := FIsOpen;
end;

{ Reads the cells data of the current worksheet row and
  copies them to the buffer. }
procedure TsWorksheetDataset.LoadWorksheetToBuffer(Buffer: TRecordBuffer;
  ARecNo: Integer);
var
  field: TField;
  row: TRowIndex;
  col: TColIndex;
  cell: PCell;
  s: String;
  ws: WideString;
  {%H-}i: Integer;
  {%H-}si: SmallInt;
  {%H-}b: Byte;
  {%H-}w: word;
  {%H-}li: LargeInt;
  {%H-}wb: WordBool;
  nullMask: Pointer;
  maxLen: Integer;
  fs: Integer;
begin
  nullMask := GetNullMaskPtr(Buffer);
  row := GetRowIndexFromRecNo(ARecNo);
  for field in Fields do
  begin
    col := ColIndexFromField(field);
    if col = -1 then  // this happens for calculated fields.
      continue;
    // Find the cell at the column and row. BUT: For bookmark support, we need
    // a cell even when there is none. So: Find the cell by calling GetCell
    // which adds a blank cell in such a case.
    cell := FWorksheet.GetCell(row, col);
    ClearFieldIsNull(nullMask, field.FieldNo);
    if field is TBlobField then
      // BLOB fields have zero DataSize although they occupy space in the buffer
      fs := SizeOf(TsBlobData)
    else
      fs := field.DataSize;
    case cell^.ContentType of
      cctUTF8String:
        begin
          s := FWorksheet.ReadAsText(cell);
          if s = '' then
            SetFieldIsNull(nullMask, field.FieldNo)
          else
          if field.DataType = ftMemo then
          begin
            with PsBlobData(Buffer)^ do
            begin
              SetLength(Data, Length(s));
              Move(s[1], Data[0], Length(s));
            end;
          end else
          if field.DataType in [ftWideString, ftFixedWideChar] then
          begin
            maxLen := (field.DataSize - 2) div 2;
            ws := UTF16Copy(UTF8Decode(s), 1, maxLen) + #0#0;
            Move(ws[1], Buffer^, Length(ws)*2);
          end else
          begin
            maxLen := field.DataSize - 1;
            s := UTF8Copy(s, 1, maxLen) + #0;
            Move(s[1], Buffer^, Length(s));
          end;
        end;
      cctNumber:
        case field.DataType of
          ftFloat:
            Move(cell^.NumberValue, Buffer^, SizeOf(cell^.NumberValue));
          ftCurrency:
            Move(cell^.NumberValue, Buffer^, SizeOf(cell^.Numbervalue));
          ftInteger, ftAutoInc:
            begin
              i := Round(cell^.NumberValue);
              Move(i, Buffer^, SizeOf(i));
              if field.DataType = ftAutoInc then
                FAutoIncField := TAutoIncField(field);
            end;
          {$IF FPC_FullVersion >= 30202}
          ftByte:
            begin
              b := byte(round(cell^.NumberValue));
              Move(b, Buffer^, SizeOf(b));
            end;
          {$IFEND}
          ftSmallInt:
            begin
              si := SmallInt(round(cell^.NumberValue));
              Move(si, Buffer^, SizeOf(si));
            end;
          ftWord:
            begin
              w := word(round(cell^.NumberValue));
              Move(w, Buffer^, SizeOf(w));
            end;
          ftLargeInt:
            begin
              li := LargeInt(round(cell^.NumberValue));
              Move(li, Buffer^, SizeOf(li));
            end;
          ftString, ftFixedChar:
            begin
              s := FWorksheet.ReadAsText(cell) + #0;
              Move(s[1], Buffer^, Length(s));
            end;
          else
            ;
        end;
      cctDateTime:
        // TDataset handles date/time value as TDateTimeRec but expects them
        // to be TDateTime in the buffer. How strange!
        Move(cell^.DateTimeValue, Buffer^, SizeOf(TDateTime));
      cctBool:
        begin
          wb := cell^.BoolValue;    // Boolean field stores value as wordbool
          Move(wb, Buffer^, SizeOf(wb));
        end;
      cctEmpty:
        SetFieldIsNull(nullMask, field.FieldNo);
      else
        ;
    end;
    inc(Buffer, fs);
  end;
end;

{ Searches the first record for which the fields specified by Keyfields
  (semicolon-separated list) have the values defined in KeyValues.
  Returns false, when no such record is found.
  Code from TMemDataset. }
function TsWorksheetDataset.Locate(const KeyFields: string;
  const KeyValues: Variant; Options: TLocateOptions): boolean;
var
  ARecNo: integer;
begin
  // Call inherited to make sure the dataset is bi-directional
  Result := inherited;
  CheckActive;

  Result := LocateRecord(KeyFields, KeyValues, Options, ARecNo);
  if Result then begin
    // TODO: generate scroll events if matched record is found
    FRecNo := ARecNo;
    Resync([]);
  end;
end;

{ Helper function for locating records.
  Taken from TMemDataset: This implements a simple search from record to record.
  To do: introduce an index for faster searching. }
function TsWorksheetDataset.LocateRecord(
  const KeyFields: string;
  const KeyValues: Variant; Options: TLocateOptions;
  out ARecNo: integer): Boolean;
var
  SaveState: TDataSetState;
  lKeyFields: TList;
  Matched: boolean;
  AKeyValues: variant;
  i: integer;
  field: TField;
  s1,s2: String;
begin
  Result := false;
  SaveState := SetTempState(dsFilter);
  FFilterBuffer := TempBuffer;
  lKeyFields := TList.Create;
  try
    GetFieldList(lKeyFields, KeyFields);
    if VarArrayDimCount(KeyValues) = 0 then
    begin
      Matched := lKeyFields.Count = 1;
      AKeyValues := VarArrayOf([KeyValues]);
    end else
    if VarArrayDimCount(KeyValues) = 1 then
    begin
      Matched := VarArrayHighBound(KeyValues,1) + 1 = lKeyFields.Count;
      AKeyValues := KeyValues;
    end
    else
      Matched := false;

    if Matched then
    begin
      ARecNo := 0;
      while ARecNo < RecordCount do
      begin
        LoadWorksheetToBuffer(FFilterBuffer, ARecNo);
        if Filtered then
          Result := FilterRecord(FFilterBuffer)
        else
          Result := true;
        // compare field by field
        i := 0;
        while Result and (i < lKeyFields.Count) do
        begin
          field := TField(lKeyFields[i]);
          // string fields
          if field.DataType in [ftString, ftFixedChar] then
          begin
            {$IF FPC_FullVersion >= 30200}
            if TStringField(field).CodePage=CP_UTF8 then
            begin
              s1 := field.AsUTF8String;
              s2 := UTF8Encode(VarToUnicodeStr(AKeyValues[i]));
            end else
            {$IFEND}
            begin
              s1 := field.AsString;
              s2 := VarToStr(AKeyValues[i]);
            end;
            if loPartialKey in Options then
              s1 := copy(s1, 1, length(s2));
            if loCaseInsensitive in Options then
              Result := AnsiCompareText(s1, s2)=0
            else
              Result := s1=s2;
          end
          // all other fields
          else
            Result := (field.Value=AKeyValues[i]);
          inc(i);
        end;
        if Result then
          break;
        inc(ARecNo);
      end;
    end;
  finally
    lKeyFields.Free;
    RestoreState(SaveState);
  end;
end;

{ Searches the first record for which the fields specified by KeyFields
  (semicolon-separated list of field names) have the values defined in KeyValues.
  Returns the field values of the ResultFields (a semicolon-separated list of field
  names), or NULL if there is no match.
  Code from TMemDataset. }
function TsWorksheetDataset.Lookup(const KeyFields: string; const KeyValues: Variant;
  const ResultFields: string): Variant;
var
  ARecNo: integer;
  SaveState: TDataSetState;
begin
  if LocateRecord(KeyFields, KeyValues, [], ARecNo) then
  begin
    SaveState := SetTempState(dsCalcFields);
    try
      // FFilterBuffer contains found record
      CalculateFields(FFilterBuffer); // CalcBuffer is set to FFilterBuffer
      Result := FieldValues[ResultFields];
    finally
      RestoreState(SaveState);
    end;
  end
  else
    Result := Null;
end;

// from TBufDataset
procedure TsWorksheetDataset.ParseFilter(const AFilter: string);
begin
  // parser created?
  if Length(AFilter) > 0 then
  begin
    if (FParser = nil) and IsCursorOpen then
      FParser := TBufDatasetParser.Create(Self);
    // is there a parser now?
    if FParser <> nil then
    begin
      // set options
      FParser.PartialMatch := not (foNoPartialCompare in FilterOptions);
      FParser.CaseInsensitive := foCaseInsensitive in FilterOptions;
      // parse expression
      FParser.ParseExpression(AFilter);
    end;
  end;
end;

procedure TsWorksheetDataset.SetBookmarkData(Buffer: TRecordBuffer; Data: Pointer);
begin
  if Data <> nil then
    GetRecordInfoPtr(Buffer)^.Bookmark := PPCell(Data)^
  else
    GetRecordInfoPtr(Buffer)^.Bookmark := nil;
end;

procedure TsWorksheetDataset.SetBookmarkFlag(Buffer: TRecordBuffer;
  Value: TBookmarkFlag);
begin
  GetRecordInfoPtr(Buffer)^.BookmarkFlag := Value;
end;

procedure TsWorksheetDataset.SetCurrentRow(ARow: TRowIndex);
begin
  FRecNo := ARow - GetFirstDataRowIndex;
end;

{ Copies the data to which Buffer points to the position in the active buffer
  which belongs to the specified field.
  Adapted from TMemDataset. }
procedure TsWorksheetDataset.SetFieldData(Field: TField; Buffer: Pointer);
var
  destBuffer: TRecordBuffer;
  idx: Integer;
  fsize: Integer;
  {%H-}dt: TDateTime;
  dtr: TDateTimeRec;
begin
  if not GetActiveBuffer(destBuffer) then
    exit;

  idx := Field.FieldNo - 1;
  if idx >= 0 then
  begin
    if State in [dsEdit, dsInsert, dsNewValue] then
      Field.Validate(Buffer);
    if Buffer = nil then
      SetFieldIsNull(GetNullMaskPtr(destBuffer), Field.FieldNo)
    else
    begin
      ClearFieldIsNull(GetNullMaskPtr(destBuffer), Field.FieldNo);
      inc(destBuffer, FFieldOffsets[idx]);
      if Field.DataType in [ftDate, ftTime, ftDateTime] then
      begin
        // Special treatment for date/time values: TDataset expects them
        // to be TDateTime in the destBuffer, but to be TDateTimeRec in the
        // input Buffer.
        dtr := Default(TDateTimeRec);
        Move(Buffer^, dtr, SizeOf(dtr));
        dt := DateTimeRecToDateTime(Field.DataType, dtr);
        Move(dt, destBuffer^, SizeOf(dt));
      end else
      begin
        fsize := Field.DataSize;
        if Field.DataType in [ftString, ftFixedChar] then
          dec(fSize);  // Do not move terminating 0 which is included in DataSize
        if Field.DataType in [ftWideString, ftFixedWideChar] then
          dec(fSize, 2);
        Move(Buffer^, destBuffer^, fsize);
      end;
    end;
  end else
  begin  // Calculated, Lookup
    inc(destBuffer, RecordSize + Field.Offset);
    Boolean(destBuffer[0]) := Buffer <> nil;
    if Assigned(Buffer) then
      Move(Buffer^, DestBuffer[1], Field.DataSize);
  end;

  if not (State in [dsCalcFields, dsFilter, dsNewValue]) then
    DataEvent(deFieldChange, PtrInt(Field));
end;

// From TBufDataset
procedure TsWorksheetDataset.SetFiltered(Value: Boolean);
begin
  if Value = Filtered then
    exit;

  // Pass on to ancestor
  inherited;

  // Only refresh if active
  if IsCursorOpen then
    Resync([]);
end;

// From TBufDataset
procedure TsWorksheetDataset.SetFilterText(const Value: string);
begin
  if Value = Filter then
    exit;

  // Parse
  ParseFilter(Value);

  // Call dataset method
  inherited;

  // Refilter dataset if filtered
  if IsCursorOpen and Filtered then Resync([]);
end;

procedure TsWorksheetDataset.SetRecNo(Value: Integer);
begin
  CheckBrowseMode;
  if (Value >= 1) and (Value <= RecordCount) then
  begin
    FRecNo := Value-1;
    Resync([]);
  end;
end;

procedure TsWorksheetDataset.SetupAutoInc;
var
  f: TField;
  c: TColIndex;
  r: Integer;
  mx: Integer;
  cell: PCell;
begin
  // Search for autoinc field
  FAutoIncField := nil;
  FAutoIncValue := -1;
  for f in Fields do
    if f is TAutoIncField then
    begin
      FAutoIncField := TAutoIncField(f);
      break;
    end;

  if FAutoIncField = nil then
    exit;

  mx := -MaxInt;
  c := ColIndexFromField(f);
  r := GetFirstDataRowIndex;
  for r := GetFirstDataRowIndex to FLastRow do
  begin
    cell := FWorksheet.FindCell(r, c);
    if cell <> nil then
    begin
      if (cell^.ContentType <> cctNumber) then
        DatabaseError('AutoInc field must be a assigned to numeric cells.');
      mx := Max(mx, round(FWorksheet.ReadAsNumber(cell)));
    end;
    FAutoIncValue := mx + 1;
  end;
end;


procedure TsWorksheetDataset.WriteBufferToWorksheet(Buffer: TRecordBuffer);
var
  row: TRowIndex;
  col: TColIndex;
  cell: PCell;
  field: TField;
  P: Pointer;
  s: String = '';
  ws: WideString = '';
begin
  row := GetCurrentRowIndex;
  P := Buffer;
  for field in Fields do begin
    col := ColIndexFromField(field);
    cell := FWorksheet.FindCell(row, col);
    if GetFieldIsNull(GetNullMaskPtr(Buffer), field.FieldNo) then
      FWorksheet.WriteBlank(cell)
    else
    begin
      P := Buffer + FFieldOffsets[field.FieldNo-1];
      cell := FWorksheet.GetCell(row, col);
      case field.DataType of
        ftFloat:
          if (TFloatField(field).Precision >= 15) or (TFloatField(field).Precision < 0) then
            FWorksheet.WriteNumber(cell, PDouble(P)^, nfGeneral)
          else
            FWorksheet.WriteNumber(cell, PDouble(P)^, nfFixed, TFloatField(field).Precision);
        ftCurrency:
          FWorksheet.WriteCurrency(cell, PDouble(P)^, nfCurrency, 2);
        ftInteger, ftAutoInc:
          FWorksheet.WriteNumber(cell, PInteger(P)^);
        {$IF FPC_FullVersion >= 30202}
        ftByte:
          FWorksheet.WriteNumber(cell, PByte(P)^);
        {$IFEND}
        ftSmallInt:
          FWorksheet.WriteNumber(cell, PSmallInt(P)^);
        ftWord:
          FWorksheet.WriteNumber(cell, PWord(P)^);
        ftLargeInt:
          FWorksheet.WriteNumber(cell, PLargeInt(P)^);
        ftDateTime:
          FWorksheet.WriteDateTime(cell, PDateTime(P)^, nfShortDateTime);
        ftDate:
          FWorksheet.WriteDateTime(Cell, PDateTime(P)^, nfShortDate);
        ftTime:
          FWorksheet.WriteDateTime(cell, PDateTime(P)^, nfLongTime);
        ftBoolean:
          FWorksheet.WriteBoolValue(cell, PWordBool(P)^);
        ftString, ftFixedChar:
          FWorksheet.WriteText(cell, StrPas(PChar(P)));
        ftWideString, ftFixedWideChar:
          begin
            Setlength(ws, StrLen(PWideChar(P)));
            Move(P^, ws[1], Length(ws)*2);
            FWorksheet.WriteText(cell, UTF8Encode(ws));
          end;
        ftMemo:
          begin
            SetLength(s, Length(PsBlobData(P)^.Data));
            if Length(PsBlobData(P)^.Data) > 0 then
              Move(PsBlobData(P)^.Data[0], s[1], Length(s));
            FWorksheet.WriteText(cell, s);
          end;
        else
          ;
      end;
    end;
  end;
  FModified := true;
end;


end.

