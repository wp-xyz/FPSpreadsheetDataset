{@@ ----------------------------------------------------------------------------
  Unit **fpsDataset** implements a TDataset based on spreadsheet data.
  This way spreadsheets can be accessed in a database-like manner.
  Of course, it is required that all cells in a column have the same type.

  AUTHORS: Werner Pamler

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
           distribution, for details about the license.

  Documentation used:
  * https://www.delphipower.xyz/guide_8/building_custom_datasets.html
  * http://etutorials.org/Programming/mastering+delphi+7/Part+III+Delphi+Database-Oriented+Architectures/Chapter+17+Writing+Database+Components/Building+Custom+Datasets/
  * http://216.19.73.24/articles/customds.asp
  * https://delphi.cjcsoft.net/viewthread.php?tid=44220

  Much of the code is adapted from TMemDataset.

  Current status (Aug 27, 2021):
  * Field defs: determined automatically from file
  * Field defs defined by user: working (requires AutoFieldDefs = false)
  * Fields: working
  * Field types: ftFloat, ftInteger, ftDateTime, ftDate, ftTime, ftString, ftBoolean
  * Calculated fields: in code, but not tested, yet.
  * Persistent fields: to be done
  * Locate: working
  * Lookup: working
  * Edit, Delete, Insert: working, Post and Cancel ok
  * NULL fields: working
  * GetBookmark, GotoBookmark: working
  * Filter: only by OnFilter event, working.

  ' Field defs: Required, Unique etc not supported ATM.
  * Indexes: not implemented
  * Sorting: not implemented

  Issues
  * Bookmark moves up by 1 when a record is inserted before bookmark
-------------------------------------------------------------------------------}

unit fpsDataset;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, DB,
  fpSpreadsheet, fpsTypes, fpsUtils;

type
  TRowIndex = Int64;
  TColIndex = Int64;

  { TsFieldDef }
  TsFieldDef = class(TFieldDef)
  private
    FColumn: TColIndex;
  public
    constructor Create(ACollection: TCollection); override;
  published
    property Column: TColIndex read FColumn write FColumn default -1;
  end;

  { TsFieldDefs }
  TsFieldDefs = class(TFieldDefs)
  protected
    class function FieldDefClass : TFieldDefClass; override;
  end;

  { TsRecordInfo }
  TsRecordInfo = record
    Bookmark: LongInt;
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
    FDataSize: Integer;              // Total size of the field data
    FFieldOffsets: array of Integer; // Offset to field start in buffer
    FModified: Boolean;              // Flag to show that workbook needs saving
    FFilterBuffer: TRecordBuffer;
    FTableCreated: boolean;
    FAutoFieldDefs: Boolean;
    FIsOpen: boolean;
  private
    function FixFieldName(const AText: String): String;
    function GetActiveBuffer(out Buffer: TRecordBuffer): Boolean;
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
    procedure SetRecNo(Value: Integer); override;
    // new methods
    procedure CalcFieldOffsets;
    function ColIndexFromField(AField: TField): TColIndex;
    procedure DetectFieldDefs;
    function FilterRecord(Buffer: TRecordBuffer): Boolean;
    procedure FreeWorkbook;
    function GetDataSize: Integer;
    procedure LoadRecordToBuffer(Buffer: TRecordBuffer; ARecNo: Integer);
    function LocateRecord(const KeyFields: string; const KeyValues: Variant;
      Options: TLocateOptions; out ARecNo: integer): Boolean;
    procedure SetFilterText(const {%H-}AValue: String); override;
    procedure WriteBufferToWorksheet(Buffer: TRecordBuffer);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function BookmarkValid(ABookmark: TBookmark): Boolean; override;
    procedure Clear;
    procedure Clear(ClearDefs: Boolean);
    function CompareBookmarks(Bookmark1, Bookmark2: TBookmark): Longint; override;
    procedure CreateTable;
    function GetFieldData(Field: TField; Buffer: Pointer): Boolean; override;
    function Locate(const KeyFields: String; const KeyValues: Variant;
      Options: TLocateOptions): boolean; override;
    function Lookup(const Keyfields: String; const KeyValues: Variant;
      const ResultFields: String): Variant; override;
    procedure SetFieldData(Field: TField; Buffer: Pointer); override;
    property Filter; unimplemented;  // Use OnFilter instead
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

    property Active;
    property Filtered;
    property AfterCancel;
    property AfterClose;
    property AfterOpen;
    property AfterPost;
    property AfterScroll;
    property BeforeCancel;
    property BeforeClose;
    property BeforeOpen;
    property BeforePost;
    property BeforeScroll;
    property OnFilterRecord;
    property OnPostError;
  end;

implementation

uses
  Math, Variants;

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
  FColumn := -1;
end;

{ TsFieldDefs }

class function TsFieldDefs.FieldDefClass: TFieldDefClass;
begin
  Result := TsFieldDef;
end;

{ TsWorksheetDataset }

constructor TsWorksheetDataset.Create(AOwner: TComponent);
begin
  inherited;
  FAutoFieldDefs := true;
  FRecordCount := -1;
  FDataSize := -1;
  FRecordBufferSize := -1;
  FRecNo := -1;
  BookmarkSize := SizeOf(Longint);
end;

destructor TsWorksheetDataset.Destroy;
begin
  Close;
  inherited;
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
end;

{ Returns whether the specified bookmark is valid, i.e. the worksheet row index
  associated with the bookmark is between first and last data rows. }
function TsWorksheetDataset.BookmarkValid(ABookmark: TBookmark): Boolean;
var
  reqBookmark: Integer;
begin
  Result := False;
  if ABookMark=nil then exit;
  reqBookmark := PInteger(ABookmark)^;
  Result := (reqBookmark >= 0) and (reqBookmark <= GetRecordCount);
end;

procedure TsWorksheetDataset.CalcFieldOffsets;
var
  i: Integer;
  fs: Integer;  // field size
begin
  SetLength(FFieldOffsets, FieldDefs.Count);
  FFieldOffsets[0] := 0;
  for i := 1 to FieldDefs.Count-1 do
  begin
    case FieldDefs[i-1].DataType of
      ftString: fs := FieldDefs[i-1].Size + 1;  // +1 for zero termination
      ftInteger: fs := SizeOf(Integer);
      ftFloat: fs := SizeOf(Double);
      ftDateTime, ftDate, ftTime: fs := SizeOf(TDateTime);  // date/time values are TDateTime in the buffer
      ftBoolean: fs := SizeOf(WordBool);  // boolean is expected by TBooleanField to be WordBool
      else ;
    end;
    FFieldOffsets[i] := FFieldOffsets[i-1] + fs;
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
  FDataSize := -1;
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
  Result := fieldDef.Column;
end;

// Compares two bookmarks (row indices). This tricky handling of nil is
// "borrowed" from TMemDataset
function TsWorksheetDataset.CompareBookmarks(Bookmark1, Bookmark2: TBookmark): Longint;
const
  r: array[Boolean, Boolean] of ShortInt = ((2,-1),(1,0));
begin
  Result := r[Bookmark1=nil, Bookmark2=nil];
  if Result = 2 then
    Result := PInteger(Bookmark1)^ - PInteger(Bookmark2)^;
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

  noWorkbook := (FWorkbook = nil);
  if noWorkbook then
  begin
    FWorkbook := TsWorkbook.Create;
    FWorkSheet := FWorkbook.AddWorksheet(FSheetName);
  end;

  for i := 0 to FieldDefs.Count-1 do
  begin
    fd := FieldDefs[i] as TsFieldDef;
    FWorksheet.WriteText(0, fd.Column, fd.Name);
  end;
  FWorkbook.WriteToFile(FFileName, true);

  if noWorkbook then
  begin
    FreeAndNil(FWorkbook);
    FWorksheet := nil;
  end;

  FTableCreated := true;
end;

// Determines the offsets for each field in the buffers to be used when
// accessing records.
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
      case cell^.ContentType of
        cctNumber: ft := ftInteger;    // float will be checked below
        cctUTF8String: ft := ftString;
        cctDateTime: ft := ftDateTime; // ftDate, ftTime will be checked below
        cctBool: ft := ftBoolean;
        else continue;
      end;
      break;
    end;

    // Determine field size
    fs := 0;
    case ft of
      ftString:
        begin
          // Find longest text in column...
          for r := GetFirstDataRowIndex to GetLastDataRowIndex do
            fs := Max(fs, Length(FWorksheet.ReadAsText(r, c)));
          // ... and round it up to a multiple of 10 for edition ---> VarChars to be introduced later!
          fs := (fs div 10) * 10 + 10;
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

    // Add field def and set its properties
    fd := FieldDefs.AddFieldDef;
    fd.Name := FixFieldName(fn);
    fd.DataType := ft;
    fd.Size := fs;
    TsFieldDef(fd).Column := c;
  end;

  // Determine the offsets at which the field data will begin in the buffer.
  CalcFieldOffsets;
end;

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

function TsWorksheetDataset.FilterRecord(Buffer: TRecordBuffer): Boolean;
var
  SaveState: TDatasetState;
begin
  Result := True;
  if not Assigned(OnFilterRecord) then
    Exit;
  SaveState := SetTempState(dsFilter);
  try
    FFilterBuffer := Buffer;
    OnFilterRecord(Self, Result);
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

// Frees a record buffer.
procedure TsWorksheetDataset.FreeRecordBuffer(var Buffer: TRecordBuffer);
begin
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

// Extracts the bookmark (worksheet row index) from the specified buffer.
procedure TsWorksheetDataset.GetBookmarkData(Buffer: TRecordBuffer; Data: Pointer);
begin
  if Data <> nil then
    PInteger(Data)^ := GetRecordInfoPtr(Buffer)^.Bookmark;
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

// Returns the size of the data part in a buffer. This is the sume of all
// field sizes.
function TsWorksheetDataset.GetDataSize: Integer;
var
  f: TField;
begin
  if FDataSize = -1 then
  begin
    FDataSize := 0;
    for f in Fields do
      FDataSize := FDataSize + f.DataSize;
  end;
  Result := FDataSize;
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
    if Result and Assigned(Buffer) then
    begin
      inc(srcBuffer, FFieldOffsets[idx]);
      if (Field.DataType in [ftDate, ftTime, ftDateTime]) then
      begin
        // The srcBuffer contains date/time values as TDateTime, but
        // TDataset expects them to be TDateTimeRec --> convert to TDateTimeRec
        Move(srcBuffer^, dt, SizeOf(TDateTime));
        dtr := DateTimeToDateTimeRec(Field.DataType, dt);
        Move(dtr, Buffer^, SizeOf(TDateTimeRec));
      end else begin
        Move(srcBuffer^, Buffer^, Field.DataSize);
      end;
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
  inc(Result, GetDataSize);
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
      LoadRecordToBuffer(Buffer, FRecNo);
      with GetRecordInfoPtr(Buffer)^ do
      begin
        Bookmark := FRecNo;
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
  Result := PsRecordInfo(Buffer + GetDataSize + GetNullMaskSize);
end;

{ Determines the size of the full record buffer:
  - data block: a contiguous field of bytes consisting of the field values
  - null mask: a bit mask storing the information that a field is null
  - Record Info: the bookmark part of the record }
function TsWorksheetDataset.GetRecordSize: Word;
begin
  if FRecordBufferSize = -1 then
    FRecordBufferSize := GetDataSize + GetNullMaskSize + SizeOf(TsRecordInfo);
  Result := FRecordBufferSize;
end;

function TsWorksheetDataset.GetRowIndexFromRecNo(ARecNo: Integer): TRowIndex;
begin
  Result := GetFirstDataRowIndex + ARecNo;
end;

procedure TsWorksheetDataset.InternalAddRecord(Buffer: Pointer; DoAppend: Boolean);
var
  row: TRowIndex;
begin
  row := GetRowIndexFromRecNo(FRecNo);
  FWorksheet.InsertRow(row);
  inc(FLastRow);
  Inc(FRecordCount);
  WriteBufferToWorksheet(Buffer);
  FModified := true;
end;

// Closes the dataset
procedure TsWorksheetDataset.InternalClose;
begin
  FIsOpen := false;

  if FModified then begin
    FWorkbook.WriteToFile(FFileName, true);
    FModified := false;
  end;
  FreeWorkbook;

  if DefaultFields then
    DestroyFields;
  FDataSize := -1;
  FRecordBufferSize := -1;
  FRecNo := -1;
end;

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

// Moves the cursor to the first record, the first data row in the worksheet
procedure TsWorksheetDataset.InternalFirst;
begin
  FRecNo := -1;
end;

// Internally, a bookmark is the row index of the worksheet.
procedure TsWorksheetDataset.InternalGotoBookmark(ABookmark: Pointer);
var
  reqBookmark: Integer;
begin
  reqBookmark := PInteger(ABookmark)^;
  if (reqBookmark >= 0) and (reqBookmark <= GetRecordCount) then
    FRecNo := reqBookmark
  else
    DatabaseError('Bookmark not found.');
end;

// Initializes the field defs.
procedure TsWorksheetDataset.InternalInitFieldDefs;
begin
  if FAutoFieldDefs and (FieldDefs.Count = 0) then
    DetectFieldDefs;
  CalcFieldOffsets;
end;

// Moves the cursor to the last record, the last data row of the worksheet
procedure TsWorksheetDataset.InternalLast;
begin
  FRecNo := RecordCount;
end;

// Opens the dataset: Opens the workbook, initialized field defs, creates fields
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
    FDataSize := -1;
    FRecordBufferSize := -1;

    InternalInitFieldDefs;
    if DefaultFields then
      CreateFields;
    BindFields(True);  // Computes CalcFieldsSize
    GetDataSize;
    GetRecordSize;
    FRecNo := -1;
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

procedure TsWorksheetDataset.InternalPost;
begin
  CheckActive;
  if not (State in [dsEdit, dsInsert]) then
    Exit;
  inherited InternalPost;
  if (State=dsEdit) then
    WriteBufferToWorksheet(ActiveBuffer)
  else
    InternalAddRecord(ActiveBuffer, True);
end;

// Reinitializes a buffer which has been allocated previously -> zero out everything
procedure TsWorksheetDataset.InternalInitRecord(Buffer: TRecordBuffer);
begin
  FillChar(Buffer^, FRecordBufferSize, 0);
end;

// Sets the database cursor to the record specified by the given buffer.
// We extract here the bookmark associated with the buffer and go to this
// bookmark.
procedure TsWorksheetDataset.InternalSetToRecord(Buffer: TRecordBuffer);
var
  reqBookmark: Integer;
begin
  reqBookmark := GetRecordInfoPtr(Buffer)^.Bookmark;
  InternalGotoBookmark(@reqBookmark);
end;

function TsWorksheetDataset.IsCursorOpen: boolean;
begin
  Result := FIsOpen;
end;

// Reads the cells data of the current worksheet row
// and copies them to the buffer.
procedure TsWorksheetDataset.LoadRecordToBuffer(Buffer: TRecordBuffer; ARecNo: Integer);
var
  field: TField;
  row: TRowIndex;
  col: TColIndex;
  cell: PCell;
  {%H-}i: Integer;
  s: String;
  {%H-}b: WordBool;
  bufferStart: TRecordBuffer;
begin
  bufferStart := Buffer;
  //Q := Buffer;

  row := GetRowIndexFromRecNo(ARecNo);
  for field in Fields do
  begin
    col := ColIndexFromField(field);
    cell := FWorksheet.FindCell(row, col);
    if cell = nil then
      SetFieldIsNull(GetNullMaskPtr(bufferStart), field.FieldNo)
    else
    begin
      ClearFieldIsNull(GetNullMaskPtr(bufferStart), field.FieldNo);
      case cell^.ContentType of
        cctUTF8String:
          begin
            s := FWorksheet.ReadAsText(cell) + #0;
            Move(s[1], Buffer^, Length(s));
          end;
        cctNumber:
          case field.DataType of
            ftFloat:
              Move(cell^.NumberValue, Buffer^, SizeOf(cell^.NumberValue));
            ftInteger:
              begin
                i := Round(cell^.NumberValue);
                Move(i, Buffer^, SizeOf(i));
              end;
            ftString:
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
            b := cell^.BoolValue;    // Boolean field stores value as wordbool
            Move(b, Buffer^, SizeOf(b));
          end;
        cctEmpty:
          SetFieldIsNull(GetNullMaskPtr(bufferStart), field.FieldNo);
        else
          ;
      end;
    end;
    inc(Buffer, field.DataSize);
  end;
                       (*
  for field in Fields do
  begin
    P := P + FFieldOffsets[field.Index];
    case field.Datatype of
      ftString: WriteLn(field.Index, ': ', PChar(P));
      ftInteger: WriteLn(field.Index, ': ', PInteger(P)^);
      ftFloat: WriteLn(field.Index, ': ', PDouble(P)^);
      ftDateTime: WriteLn(field.Index, ': ', PDateTime(P)^);
      ftBoolean: WriteLn(field.Index, ': ', PWordBool(P)^);
      else ;
    end;
  end;

  for i := 0 to RecordSize-1 do
  begin
    Write(Format('%.2x ', [byte(Q^)]));
    inc(Q);
  end;
  WriteLn;
  *)

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
        LoadRecordToBuffer(FFilterBuffer, ARecNo);
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
            if TStringField(field).CodePage=CP_UTF8 then
            begin
              s1 := field.AsUTF8String;
              s2 := UTF8Encode(VarToUnicodeStr(AKeyValues[i]));
            end else
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

procedure TsWorksheetDataset.SetBookmarkData(Buffer: TRecordBuffer; Data: Pointer);
begin
  if Data <> nil then
    GetRecordInfoPtr(Buffer)^.Bookmark := PInteger(Data)^
  else
    GetRecordInfoPtr(Buffer)^.Bookmark := 0;
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
        if Field.DataType = ftString then
          dec(fSize);  // Do not move terminating 0 which is included in DataSize
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

procedure TsWorksheetDataset.SetFiltered(Value: Boolean);
var
  changed: Boolean;
begin
  changed := Value <> inherited Filtered;
  inherited;
  if changed then
    Refresh;
end;

procedure TsWorksheetDataset.SetFilterText(const AValue: string);
begin
  // Just do nothing; filter is not implemented
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

procedure TsWorksheetDataset.WriteBufferToWorksheet(Buffer: TRecordBuffer);
var
  row: TRowIndex;
  col: TColIndex;
  cell: PCell;
  field: TField;
  P: Pointer;
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
          if TFloatField(field).Precision >= 15 then
            FWorksheet.WriteNumber(cell, PDouble(P)^, nfGeneral)
          else
            FWorksheet.WriteNumber(cell, PDouble(P)^, nfFixed, TFloatField(field).Precision);
        ftInteger:
          FWorksheet.WriteNumber(cell, PInteger(P)^);
        ftDateTime:
          FWorksheet.WriteDateTime(cell, PDateTime(P)^, nfShortDateTime);
        ftDate:
          FWorksheet.WriteDateTime(Cell, PDateTime(P)^, nfShortDate);
        ftTime:
          FWorksheet.WriteDateTime(cell, PDateTime(P)^, nfLongTime);
        ftBoolean:
          FWorksheet.WriteBoolValue(cell, PWordBool(P)^);
        ftString:
          FWorksheet.WriteText(cell, StrPas(PChar(P)));
        else
          ;
      end;
    end;
  end;
  FModified := true;
end;


end.

