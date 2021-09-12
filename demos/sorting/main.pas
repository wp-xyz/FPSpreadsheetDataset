unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, DB, DBGrids, ComCtrls,
  Menus, ExtCtrls, DBCtrls, StdCtrls, fpsTypes, fpsDataset;

type

  { TForm1 }

  TForm1 = class(TForm)
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    ImageList1: TImageList;
    RecordInfo: TLabel;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    Panel1: TPanel;
    PopupMenu1: TPopupMenu;
    sWorksheetDataset1: TsWorksheetDataset;
    procedure DBGrid1TitleClick(Column: TColumn);
    procedure FormCreate(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure sWorksheetDataset1AfterOpen(DataSet: TDataSet);
    procedure sWorksheetDataset1AfterScroll(DataSet: TDataSet);
  private
    FSortColumn: TColumn;
    FFilterField: TField;
    FFilterText: String;
    procedure FilterRecord({%H-}ADataSet: TDataSet; var Accept: Boolean);
    procedure UpdateRecordInfo;

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

uses
  ListboxDlg;

{ TForm1 }

procedure TForm1.FilterRecord(ADataSet: TDataSet; var Accept: Boolean);
begin
  Accept := FFilterField.AsString = FFilterText;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  f: TField;
begin
  // Open the spreadsheet file as dataset.
  sWorksheetDataset1.FileName := 'PlantList.xls';
  sWorksheetDataset1.Open;

  // Avoid too many decimal places in floating point fields.
  for f in sWorksheetDataset1.Fields do
    if (f is TFloatField) then
      TFloatField(f).DisplayFormat := '0.###';
end;

procedure TForm1.MenuItem1Click(Sender: TObject);
begin
  sWorksheetDataset1.Filtered := false;
  sWorksheetDataset1.OnFilterRecord := nil;
end;

procedure TForm1.MenuItem2Click(Sender: TObject);

  procedure GetFieldValues(ADataset: TDataset; AField: TField; AList: TStrings);
  var
    bm: TBookmark;
    L: TStringList;
  begin
    bm := ADataset.GetBookmark;
    ADataset.DisableControls;
    L := TStringList.Create;
    try
      L.Sorted := true;
      L.Duplicates := dupIgnore;
      ADataset.First;
      while not ADataset.EOF do
      begin
        L.Add(AField.AsString);
        ADataset.Next;
      end;
      AList.Assign(L);
    finally
      L.Free;
      if ADataset.BookmarkValid(bm) then
      begin
        ADataset.GotoBookmark(bm);
        ADataset.FreeBookmark(bm);
      end;
      ADataset.EnableControls;
    end;
  end;

var
  P: TPoint;
  F: TListboxForm;
begin
  sWorksheetDataset1.OnFilterRecord := nil;
  sWorksheetDataset1.Filtered := false;

  if DBGrid1.SelectedColumn = nil then
  begin
    ShowMessage('No column selected.');
    exit;
  end;

  FFilterField := DBGrid1.SelectedColumn.Field;
  FFilterText := FFilterField.AsString;

  F := TListboxForm.Create(nil);
  try
    F.Caption := 'Filter';
    F.Prompt.Caption := FFilterField.FieldName + ' matches...';
    GetFieldValues(sWorksheetDataset1, FFilterField, F.Listbox.Items);
    F.Listbox.ItemIndex := F.Listbox.Items.IndexOf(FFilterText);
    if F.ShowModal = mrOK then
    begin
      FFilterText := F.Listbox.Items[F.Listbox.ItemIndex];
      sWorksheetDataset1.Filtered := false;
      sWorksheetDataset1.OnFilterRecord := @FilterRecord;
      sWorksheetDataset1.Filtered := true;
      UpdateRecordInfo;
    end;
  finally
    F.Free;
  end;
end;

procedure TForm1.sWorksheetDataset1AfterOpen(DataSet: TDataSet);
begin
  UpdateRecordInfo;
end;

procedure TForm1.sWorksheetDataset1AfterScroll(DataSet: TDataSet);
begin
  UpdateRecordInfo;
end;

procedure TForm1.UpdateRecordInfo;
begin
  RecordInfo.Caption := Format('Record %d of %d (relative to unfiltered dataset)', [
    sWorksheetDataset1.RecNo,
    sWorksheetDataset1.RecordCount
  ]);
end;

{ Sorts the grid (and worksheet) when a grid header is clicked. A sort indicator
  image is displayed at the right of the column title. Requires an ImageList
  assigned to the grid's TitleImageList having the image for ascending and
  descending sorts at index 0 and 1, respectively. }
procedure TForm1.DBGrid1TitleClick(Column: TColumn);
var
  options: TsSortOptions;
begin
  options := [];  // [] --> ascending sort

  if FSortColumn = Column then
  // Previously selected sort column was clicked another time...
  begin
    // Toggle between ascending and descending sort images
    FSortColumn.Title.ImageIndex := (FSortColumn.Title.ImageIndex + 1) mod 2;
    if FSortColumn.Title.ImageIndex = 1 then
      options := [ssoDescending];
  end
  else
  // A previously unsorted column was clicked...
  begin
    // Remove sort image from old sort column
    if FSortColumn <> nil then FSortColumn.Title.ImageIndex := -1;
    // Store clicked column as new SortColumn
    FSortColumn := Column;
    // Set sort image index to "ascending sort"
    FSortColumn.Title.ImageIndex := 0;
  end;

  // Execute the sorting operation.
  sWorksheetDataset1.SortOnField(FSortColumn.Field.FieldName, options);
end;

end.

