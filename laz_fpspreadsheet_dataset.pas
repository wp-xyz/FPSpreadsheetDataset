{ This file was automatically created by Lazarus. Do not edit!
  This source is only used to compile and install the package.
 }

unit laz_fpspreadsheet_dataset;

{$warn 5023 off : no warning about unused units}
interface

uses
  fpsDataset, LazarusPackageIntf;

implementation

procedure Register;
begin
  RegisterUnit('fpsDataset', @fpsDataset.Register);
end;

initialization
  RegisterPackage('laz_fpspreadsheet_dataset', @Register);
end.
