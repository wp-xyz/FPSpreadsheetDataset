program GuiTestProject;

{$mode objfpc}{$H+}

uses
  Interfaces, Forms, GuiTestRunner, 
  ReadFieldsTestUnit, SortTestUnit, FilterTestUnit, PostTestUnit, 
  EmptyColumnsTestUnit;

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TGuiTestRunner, TestRunner);
  Application.Run;
end.

