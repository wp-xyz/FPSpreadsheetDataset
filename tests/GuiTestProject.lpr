program GuiTestProject;

{$mode objfpc}{$H+}

uses
  Interfaces, Forms, GuiTestRunner, ReadFieldsTest;

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TGuiTestRunner, TestRunner);
  Application.Run;
end.

