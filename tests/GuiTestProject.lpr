program GuiTestProject;

{$mode objfpc}{$H+}

uses
  Interfaces, Forms, GuiTestRunner, ReadFieldsTest, sorttestunit;

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TGuiTestRunner, TestRunner);
  Application.Run;
end.

