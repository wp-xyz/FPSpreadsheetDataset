unit ListboxDlg;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ButtonPanel;

type

  { TListboxForm }

  TListboxForm = class(TForm)
    ButtonPanel1: TButtonPanel;
    Prompt: TLabel;
    ListBox: TListBox;
  private

  public

  end;

var
  ListboxForm: TListboxForm;

implementation

{$R *.lfm}

end.

