unit FormPropertySheet;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TPropertySheetForm = class(TForm)
    MemoFiles: TMemo;
  public
    constructor Create(AOwner : TComponent); override;

    procedure SetFiles(const FileList : TStrings);
  end;


////////////////////////////////////////////////////////////////////////////////

implementation

{$R *.dfm}

////////////////////////////////////////////////////////////////////////////////

constructor TPropertySheetForm.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  MemoFiles.Clear;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TPropertySheetForm.SetFiles(const FileList: TStrings);
begin
  if (FileList <> nil) then
    MemoFiles.Lines.AddStrings(FileList);
end;

////////////////////////////////////////////////////////////////////////////////

end.
