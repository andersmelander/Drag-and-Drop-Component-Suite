unit Main;

interface

uses
  DragDrop,
  DropSource,
  DragDropFile,
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, ComCtrls;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Panel2: TPanel;
    ButtonClose: TButton;
    DropFileSource1: TDropFileSource;
    ListView1: TListView;
    procedure ButtonCloseClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ListView1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
  public
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

procedure TForm1.ButtonCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  Path: string;
  SearchRec: TSearchRec;
  Res: integer;
  NewItem: TListItem;
begin
  (*
  ** Fill listview with list of files from current directory...
  *)
  Path := ExtractFilePath(Application.ExeName);
  Res := FindFirst(path+'*.*', 0, SearchRec);
  try
    while (Res = 0) do
    begin
      if (SearchRec.Name <> '.') and (SearchRec.Name <> '..') then
      begin
        NewItem := Listview1.Items.Add;
        NewItem.Caption := Path+SearchRec.Name;
      end;
      Res := FindNext(SearchRec);
    end;
  finally
    FindClose(SearchRec);
  end;
end;

procedure TForm1.ListView1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  i: integer;
begin
  // If no files selected then we can't drag.
  if (Listview1.SelCount = 0) then
    Exit;

  // Wait for user to move mouse before we start the drag/drop.
  if (DragDetectPlus(TWinControl(Sender).Handle, Point(X,Y))) then
  begin
    // Delete anything from a previous drag.
    DropFileSource1.Files.Clear;

    // Fill DropSource1.Files with selected files from ListView1.
    for i := 0 to Listview1.Items.Count-1 do
      if (Listview1.items.Item[i].Selected) then
        DropFileSource1.Files.Add(Listview1.items.Item[i].Caption);

    // Start the drag operation.
    DropFileSource1.Execute;
  end;
end;

end.
