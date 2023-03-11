unit Main;

interface

uses
  DragDrop,
  DropTarget,
  DragDropFile,
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls, ExtCtrls;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Panel2: TPanel;
    ButtonClose: TButton;
    DropFileTarget1: TDropFileTarget;
    ListView1: TListView;
    procedure ButtonCloseClick(Sender: TObject);
    procedure DropFileTarget1Drop(Sender: TObject; ShiftState: TShiftState;
      Point: TPoint; var Effect: Integer);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

procedure TForm1.ButtonCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TForm1.DropFileTarget1Drop(Sender: TObject;
  ShiftState: TShiftState; Point: TPoint; var Effect: Integer);
var
  i: integer;
  Item: TListItem;
begin
  (*
  ** The OnDrop event handler is called when the user drags files and drops them
  ** onto your application.
  *)

  Listview1.Items.Clear;

  // Display mapped names if present.
  // Mapped names are usually only present when dragging from the recycle bin
  // (try it).
  if (DropFileTarget1.MappedNames.Count > 0) then
    Listview1.Columns[1].Width := 100
  else
    Listview1.Columns[1].Width := 0;

  // Copy the file names from the DropTarget component into the list view.
  for i := 0 to DropFileTarget1.Files.Count-1 do
  begin
    Item := Listview1.Items.Add;
    Item.Caption := DropFileTarget1.Files[i];

    // Display mapped names if present.
    if (DropFileTarget1.MappedNames.Count > 0) then
      Item.SubItems.Add(DropFileTarget1.MappedNames[i]);
  end;
end;

end.
