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
    procedure DropFileTarget1Enter(Sender: TObject; ShiftState: TShiftState;
      APoint: TPoint; var Effect: Integer);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

uses
  ActiveX,
  DragDropFormats;

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
  ** The OnDrop event handler is called when the user drag and drop files
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
    if (DropFileTarget1.MappedNames.Count > i) then
      Item.SubItems.Add(DropFileTarget1.MappedNames[i]);
  end;

  // For this demo we reject the drop if a move operation was performed. This is
  // done so the drop source doesn't think we actually did something usefull
  // with the data.
  // This is important when moving data or when dropping from the recycle bin;
  // If we do not reject a move, the source will assume that it is safe to
  // delete the source data. See also "Optimized move".
  if (Effect = DROPEFFECT_MOVE) then
    Effect := DROPEFFECT_NONE;
end;

const
  // Just some random GUID (press Ctrl+Shift+G in the IDE to generate a GUID)
  MyClassID: TGUID = '{228FF208-AD46-46DC-B02D-7906787BF8F7}';

procedure TForm1.DropFileTarget1Enter(Sender: TObject; ShiftState: TShiftState;
  APoint: TPoint; var Effect: Integer);
var
  Medium: TStgMedium;
begin
  // The following code is an example of how to write the TargetCLSID value back
  // to the drop source in order to identify the target application.
  // Unless you have an explicit need for this feature you don't need to do this
  // in your own application.
  with TTargetCLSIDClipboardFormat.Create do
    try
      CLSID := MyClassID;
      SetData(TCustomDropTarget(Sender).DataObject, FormatEtc, Medium);
    finally
      Free;
    end;
end;

end.
