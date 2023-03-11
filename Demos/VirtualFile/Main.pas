unit Main;

interface

uses
  Windows, Classes, Controls, Forms, ExtCtrls, StdCtrls,
  DragDrop, DropSource, DragDropFile, DropTarget, Graphics, ImgList, Menus;

type
  (*
  ** This is a custom data format.
  ** The data format supports TFileContentsClipboardFormat and
  ** TFileGroupDescritorClipboardFormat.
  *)
  TVirtualFileDataFormat = class(TCustomDataFormat)
  private
    FContents: string;
    FFileName: string;
  public
    constructor Create(AOwner: TDragDropComponent); override;
    function Assign(Source: TClipboardFormat): boolean; override;
    function AssignTo(Dest: TClipboardFormat): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function NeedsData: boolean; override;
    property FileName: string read FFileName write FFileName;
    property Contents: string read FContents write FContents;
  end;

  TFormMain = class(TForm)
    EditFilename: TEdit;
    Label1: TLabel;
    MemoContents: TMemo;
    Label2: TLabel;
    DropDummy1: TDropDummy;
    Panel2: TPanel;
    PanelDragDrop: TPanel;
    Image1: TImage;
    ImageList1: TImageList;
    PopupMenu1: TPopupMenu;
    MenuCopy: TMenuItem;
    MenuPaste: TMenuItem;
    DropEmptySource1: TDropEmptySource;
    DropEmptyTarget1: TDropEmptyTarget;
    Label3: TLabel;
    procedure OnMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormCreate(Sender: TObject);
    procedure DropFileTarget1Drop(Sender: TObject; ShiftState: TShiftState;
      Point: TPoint; var Effect: Integer);
    procedure PopupMenu1Popup(Sender: TObject);
    procedure MenuCopyClick(Sender: TObject);
    procedure MenuPasteClick(Sender: TObject);
  private
    { Private declarations }
    FSourceDataFormat: TVirtualFileDataFormat;
    FTargetDataFormat: TVirtualFileDataFormat;
  public
    { Public declarations }
  end;

var
  FormMain: TFormMain;

implementation

{$R *.DFM}

uses
  DragDropFormats,
  ShlObj,
  ComObj,
  ActiveX,
  SysUtils;

procedure TFormMain.FormCreate(Sender: TObject);
begin
  // Add our own custom data format to the drag/drop components.
  FSourceDataFormat := TVirtualFileDataFormat.Create(DropEmptySource1);
  FTargetDataFormat := TVirtualFileDataFormat.Create(DropEmptyTarget1);
end;

procedure TFormMain.OnMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if DragDetectPlus(Handle, Point(X,Y)) then
  begin
    // Transfer the file name and contents to the data format...
    FSourceDataFormat.FileName := EditFileName.Text;
    FSourceDataFormat.Contents := MemoContents.Lines.Text;

    // ...and let it rip!
    DropEmptySource1.Execute;
  end;
end;

procedure TFormMain.DropFileTarget1Drop(Sender: TObject;
  ShiftState: TShiftState; Point: TPoint; var Effect: Integer);
begin
  // Transfer the file name and contents from the data format.
  EditFileName.Text := FTargetDataFormat.FileName;

  // Limit the amount of data to 32Kb. If someone drops a huge amount on data on
  // us (e.g. the AsyncTransferSource demo which transfers 10Mb of data) we need
  // to limit how much data we try to stuff into the poor memo field. Otherwise
  // we could wait for hours before transfer was finished.
  MemoContents.Lines.Text := Copy(FTargetDataFormat.Contents, 1, 1024*32);
end;

procedure TFormMain.PopupMenu1Popup(Sender: TObject);
var
  DataObject: IDataObject;
begin
  MenuCopy.Enabled := (EditFilename.Text <> '');

  (*
  ** The following shows two diffent methods of determining if the clipboard
  ** contains any data that we can paste. The two methods are in fact identical.
  ** Only the first method is used.
  *)

  (*
  ** Method 1: Get the clipboard data object and test against it.
  *)
  // Open the clipboard as an IDataObject
  OleCheck(OleGetClipboard(DataObject));
  try
    // Enable paste menu if the clipboard contains data in any of
    // the supported formats.
    MenuPaste.Enabled := DropEmptyTarget1.HasValidFormats(DataObject);
  finally
    DataObject := nil;
  end;

  (*
  ** Method 2: Let the component do it for us.
  *)
  // MenuPaste.Enabled := DropEmptyTarget1.CanPasteFromClipboard;

end;

procedure TFormMain.MenuCopyClick(Sender: TObject);
begin
  // Transfer the file name and contents to the data format...
  FSourceDataFormat.FileName := EditFileName.Text;
  FSourceDataFormat.Contents := MemoContents.Lines.Text;

  // ...and copy to clipboard.
  DropEmptySource1.CopyToClipboard;
end;

procedure TFormMain.MenuPasteClick(Sender: TObject);
begin
  DropEmptyTarget1.PasteFromClipboard;
end;


{ TVirtualFileDataFormat }

constructor TVirtualFileDataFormat.Create(AOwner: TDragDropComponent);
begin
  inherited Create(AOwner);

  // Add the "file group descriptor" and "file contents" clipboard formats to
  // the data format's list of compatible formats.
  // Note: This is normally done via TCustomDataFormat.RegisterCompatibleFormat,
  // but since this data format is only used here, it is just as easy for us to
  // add the formats manually.
  CompatibleFormats.Add(TFileContentsClipboardFormat.Create);
  CompatibleFormats.Add(TFileGroupDescritorClipboardFormat.Create);
end;

function TVirtualFileDataFormat.Assign(Source: TClipboardFormat): boolean;
begin
  Result := True;

  (*
  ** TFileContentsClipboardFormat
  *)
  if (Source is TFileContentsClipboardFormat) then
  begin
    FContents := TFileContentsClipboardFormat(Source).Data
  end else
  (*
  ** TFileGroupDescritorClipboardFormat
  *)
  if (Source is TFileGroupDescritorClipboardFormat) then
  begin
    if (TFileGroupDescritorClipboardFormat(Source).FileGroupDescriptor^.cItems > 0) then
      FFileName := TFileGroupDescritorClipboardFormat(Source).FileGroupDescriptor^.fgd[0].cFileName;
  end else
  (*
  ** None of the above...
  *)
    Result := inherited Assign(Source);
end;

function TVirtualFileDataFormat.AssignTo(Dest: TClipboardFormat): boolean;
var
  FGD: TFileGroupDescriptor;
begin
  Result := True;

  (*
  ** TFileContentsClipboardFormat
  *)
  if (Dest is TFileContentsClipboardFormat) then
  begin
    TFileContentsClipboardFormat(Dest).Data := FContents;
  end else
  (*
  ** TFileGroupDescritorClipboardFormat
  *)
  if (Dest is TFileGroupDescritorClipboardFormat) then
  begin
    FillChar(FGD, SizeOf(FGD), 0);
    FGD.cItems := 1;
    StrLCopy(@FGD.fgd[0].cFileName[0], PChar(FFileName),
      SizeOf(FGD.fgd[0].cFileName));
    TFileGroupDescritorClipboardFormat(Dest).CopyFrom(@FGD);
  end else
  (*
  ** None of the above...
  *)
    Result := inherited AssignTo(Dest);
end;

procedure TVirtualFileDataFormat.Clear;
begin
  FFileName := '';
  FContents := ''
end;

function TVirtualFileDataFormat.HasData: boolean;
begin
  Result := (FFileName <> '');
end;

function TVirtualFileDataFormat.NeedsData: boolean;
begin
  Result := (FFileName = '') or (FContents = '');
end;

end.

