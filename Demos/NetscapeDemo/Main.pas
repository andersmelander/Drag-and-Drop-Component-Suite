unit Main;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, DragDrop, DropTarget, DragDropInternet, ComCtrls;

type
  TFormMain = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Edit1: TEdit;
    DropURLTarget1: TDropURLTarget;
    Panel2: TPanel;
    Panel3: TPanel;
    Button1: TButton;
    PageControl1: TPageControl;
    TabSheetContents: TTabSheet;
    TabSheetRaw: TTabSheet;
    MemoRaw: TMemo;
    MemoContent: TMemo;
    TabSheetHeader: TTabSheet;
    ListViewHeader: TListView;
    Label2: TLabel;
    procedure DropURLTarget1Drop(Sender: TObject; ShiftState: TShiftState;
      APoint: TPoint; var Effect: Integer);
    procedure Button1Click(Sender: TObject);
    procedure DropURLTarget1Enter(Sender: TObject; ShiftState: TShiftState;
      APoint: TPoint; var Effect: Integer);
  private
    procedure ParseRFC822(const Msg: string);
  public
  end;

var
  FormMain: TFormMain;

implementation

{$R *.DFM}

{$include ..\..\Components\DragDrop.inc}

uses
{$ifdef VER14_PLUS}
  Variants,
{$endif}
  DragDropFormats,
  ActiveX,
  Registry,
  ShellAPI,
  ComObj;

procedure TFormMain.DropURLTarget1Drop(Sender: TObject;
  ShiftState: TShiftState; APoint: TPoint; var Effect: Integer);
var
  Netscape: Variant;
  Buffer: WideString;
  Size: integer;
  s: string;
begin
  Edit1.Text := DropURLTarget1.URL;

  Netscape := CreateOleObject('Netscape.Network.1');
  try

    if (Netscape.Open(DropURLTarget1.URL, 0, '', 0, '') = 0) then
      raise Exception.CreateFmt('Failed to open URL: %s', [DropURLTarget1.URL]);

    Size := 1;
    SetLength(Buffer, 255);
    s := '';

    // Note: Netscape Messenger must be running in order to read the mail.
    while (Size > 0) do
    begin
      Size := Netscape.Read(Buffer, Length(Buffer));
      if (Size > 0) then
        s := s+Copy(Buffer, 1, Size);
    end;

    ParseRFC822(s);

    if (Size = -1) and (Netscape.GetStatus <> 0) then
      raise Exception.CreateFmt('Failed to read from URL. Status: %d',
        [integer(Netscape.GetStatus)]);
  finally
    Netscape := Null;
  end;
end;

procedure TFormMain.Button1Click(Sender: TObject);
var
  s: string;
begin
  s := '';

{$ifdef VER120}
  with TRegistry.Create do
{$else}
  with TRegistry.Create(KEY_READ) do
{$endif}
    try
      RootKey := HKEY_LOCAL_MACHINE;

      // Note: The following registry keys should work for all versions of
      // Navigator 4.x. Previous versions has not been tested/verified.
      if (OpenKeyReadOnly('\SOFTWARE\Netscape\Netscape Navigator')) then
      begin
        s := ReadString('CurrentVersion');
        if (OpenKeyReadOnly(format('\SOFTWARE\Netscape\Netscape Navigator\%s\Main', [s]))) then
        begin
          s := ReadString('Install Directory');
          if (s <> '') then
            s := format('%s\Program\netscape.exe', [s]);
        end;
      end else
      // Note: The following registry keys are valid for Navigator 6, but the
      // ShellExecute haven't been tested with version 6.
      if (OpenKeyReadOnly('\SOFTWARE\Netscape\Netscape 6')) then
      begin
        s := ReadString('CurrentVersion');
        if (OpenKeyReadOnly(format('\SOFTWARE\Netscape\Netscape 6\%s\Main', [s]))) then
          s := ReadString('PathToExe')
        else
          s := '';
      end;

    finally
      Free;
    end;
  if (s <> '') then
    ShellExecute(Handle, 'Open', PChar(s), PChar(Edit1.Text), '', SW_SHOWNORMAL);
end;

type
  TNetscapeMessageClipboardFormat = class(TCustomTextClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    property Data;
  end;

var
  CF_NETSCAPEMESSAGE: TClipFormat = 0;

function TNetscapeMessageClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_NETSCAPEMESSAGE = 0) then
    CF_NETSCAPEMESSAGE := RegisterClipboardFormat('Netscape Message');
  Result := CF_NETSCAPEMESSAGE;
end;

procedure TFormMain.DropURLTarget1Enter(Sender: TObject;
  ShiftState: TShiftState; APoint: TPoint; var Effect: Integer);
begin
  // Reject drop unless the "Netscape Message" format is present in the data
  // object. Even though this format doesn't contain any usefull information,
  // its presence does indicate that the drop is a Netscape message and not a
  // regular URL.
  with TNetscapeMessageClipboardFormat.Create do
    try
      if not(HasValidFormats(DropURLTarget1.DataObject)) then
        Effect := DROPEFFECT_NONE;
    finally
      Free;
    end;
  // Another way to do the same could be to set GetDataOnEnter = True and
  // then parse the URL here. If the URL starts with "mailbox:" the drop
  // contains a mail message.
end;

procedure TFormMain.ParseRFC822(const Msg: string);
var
  Lines: TStringList;
  i: integer;
  n: integer;
  s: string;
  Value: string;

  function Trim(const Str: string): string;
  var
    p: PChar;
  begin
    p := PChar(Str);
    while (p^ in [' ', #9]) and (p^ <> #0) do
      inc(p);
    Result := p;
  end;

begin
  (*
  ** Parse a text message in RFC 822 format.
  **
  ** Note: This is a very simple implementation. I have not read the RFC822
  ** specifications and do not know if what I do here is correct. If you need
  ** a RFC822 decoder you should probably look elsewhere for a reference
  ** implementation.
  *)
  ListViewHeader.Items.Clear;
  MemoContent.Lines.Clear;
  MemoRaw.Lines.Text := Msg;
  Lines := TStringList.Create;
  try
    Lines.Text := Msg;
    i := 0;
    while (i < Lines.Count) do
    begin
      s := Lines[i];
      n  := Pos(':', s);
      if (n = 0) then
      begin
        if (s = '') then
          Inc(i);
        break;
      end;

      with ListViewHeader.Items.Add do
      begin
        Caption := Copy(s, 1, n-1);
        Value := Copy(s, n+1, Length(s));
        Inc(i);
        while (i < Lines.Count) do
        begin
          s := Lines[i];
          if (s <> '') and (s[1] in [' ',#9]) then
          begin
            Value := Value+' '+Trim(s);
            Inc(i);
          end else
            break;
        end;
        SubItems.Add(Value);
      end;
    end;

    while (i < Lines.Count) do
    begin
      s := Lines[i];
      Inc(i);
      while (Copy(s, Length(s)-2, 3) = '=20') and (i < Lines.Count) do
      begin
        s := Copy(s, 1, Length(s)-3)+Lines[i];
        Inc(i);
      end;
      MemoContent.Lines.Add(s);
    end;
  finally
    Lines.Free;
  end;
end;

end.
