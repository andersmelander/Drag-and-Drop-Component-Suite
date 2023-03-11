unit DragDropText;

// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite.
// Module:          DragDropText
// Description:     Implements Dragging and Dropping of different text formats.
// Version:         4.0
// Date:            18-MAY-2001
// Target:          Win32, Delphi 5-6
// Authors:         Anders Melander, anders@melander.dk, http://www.melander.dk
// Copyright        © 1997-2001 Angus Johnson & Anders Melander
// -----------------------------------------------------------------------------
interface

uses
  DragDrop,
  DropTarget,
  DropSource,
  DragDropFormats,
  ActiveX,
  Windows,
  Classes;

type
////////////////////////////////////////////////////////////////////////////////
//
//		TRichTextClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TRichTextClipboardFormat = class(TCustomTextClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    function HasData: boolean; override;
    function Assign(Source: TCustomDataFormat): boolean; override;
    function AssignTo(Dest: TCustomDataFormat): boolean; override;
    property Text;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TUnicodeTextClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TUnicodeTextClipboardFormat = class(TCustomWideTextClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    function Assign(Source: TCustomDataFormat): boolean; override;
    function AssignTo(Dest: TCustomDataFormat): boolean; override;
    property Text;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TOEMTextClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TOEMTextClipboardFormat = class(TCustomTextClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    function Assign(Source: TCustomDataFormat): boolean; override;
    function AssignTo(Dest: TCustomDataFormat): boolean; override;
    property Text;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TCSVClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TCSVClipboardFormat = class(TCustomStringListClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    function Assign(Source: TCustomDataFormat): boolean; override;
    function AssignTo(Dest: TCustomDataFormat): boolean; override;
    property Lines;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TLocaleClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TLocaleClipboardFormat = class(TCustomDWORDClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    function HasData: boolean; override;
    function Assign(Source: TCustomDataFormat): boolean; override;
    function AssignTo(Dest: TCustomDataFormat): boolean; override;
    property Locale: DWORD read GetValueDWORD;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDropTextTarget
//
////////////////////////////////////////////////////////////////////////////////
  TDropTextTarget = class(TCustomDropMultiTarget)
  private
    FTextFormat		: TTextDataFormat;
  protected
    function GetText: string;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    property Text: string read GetText;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDropTextSource
//
////////////////////////////////////////////////////////////////////////////////
  TDropTextSource = class(TCustomDropMultiSource)
  private
    FTextFormat		: TTextDataFormat;
  protected
    function GetText: string;
    procedure SetText(const Value: string);
  public
    constructor Create(aOwner: TComponent); override;
    destructor Destroy; override;
  published
    property Text: string read GetText write SetText;
  end;

  
////////////////////////////////////////////////////////////////////////////////
//
//		Component registration
//
////////////////////////////////////////////////////////////////////////////////
procedure Register;


////////////////////////////////////////////////////////////////////////////////
//
//		Misc.
//
////////////////////////////////////////////////////////////////////////////////
function IsRTF(const s: string): boolean;
function MakeRTF(const s: string): string;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
//
//			IMPLEMENTATION
//
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
implementation

uses
  SysUtils;

////////////////////////////////////////////////////////////////////////////////
//
//		Component registration
//
////////////////////////////////////////////////////////////////////////////////
procedure Register;
begin
  RegisterComponents(DragDropComponentPalettePage, [TDropTextTarget,
    TDropTextSource]);
end;

////////////////////////////////////////////////////////////////////////////////
//
//		Utilities
//
////////////////////////////////////////////////////////////////////////////////
function IsRTF(const s: string): boolean;
begin
  // This probably isn't a valid test, but it will have to do until I have
  // time to research the RTF specifications.
  { TODO -oanme -cImprovement : Need a solid test for RTF format. }
  Result := (AnsiStrLIComp(PChar(s), '{\rtf', 5) = 0);
end;

{ TODO -oanme -cImprovement : Needs RTF to text conversion. Maybe ITextDocument can be used. }
function MakeRTF(const s: string): string;
begin
  { TODO -oanme -cImprovement : Needs to escape \ in text to RTF conversion. }
  { TODO -oanme -cImprovement : Needs better text to RTF conversion. }
  if (not IsRTF(s)) then
    Result := '{\rtf1\ansi ' + s + '}'
  else
    Result := s;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TRichTextClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_RTF: TClipFormat = 0;

function TRichTextClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  // Note: The string 'Rich Text Format', is also defined in the RichEdit
  // unit as CF_RTF
  if (CF_RTF = 0) then
    CF_RTF := RegisterClipboardFormat('Rich Text Format'); // *** DO NOT LOCALIZE ***
  Result := CF_RTF;
end;

function TRichTextClipboardFormat.HasData: boolean;
begin
  Result := inherited HasData and IsRTF(Text);
end;

function TRichTextClipboardFormat.Assign(Source: TCustomDataFormat): boolean;
begin
  if (Source is TTextDataFormat) then
  begin
    Text := MakeRTF(TTextDataFormat(Source).Text);
    Result := True;
  end else
    Result := inherited Assign(Source);
end;

function TRichTextClipboardFormat.AssignTo(Dest: TCustomDataFormat): boolean;
begin
  if (Dest is TTextDataFormat) then
  begin
    TTextDataFormat(Dest).Text := Text;
    Result := True;
  end else
    Result := inherited AssignTo(Dest);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TUnicodeTextClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
function TUnicodeTextClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  Result := CF_UNICODETEXT;
end;

function TUnicodeTextClipboardFormat.Assign(Source: TCustomDataFormat): boolean;
begin
  if (Source is TTextDataFormat) then
  begin
    Text := TTextDataFormat(Source).Text;
    Result := True;
  end else
    Result := inherited Assign(Source);
end;

function TUnicodeTextClipboardFormat.AssignTo(Dest: TCustomDataFormat): boolean;
begin
  if (Dest is TTextDataFormat) then
  begin
    TTextDataFormat(Dest).Text := Text;
    Result := True;
  end else
    Result := inherited AssignTo(Dest);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TOEMTextClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
function TOEMTextClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  Result := CF_OEMTEXT;
end;

function TOEMTextClipboardFormat.Assign(Source: TCustomDataFormat): boolean;
var
  OEMText		: string;
begin
  if (Source is TTextDataFormat) then
  begin
    // First convert ANSI string to OEM string...
    SetLength(OEMText, Length(TTextDataFormat(Source).Text));
    CharToOemBuff(PChar(TTextDataFormat(Source).Text), PChar(OEMText),
      Length(TTextDataFormat(Source).Text));
    // ...then assign OEM string
    Text := OEMText;
    Result := True;
  end else
    Result := inherited Assign(Source);
end;

function TOEMTextClipboardFormat.AssignTo(Dest: TCustomDataFormat): boolean;
var
  AnsiText		: string;
begin
  if (Dest is TTextDataFormat) then
  begin
    // First convert OEM string to ANSI string...
    SetLength(AnsiText, Length(Text));
    OemToCharBuff(PChar(Text), PChar(AnsiText), Length(Text));
    // ...then assign ANSI string
    TTextDataFormat(Dest).Text := AnsiText;
    Result := True;
  end else
    Result := inherited AssignTo(Dest);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TCSVClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_CSV: TClipFormat = 0;

function TCSVClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_CSV = 0) then
    CF_CSV := RegisterClipboardFormat('CSV'); // *** DO NOT LOCALIZE ***
  Result := CF_CSV;
end;

function TCSVClipboardFormat.Assign(Source: TCustomDataFormat): boolean;
begin
  if (Source is TTextDataFormat) then
  begin
    Lines.Text := TTextDataFormat(Source).Text;
    Result := True;
  end else
    Result := inherited AssignTo(Source);
end;

function TCSVClipboardFormat.AssignTo(Dest: TCustomDataFormat): boolean;
begin
  if (Dest is TTextDataFormat) then
  begin
    TTextDataFormat(Dest).Text := Lines.Text;
    Result := True;
  end else
    Result := inherited AssignTo(Dest);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TLocaleClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
function TLocaleClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  Result := CF_LOCALE;
end;

function TLocaleClipboardFormat.HasData: boolean;
begin
  Result := (Locale <> 0);
end;

function TLocaleClipboardFormat.Assign(Source: TCustomDataFormat): boolean;
begin
  // So far we have no one to play with...
  Result := inherited Assign(Source);
end;

function TLocaleClipboardFormat.AssignTo(Dest: TCustomDataFormat): boolean;
begin
  // So far we have no one to play with...
  Result := inherited AssignTo(Dest);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TDropTextTarget
//
////////////////////////////////////////////////////////////////////////////////
constructor TDropTextTarget.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FTextFormat := TTextDataFormat.Create(Self);
end;

destructor TDropTextTarget.Destroy;
begin
  FTextFormat.Free;
  inherited Destroy;
end;

function TDropTextTarget.GetText: string;
begin
  Result := FTextFormat.Text;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TDropTextSource
//
////////////////////////////////////////////////////////////////////////////////
constructor TDropTextSource.Create(aOwner: TComponent);
begin
  inherited Create(aOwner);
  FTextFormat := TTextDataFormat.Create(Self);
end;

destructor TDropTextSource.Destroy;
begin
  FTextFormat.Free;
  inherited Destroy;
end;

function TDropTextSource.GetText: string;
begin
  Result := FTextFormat.Text;
end;

procedure TDropTextSource.SetText(const Value: string);
begin
  FTextFormat.Text := Value;
end;

////////////////////////////////////////////////////////////////////////////////
//
//		Initialization/Finalization
//
////////////////////////////////////////////////////////////////////////////////

initialization
  // Clipboard format registration
  TTextDataFormat.RegisterCompatibleFormat(TUnicodeTextClipboardFormat, 1, csSourceTarget, [ddRead]);
  TTextDataFormat.RegisterCompatibleFormat(TRichTextClipboardFormat, 2, csSourceTarget, [ddRead]);
  TTextDataFormat.RegisterCompatibleFormat(TOEMTextClipboardFormat, 2, csSourceTarget, [ddRead]);
  TTextDataFormat.RegisterCompatibleFormat(TCSVClipboardFormat, 3, csSourceTarget, [ddRead]);

finalization
  // Clipboard format unregistration
  TUnicodeTextClipboardFormat.UnregisterClipboardFormat;
  TRichTextClipboardFormat.UnregisterClipboardFormat;
  TOEMTextClipboardFormat.UnregisterClipboardFormat;
  TCSVClipboardFormat.UnregisterClipboardFormat;
end.
