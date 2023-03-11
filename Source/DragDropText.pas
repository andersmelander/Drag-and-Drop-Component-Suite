unit DragDropText;
// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite.
// Module:          DragDropText
// Description:     Implements Dragging and Dropping of different text formats.
// Version:         4.2
// Date:            05-APR-2008
// Target:          Win32, Delphi 5-2007
// Authors:         Anders Melander, anders@melander.dk, http://melander.dk
// Copyright        © 1997-1999 Angus Johnson & Anders Melander
//                  © 2000-2008 Anders Melander
// -----------------------------------------------------------------------------

{$define DROPSOURCE_TEXTSCRAP}

interface

uses
  DragDrop,
  DropTarget,
  DropSource,
  DragDropFormats,
  ActiveX,
  Windows,
  Classes;

{$include DragDrop.inc}

type
////////////////////////////////////////////////////////////////////////////////
//
//		TRichTextClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TRichTextClipboardFormat = class(TCustomTextClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    property Text;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TUnicodeTextClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
// Note: On Windows NT/2000 the system automatically synthesizes the
// CF_UNICODETEXT format from the CF_TEXT and CF_OEMTEXT formats.
////////////////////////////////////////////////////////////////////////////////
  TUnicodeTextClipboardFormat = class(TCustomWideTextClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    property Text;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TOEMTextClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
// Note: The system automatically synthesizes the CF_OEMTEXT format from the
// CF_TEXT format. On Windows NT/2000 the system also synthesizes from the
// CF_UNICODETEXT format.
////////////////////////////////////////////////////////////////////////////////
  TOEMTextClipboardFormat = class(TCustomTextClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
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
    property Locale: DWORD read GetValueDWORD write SetValueDWORD;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TTextDataFormat
//
////////////////////////////////////////////////////////////////////////////////
  TTextDataFormat = class(TCustomDataFormat)
  private
    FText: string;
    FUnicodeText: WideString;
    FRichText: string;
    FOEMText: string;
    FCSVText: string;
    FHTML: string;
    FLocale: DWORD;
  protected
    procedure SetText(const Value: string);
    procedure SetCSVText(const Value: string);
    procedure SetOEMText(const Value: string);
    procedure SetRichText(const Value: string);
    procedure SetHTML(const Value: string);
    procedure SetUnicodeText(const Value: WideString);
    procedure SetLocale(const Value: DWORD);
  public
    function Assign(Source: TClipboardFormat): boolean; override;
    function AssignTo(Dest: TClipboardFormat): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function NeedsData: boolean; override;
    property Text: string read FText write SetText;
    property UnicodeText: WideString read FUnicodeText write SetUnicodeText;
    property RichText: string read FRichText write SetRichText;
    property OEMText: string read FOEMText write SetOEMText;
    property CSVText: string read FCSVText write SetCSVText;
    property HTML: string read FHTML write SetHTML;
    property Locale: DWORD read FLocale write SetLocale;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDropTextTarget
//
////////////////////////////////////////////////////////////////////////////////
  TDropTextTarget = class(TCustomDropMultiTarget)
  private
    FTextFormat: TTextDataFormat;
  protected
    function GetText: string;
    function GetCSVText: string;
    function GetLocale: DWORD;
    function GetOEMText: string;
    function GetRichText: string;
    function GetUnicodeText: WideString;
    function GetHTML: string;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    property Text: string read GetText;
    property UnicodeText: WideString read GetUnicodeText;
    property RichText: string read GetRichText;
    property OEMText: string read GetOEMText;
    property CSVText: string read GetCSVText;
    property HTML: string read GetHTML;
    property Locale: DWORD read GetLocale;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDropTextSource
//
////////////////////////////////////////////////////////////////////////////////
  TDropTextSource = class(TCustomDropMultiSource)
  private
    FTextFormat: TTextDataFormat;
  protected
    function GetText: string;
    function GetCSVText: string;
    function GetLocale: DWORD;
    function GetOEMText: string;
    function GetRichText: string;
    function GetUnicodeText: WideString;
    function GetHTML: string;
    procedure SetText(const Value: string);
    procedure SetCSVText(const Value: string);
    procedure SetLocale(const Value: DWORD);
    procedure SetOEMText(const Value: string);
    procedure SetRichText(const Value: string);
    procedure SetUnicodeText(const Value: WideString);
    procedure SetHTML(const Value: string);
  public
    constructor Create(aOwner: TComponent); override;
    destructor Destroy; override;
  published
    property Text: string read GetText write SetText;
    property UnicodeText: WideString read GetUnicodeText write SetUnicodeText;
    property RichText: string read GetRichText write SetRichText;
    property OEMText: string read GetOEMText write SetOEMText;
    property CSVText: string read GetCSVText write SetCSVText;
    property HTML: string read GetHTML write SetHTML;
    property Locale: DWORD read GetLocale write SetLocale default 0;
  end;


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
  ShlObj,
  SysUtils;

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

////////////////////////////////////////////////////////////////////////////////
//
//		TUnicodeTextClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
function TUnicodeTextClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  Result := CF_UNICODETEXT;
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


////////////////////////////////////////////////////////////////////////////////
//
//		TTextDataFormat
//
////////////////////////////////////////////////////////////////////////////////

function TTextDataFormat.Assign(Source: TClipboardFormat): boolean;
begin
  Result := True;

  if (Source is TTextClipboardFormat) then
    FText := TTextClipboardFormat(Source).Text

  else if (Source is TRichTextClipboardFormat) then
    FRichText := TRichTextClipboardFormat(Source).Text

  else if (Source is TUnicodeTextClipboardFormat) then
    FUnicodeText := TUnicodeTextClipboardFormat(Source).Text

  else if (Source is TOEMTextClipboardFormat) then
    FOEMText := TOEMTextClipboardFormat(Source).Text

  else if (Source is TCSVClipboardFormat) then
    FCSVText := TCSVClipboardFormat(Source).Lines.Text

  else if (Source is TLocaleClipboardFormat) then
    FLocale := TLocaleClipboardFormat(Source).Locale

{$ifdef DROPSOURCE_TEXTSCRAP}
  else if (Source is TFileContentsClipboardFormat) then
    FText := TFileContentsClipboardFormat(Source).Data
{$endif}

  else
    Result := inherited Assign(Source);
end;

function TTextDataFormat.AssignTo(Dest: TClipboardFormat): boolean;
var
{$ifdef DROPSOURCE_TEXTSCRAP}
  FGD: TFileGroupDescriptor;
  FGDW: DragDropFormats.TFileGroupDescriptorW;
{$endif}
  s: string;
{$ifdef DROPSOURCE_TEXTSCRAP}
resourcestring
  // Name of the text scrap file.
  sTextScrap = 'Text scrap.txt';
{$endif}
begin
  Result := True;

  if (Dest is TTextClipboardFormat) then
  begin
    if (FText <> '') then
      TTextClipboardFormat(Dest).Text := FText
    else
    // Synthesize ANSI text from Unicode text.
    if (FUnicodeText <> '') then
    begin
      // TODO: Take Locale into account
      TTextClipboardFormat(Dest).Text := FUnicodeText;
    end else
    // Synthesize ANSI text from OEM text.
    if (FOEMText <> '') then
    begin
      // Convert OEM string to ANSI string...
      SetLength(s, Length(FOEMText));
      OemToCharBuff(PChar(FOEMText), PChar(s),
        Length(s));
      TTextClipboardFormat(Dest).Text := s;
    end else
      TTextClipboardFormat(Dest).Text := '';
  end else

  if (Dest is TRichTextClipboardFormat) then
  begin
    if (FRichText <> '') then
      TRichTextClipboardFormat(Dest).Text := FRichText
    else
      TRichTextClipboardFormat(Dest).Text := MakeRTF(FText);
  end else

  if (Dest is TUnicodeTextClipboardFormat) then
  begin
    if (FUnicodeText <> '') then
      TUnicodeTextClipboardFormat(Dest).Text := FUnicodeText
    else
      // TODO: Take Locale into account
      // Synthesize Unicode text from ANSI text.
      TUnicodeTextClipboardFormat(Dest).Text := FText;
  end else

  if (Dest is TOEMTextClipboardFormat) then
  begin
    if (FOEMText <> '') then
      TOEMTextClipboardFormat(Dest).Text := FOEMText
    else
    // Synthesize OEM text from ANSI text.
    if (FText <> '') then
    begin
      // Convert ANSI string to OEM string.
      SetLength(s, Length(FText));
      CharToOemBuff(PChar(FText), PChar(s),
        Length(FText));
      TOEMTextClipboardFormat(Dest).Text := s;
    end else
      TOEMTextClipboardFormat(Dest).Text := '';
  end else

  if (Dest is TCSVClipboardFormat) then
  begin
    if (FCSVText <> '') then
      TCSVClipboardFormat(Dest).Lines.Text := FCSVText
    else
      TCSVClipboardFormat(Dest).Lines.Text := FText;
  end else

  if (Dest is TLocaleClipboardFormat) then
  begin
    TLocaleClipboardFormat(Dest).Locale := FLocale
  end else

{$ifdef DROPSOURCE_TEXTSCRAP}
  // TODO : Get rid of this. It doesn't belong here.
  if (Dest is TFileContentsClipboardFormat) then
  begin
    TFileContentsClipboardFormat(Dest).Data := FText
  end else

  if (Dest is TFileGroupDescritorClipboardFormat) then
  begin
    FillChar(FGD, SizeOf(FGD), 0);
    FGD.cItems := 1;
    StrPLCopy(FGD.fgd[0].cFileName, sTextScrap, SizeOf(FGD.fgd[0].cFileName));
    TFileGroupDescritorClipboardFormat(Dest).CopyFrom(@FGD);
  end else

  if (Dest is TFileGroupDescritorWClipboardFormat) then
  begin
    FillChar(FGDW, SizeOf(FGDW), 0);
    FGDW.cItems := 1;
    StringToWideChar(sTextScrap, PWideChar(@(FGDW.fgd[0].cFileName)), MAX_PATH);
    TFileGroupDescritorWClipboardFormat(Dest).CopyFrom(@FGDW);
  end else
{$endif}
    Result := inherited AssignTo(Dest);
end;

procedure TTextDataFormat.Clear;
begin
  Changing;
  FText := '';
  FOEMText := '';
  FCSVText := '';
  FRichText := '';
  FUnicodeText := '';
  FHTML := '';
  FLocale := 0;
end;

procedure TTextDataFormat.SetText(const Value: string);
begin
  Changing;
  FText := Value;
  if (FLocale = 0) and not(csDesigning in Owner.ComponentState) then
    FLocale := GetThreadLocale;
end;

procedure TTextDataFormat.SetCSVText(const Value: string);
begin
  Changing;
  FCSVText := Value;
end;

procedure TTextDataFormat.SetOEMText(const Value: string);
begin
  Changing;
  FOEMText := Value;
end;

procedure TTextDataFormat.SetRichText(const Value: string);
begin
  Changing;
  FRichText := Value;
end;

procedure TTextDataFormat.SetUnicodeText(const Value: WideString);
begin
  Changing;
  FUnicodeText := Value;
end;

procedure TTextDataFormat.SetHTML(const Value: string);
begin
  Changing;
  FHTML := Value;
end;

procedure TTextDataFormat.SetLocale(const Value: DWORD);
begin
  Changing;
  FLocale := Value;
end;

function TTextDataFormat.HasData: boolean;
begin
  Result := (FText <> '') or (FOEMText <> '') or (FCSVText <> '') or
    (FRichText <> '') or (FUnicodeText <> '') or (FHTML <> '');
end;

function TTextDataFormat.NeedsData: boolean;
begin
  Result := (FText = '') or (FOEMText = '') or (FCSVText = '') or
    (FRichText = '') or (FUnicodeText = '') or (FHTML = '');
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

function TDropTextTarget.GetCSVText: string;
begin
  Result := FTextFormat.CSVText;
end;

function TDropTextTarget.GetHTML: string;
begin
  Result := FTextFormat.HTML;
end;

function TDropTextTarget.GetLocale: DWORD;
begin
  Result := FTextFormat.Locale;
end;

function TDropTextTarget.GetOEMText: string;
begin
  Result := FTextFormat.OEMText
end;

function TDropTextTarget.GetRichText: string;
begin
  Result := FTextFormat.RichText;
end;

function TDropTextTarget.GetText: string;
begin
  Result := FTextFormat.Text
end;

function TDropTextTarget.GetUnicodeText: WideString;
begin
  Result := FTextFormat.UnicodeText
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

function TDropTextSource.GetCSVText: string;
begin
  Result := FTextFormat.CSVText;
end;

function TDropTextSource.GetHTML: string;
begin
  Result := FTextFormat.HTML;
end;

function TDropTextSource.GetLocale: DWORD;
begin
  Result := FTextFormat.Locale;
end;

function TDropTextSource.GetOEMText: string;
begin
  Result := FTextFormat.OEMText;
end;

function TDropTextSource.GetRichText: string;
begin
  Result := FTextFormat.RichText;
end;

function TDropTextSource.GetText: string;
begin
  Result := FTextFormat.Text;
end;

function TDropTextSource.GetUnicodeText: WideString;
begin
  Result := FTextFormat.UnicodeText;
end;

procedure TDropTextSource.SetCSVText(const Value: string);
begin
  FTextFormat.CSVText := Value;
end;

procedure TDropTextSource.SetHTML(const Value: string);
begin
  FTextFormat.HTML := Value;
end;

procedure TDropTextSource.SetLocale(const Value: DWORD);
begin
  FTextFormat.Locale := Value;
end;

procedure TDropTextSource.SetOEMText(const Value: string);
begin
  FTextFormat.OEMText := Value;
end;

procedure TDropTextSource.SetRichText(const Value: string);
begin
  FTextFormat.RichText := Value;
end;

procedure TDropTextSource.SetText(const Value: string);
begin
  FTextFormat.Text := Value;
end;

procedure TDropTextSource.SetUnicodeText(const Value: WideString);
begin
  FTextFormat.UnicodeText := Value;
end;

////////////////////////////////////////////////////////////////////////////////
//
//		Initialization/Finalization
//
////////////////////////////////////////////////////////////////////////////////

initialization
  // Data format registration
  TTextDataFormat.RegisterDataFormat;
  // Clipboard format registration
  TTextDataFormat.RegisterCompatibleFormat(TTextClipboardFormat, 0, csSourceTarget, [ddRead]);
  TTextDataFormat.RegisterCompatibleFormat(TUnicodeTextClipboardFormat, 0, csSourceTarget, [ddRead]);
  TTextDataFormat.RegisterCompatibleFormat(TRichTextClipboardFormat, 0, csSourceTarget, [ddRead]);
  TTextDataFormat.RegisterCompatibleFormat(TOEMTextClipboardFormat, 0, csSourceTarget, [ddRead]);
  TTextDataFormat.RegisterCompatibleFormat(TCSVClipboardFormat, 0, csSourceTarget, [ddRead]);
  TTextDataFormat.RegisterCompatibleFormat(TLocaleClipboardFormat, 0, csSourceTarget, [ddRead]);
{$ifdef DROPSOURCE_TEXTSCRAP}
  TTextDataFormat.RegisterCompatibleFormat(TFileContentsClipboardFormat, 1, csSourceTarget, [ddRead]);
  TTextDataFormat.RegisterCompatibleFormat(TFileGroupDescritorClipboardFormat, 1, [csSource], [ddRead]);
  TTextDataFormat.RegisterCompatibleFormat(TFileGroupDescritorWClipboardFormat, 1, [csSource], [ddRead]);
{$endif}

finalization
  // Data format unregistration
  TTextDataFormat.UnregisterDataFormat;
  // Clipboard format unregistration
  TUnicodeTextClipboardFormat.UnregisterClipboardFormat;
  TRichTextClipboardFormat.UnregisterClipboardFormat;
  TOEMTextClipboardFormat.UnregisterClipboardFormat;
  TCSVClipboardFormat.UnregisterClipboardFormat;
  TLocaleClipboardFormat.UnregisterClipboardFormat;
end.
