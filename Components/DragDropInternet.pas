unit DragDropInternet;

// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite.
// Module:          DragDropInternet
// Description:     Implements Dragging and Dropping of internet related data.
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
  Windows,
  Classes,
  ActiveX;

type

////////////////////////////////////////////////////////////////////////////////
//
//		TURLClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
// Implements support for the 'UniformResourceLocator' format.
////////////////////////////////////////////////////////////////////////////////

  TURLClipboardFormat = class(TCustomTextClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    property URL: string read GetString write SetString;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TNetscapeBookmarkClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
// Implements support for the 'Netscape Bookmark' format.
////////////////////////////////////////////////////////////////////////////////
  TNetscapeBookmarkClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FURL		: string;
    FTitle		: string;
  protected
    function ReadData(Value: pointer; Size: integer): boolean; override;
    function WriteData(Value: pointer; Size: integer): boolean; override;
    function GetSize: integer; override;
  public
    function GetClipboardFormat: TClipFormat; override;
    procedure Clear; override;
    property URL: string read FURL write FURL;
    property Title: string read FTitle write FTitle;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TNetscapeImageClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
// Implements support for the 'Netscape Image Format' format.
////////////////////////////////////////////////////////////////////////////////
  TNetscapeImageClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FURL		: string;
    FTitle		: string;
    FImage		: string;
    FLowRes		: string;
    FExtra		: string;
    FHeight		: integer;
    FWidth		: integer;
  protected
    function ReadData(Value: pointer; Size: integer): boolean; override;
    function WriteData(Value: pointer; Size: integer): boolean; override;
    function GetSize: integer; override;
  public
    function GetClipboardFormat: TClipFormat; override;
    procedure Clear; override;
    property URL: string read FURL write FURL;
    property Title: string read FTitle write FTitle;
    property Image: string read FImage write FImage;
    property LowRes: string read FLowRes write FLowRes;
    property Extra: string read FExtra write FExtra;
    property Height: integer read FHeight write FHeight;
    property Width: integer read FWidth write FWidth;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TVCardClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
// Implements support for the '+//ISBN 1-887687-00-9::versit::PDI//vCard'
// (vCard) format.
////////////////////////////////////////////////////////////////////////////////
  TVCardClipboardFormat = class(TCustomStringListClipboardFormat)
  protected
    function ReadData(Value: pointer; Size: integer): boolean; override;
    function WriteData(Value: pointer; Size: integer): boolean; override;
    function GetSize: integer; override;
  public
    function GetClipboardFormat: TClipFormat; override;
    property Items: TStrings read GetLines;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		THTMLClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
// Implements support for the 'HTML Format' format.
////////////////////////////////////////////////////////////////////////////////
  THTMLClipboardFormat = class(TCustomStringListClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    function HasData: boolean; override;
    function Assign(Source: TCustomDataFormat): boolean; override;
    function AssignTo(Dest: TCustomDataFormat): boolean; override;
    property HTML: TStrings read GetLines;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TRFC822ClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TRFC822ClipboardFormat = class(TCustomStringListClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    function Assign(Source: TCustomDataFormat): boolean; override;
    function AssignTo(Dest: TCustomDataFormat): boolean; override;
    property Text: TStrings read GetLines;
  end;


////////////////////////////////////////////////////////////////////////////////
//
//		TURLDataFormat
//
////////////////////////////////////////////////////////////////////////////////
// Renderer for URL formats.
////////////////////////////////////////////////////////////////////////////////
  TURLDataFormat = class(TCustomDataFormat)
  private
    FURL		: string;
    FTitle		: string;
    procedure SetTitle(const Value: string);
    procedure SetURL(const Value: string);
  protected
  public
    function Assign(Source: TClipboardFormat): boolean; override;
    function AssignTo(Dest: TClipboardFormat): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function NeedsData: boolean; override;
    property URL: string read FURL write SetURL;
    property Title: string read FTitle write SetTitle;
  end;


////////////////////////////////////////////////////////////////////////////////
//
//		THTMLDataFormat
//
////////////////////////////////////////////////////////////////////////////////
// Renderer for HTML text data.
////////////////////////////////////////////////////////////////////////////////
  THTMLDataFormat = class(TCustomDataFormat)
  private
    FHTML: TStrings;
    procedure SetHTML(const Value: TStrings);
  protected
  public
    constructor Create(AOwner: TDragDropComponent); override;
    destructor Destroy; override;
    function Assign(Source: TClipboardFormat): boolean; override;
    function AssignTo(Dest: TClipboardFormat): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function NeedsData: boolean; override;
    property HTML: TStrings read FHTML write SetHTML;
  end;


////////////////////////////////////////////////////////////////////////////////
//
//		TOutlookMailDataFormat
//
////////////////////////////////////////////////////////////////////////////////
// Renderer for Microsoft Outlook email formats.
////////////////////////////////////////////////////////////////////////////////
(*
  TOutlookMessage = class;

  TOutlookAttachments = class(TObject)
  public
    property Attachments[Index: integer]: TOutlookMessage; default;
    property Count: integer;
  end;

  TOutlookMessage = class(TObject)
  public
    property Text: string;
    property Stream: IStream;
    property Attachments: TOutlookAttachments;
  end;
*)
  TOutlookMailDataFormat = class(TCustomDataFormat)
  private
    FStorages		: TStorageInterfaceList;
  protected
  public
    constructor Create(AOwner: TDragDropComponent); override;
    destructor Destroy; override;
    function Assign(Source: TClipboardFormat): boolean; override;
    function AssignTo(Dest: TClipboardFormat): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function NeedsData: boolean; override;
    property Storages: TStorageInterfaceList read FStorages;
    // property Streams: TStreamInterfaceList;
    // property Messages: TOutlookAttachments;
  end;


////////////////////////////////////////////////////////////////////////////////
//
//		TDropURLTarget
//
////////////////////////////////////////////////////////////////////////////////
// URL drop target component.
////////////////////////////////////////////////////////////////////////////////
  TDropURLTarget = class(TCustomDropMultiTarget)
  private
    FURLFormat		: TURLDataFormat;
  protected
    function GetTitle: string;
    function GetURL: string;
    function GetPreferredDropEffect: LongInt; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    property URL: string read GetURL;
    property Title: string read GetTitle;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDropURLSource
//
////////////////////////////////////////////////////////////////////////////////
// URL drop source component.
////////////////////////////////////////////////////////////////////////////////
  TDropURLSource = class(TCustomDropMultiSource)
  private
    FURLFormat		: TURLDataFormat;
    procedure SetTitle(const Value: string);
    procedure SetURL(const Value: string);
  protected
    function GetTitle: string;
    function GetURL: string;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    property URL: string read GetURL write SetURL;
    property Title: string read GetTitle write SetTitle;
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
function GetURLFromFile(const Filename: string; var URL: string): boolean;
function GetURLFromStream(Stream: TStream; var URL: string): boolean;
function ConvertURLToFilename(const url: string): string;

function IsHTML(const s: string): boolean;
function MakeHTML(const s: string): string;


////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
//
//			IMPLEMENTATION
//
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
implementation

uses
  SysUtils,
  ShlObj,
  DragDropFile,
  DragDropPIDL;

////////////////////////////////////////////////////////////////////////////////
//
//		Component registration
//
////////////////////////////////////////////////////////////////////////////////
procedure Register;
begin
  RegisterComponents(DragDropComponentPalettePage, [TDropURLTarget,
    TDropURLSource]);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		Utilities
//
////////////////////////////////////////////////////////////////////////////////
function GetURLFromFile(const Filename: string; var URL: string): boolean;
var
  Stream		: TStream;
begin
  Stream := TFileStream.Create(Filename, fmOpenRead or fmShareDenyWrite);
  try
    Result := GetURLFromStream(Stream, URL);
  finally
    Stream.Free;
  end;
end;

function GetURLFromString(const s: string; var URL: string): boolean;
var
  Stream		: TMemoryStream;
begin
  Stream := TMemoryStream.Create;
  try
    Stream.Size := Length(s);
    Move(PChar(s)^, Stream.Memory^, Length(s));
    Result := GetURLFromStream(Stream, URL);
  finally
    Stream.Free;
  end;
end;

const
  // *** DO NOT LOCALIZE ***
  InternetShortcut	= '[InternetShortcut]';
  InternetShortcutExt	= '.url';

function GetURLFromStream(Stream: TStream; var URL: string): boolean;
var
  URLfile		: TStringList;
  i			: integer;
  s			: string;
  p			: PChar;
begin
  Result := False;
  URLfile := TStringList.Create;
  try
    URLFile.LoadFromStream(Stream);
    i := 0;
    while (i < URLFile.Count-1) do
    begin
      if (CompareText(URLFile[i], InternetShortcut) = 0) then
      begin
        inc(i);
        while (i < URLFile.Count) do
        begin
          s := URLFile[i];
          p := PChar(s);
          if (StrLIComp(p, 'URL=', length('URL=')) = 0) then
          begin
            inc(p, length('URL='));
            URL := p;
            Result := True;
            exit;
          end else
            if (p^ = '[') then
              exit;
          inc(i);
        end;
      end;
      inc(i);
    end;
  finally
    URLFile.Free;
  end;
end;

function ConvertURLToFilename(const url: string): string;
const
  Invalids	: set of char
  		= ['\', '/', ':', '?', '*', '<', '>', ',', '|', '''', '"'];
var
  i: integer;
  LastInvalid: boolean;
begin
  Result := url;
  if (AnsiStrLIComp(PChar(lowercase(Result)), 'http://', 7) = 0) then
    delete(Result, 1, 7)
  else if (AnsiStrLIComp(PChar(lowercase(Result)), 'ftp://', 6) = 0) then
    delete(Result, 1, 6)
  else if (AnsiStrLIComp(PChar(lowercase(Result)), 'mailto:', 7) = 0) then
    delete(Result, 1, 7)
  else if (AnsiStrLIComp(PChar(lowercase(Result)), 'file:', 5) = 0) then
    delete(Result, 1, 5);

  if (length(Result) > 120) then
    SetLength(Result, 120);

  // Truncate at first slash
  i := pos('/', Result);
  if (i > 0) then
    SetLength(Result, i-1);

  // Replace invalids with spaces.
  // If string starts with invalids, they are trimmed.
  LastInvalid := True;
  for i := length(Result) downto 1 do
    if (Result[i] in Invalids) then
    begin
      if (not LastInvalid) then
      begin
        Result[i] := ' ';
        LastInvalid := True;
      end else
        // Repeating invalids are trimmed.
        Delete(Result, i, 1);
    end else
      LastInvalid := False;

  if Result = '' then
    Result := 'untitled';

   Result := Result+InternetShortcutExt;
end;

function IsHTML(const s: string): boolean;
begin
  Result := (pos('<HTML>', Uppercase(s)) > 0);
end;

function MakeHTML(const s: string): string;
begin
  { TODO -oanme -cImprovement : Needs to escape special chars in text to HTML conversion. }
  { TODO -oanme -cImprovement : Needs better text to HTML conversion. }
  if (not IsHTML(s)) then
    Result := '<HTML>'#13#10'<BODY>'#13#10 + s + #13#10'</BODY>'#13#10'</HTML>'
  else
    Result := s;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TURLClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_URL: TClipFormat = 0;

function TURLClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_URL = 0) then
    CF_URL := RegisterClipboardFormat(CFSTR_SHELLURL);
  Result := CF_URL;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TNetscapeBookmarkClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_NETSCAPEBOOKMARK: TClipFormat = 0;

function TNetscapeBookmarkClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_NETSCAPEBOOKMARK = 0) then
    CF_NETSCAPEBOOKMARK := RegisterClipboardFormat('Netscape Bookmark'); // *** DO NOT LOCALIZE ***
  Result := CF_NETSCAPEBOOKMARK;
end;

function TNetscapeBookmarkClipboardFormat.GetSize: integer;
begin
  Result := 0;
  if (FURL <> '') then
  begin
    inc(Result, 1024);
    if (FTitle <> '') then
      inc(Result, 1024);
  end;
end;

function TNetscapeBookmarkClipboardFormat.ReadData(Value: pointer;
  Size: integer): boolean;
begin
  // Note: No check for missing string terminator!
  FURL := PChar(Value);
  if (Size > 1024) then
  begin
    inc(PChar(Value), 1024);
    FTitle := PChar(Value);
  end;
  Result := True;
end;

function TNetscapeBookmarkClipboardFormat.WriteData(Value: pointer;
  Size: integer): boolean;
begin
  StrLCopy(Value, PChar(FURL), Size);
  dec(Size, 1024);
  if (Size > 0) and (FTitle <> '') then
  begin
    inc(PChar(Value), 1024);
    StrLCopy(Value, PChar(FTitle), Size);
  end;
  Result := True;
end;

procedure TNetscapeBookmarkClipboardFormat.Clear;
begin
  FURL := '';
  FTitle := '';
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TNetscapeImageClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_NETSCAPEIMAGE: TClipFormat = 0;

function TNetscapeImageClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_NETSCAPEIMAGE = 0) then
    CF_NETSCAPEIMAGE := RegisterClipboardFormat('Netscape Image Format');
  Result := CF_NETSCAPEIMAGE;
end;

type
  TNetscapeImageRec = record
    Size		,
    _Unknown1		,
    Width		,
    Height		,
    HorMargin		,
    VerMargin		,
    Border		,
    OfsLowRes		,
    OfsTitle		,
    OfsURL		,
    OfsExtra		: DWORD
  end;
  PNetscapeImageRec = ^TNetscapeImageRec;

function TNetscapeImageClipboardFormat.GetSize: integer;
begin
  Result := SizeOf(TNetscapeImageRec);
  inc(Result, Length(FImage)+1);

  if (FLowRes <> '') then
    inc(Result, Length(FLowRes)+1);
  if (FTitle <> '') then
    inc(Result, Length(FTitle)+1);
  if (FUrl <> '') then
    inc(Result, Length(FUrl)+1);
  if (FExtra <> '') then
    inc(Result, Length(FExtra)+1);
end;

function TNetscapeImageClipboardFormat.ReadData(Value: pointer;
  Size: integer): boolean;
begin
  Result := (Size > SizeOf(TNetscapeImageRec));
  if (Result) then
  begin
    FWidth := PNetscapeImageRec(Value)^.Width;
    FHeight := PNetscapeImageRec(Value)^.Height;
    FImage := PChar(Value) + SizeOf(TNetscapeImageRec);
    if (PNetscapeImageRec(Value)^.OfsLowRes <> 0) then
      FLowRes := PChar(Value) + PNetscapeImageRec(Value)^.OfsLowRes;
    if (PNetscapeImageRec(Value)^.OfsTitle <> 0) then
      FTitle := PChar(Value) + PNetscapeImageRec(Value)^.OfsTitle;
    if (PNetscapeImageRec(Value)^.OfsURL <> 0) then
      FUrl := PChar(Value) + PNetscapeImageRec(Value)^.OfsUrl;
    if (PNetscapeImageRec(Value)^.OfsExtra <> 0) then
      FExtra := PChar(Value) + PNetscapeImageRec(Value)^.OfsExtra;
  end;
end;

function TNetscapeImageClipboardFormat.WriteData(Value: pointer;
  Size: integer): boolean;
var
  NetscapeImageRec		: PNetscapeImageRec;
begin
  Result := (Size > SizeOf(TNetscapeImageRec));
  if (Result) then
  begin
    NetscapeImageRec := PNetscapeImageRec(Value);
    NetscapeImageRec^.Width := FWidth;
    NetscapeImageRec^.Height := FHeight;
    inc(PChar(Value), SizeOf(TNetscapeImageRec));
    dec(Size, SizeOf(TNetscapeImageRec));
    StrLCopy(Value, PChar(FImage), Size);
    dec(Size, Length(FImage)+1);
    if (Size <= 0) then
      exit;
    if (FLowRes <> '') then
    begin
      StrLCopy(Value, PChar(FLowRes), Size);
      NetscapeImageRec^.OfsLowRes := integer(Value) - integer(NetscapeImageRec);
      dec(Size, Length(FLowRes)+1);
      inc(PChar(Value), Length(FLowRes)+1);
      if (Size <= 0) then
        exit;
    end;
    if (FTitle <> '') then
    begin
      StrLCopy(Value, PChar(FTitle), Size);
      NetscapeImageRec^.OfsTitle := integer(Value) - integer(NetscapeImageRec);
      dec(Size, Length(FTitle)+1);
      inc(PChar(Value), Length(FTitle)+1);
      if (Size <= 0) then
        exit;
    end;
    if (FUrl <> '') then
    begin
      StrLCopy(Value, PChar(FUrl), Size);
      NetscapeImageRec^.OfsUrl := integer(Value) - integer(NetscapeImageRec);
      dec(Size, Length(FUrl)+1);
      inc(PChar(Value), Length(FUrl)+1);
      if (Size <= 0) then
        exit;
    end;
    if (FExtra <> '') then
    begin
      StrLCopy(Value, PChar(FExtra), Size);
      NetscapeImageRec^.OfsExtra := integer(Value) - integer(NetscapeImageRec);
      dec(Size, Length(FExtra)+1);
      inc(PChar(Value), Length(FExtra)+1);
      if (Size <= 0) then
        exit;
    end;
  end;
end;

procedure TNetscapeImageClipboardFormat.Clear;
begin
  FURL := '';
  FTitle := '';
  FImage := '';
  FLowRes := '';
  FExtra := '';
  FHeight := 0;
  FWidth := 0;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TVCardClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_VCARD: TClipFormat = 0;

function TVCardClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_VCARD = 0) then
    CF_VCARD := RegisterClipboardFormat('+//ISBN 1-887687-00-9::versit::PDI//vCard'); // *** DO NOT LOCALIZE ***
  Result := CF_VCARD;
end;

function TVCardClipboardFormat.GetSize: integer;
var
  i			: integer;
begin
  if (Items.Count > 0) then
  begin
    Result := 22; // Length('begin:vcard'+#13+'end:vcard'+#0);
    for i := 0 to Items.Count-1 do
      inc(Result, Length(Items[i])+1);
  end else
    Result := 0;
end;

function TVCardClipboardFormat.ReadData(Value: pointer; Size: integer): boolean;
var
  i			: integer;
  s			: string;
begin
  Result := inherited ReadData(Value, Size);
  if (Result) then
  begin
    // Zap vCard header and trailer
    if (Items.Count > 0) and (CompareText(Items[0], 'begin:vcard') = 0) then
      Items.Delete(0);
    if (Items.Count > 0) and (CompareText(Items[Items.Count-1], 'end:vcard') = 0) then
      Items.Delete(Items.Count-1);
    // Convert to item/value list
    for i := 0 to Items.Count-1 do
      if (pos(':', Items[i]) > 0) then
      begin
        s := Items[i];
        s[pos(':', Items[i])] := '=';
        Items[i] := s;
      end;
  end;
end;

function DOSStringToUnixString(dos: string): string;
var
  s, d			: PChar;
  l			: integer;
begin
  SetLength(Result, Length(dos)+1);
  s := PChar(dos);
  d := PChar(Result);
  l := 1;
  while (s^ <> #0) do
  begin
    // Ignore LF
    if (s^ <> #10) then
    begin
      d^ := s^;
      inc(l);
      inc(d);
    end;
    inc(s);
  end;
  SetLength(Result, l);
end;

function TVCardClipboardFormat.WriteData(Value: pointer; Size: integer): boolean;
var
  s			: string;
begin
  Result := (Items.Count > 0);
  if (Result) then
  begin
    s := DOSStringToUnixString('begin:vcard'+#13+Items.Text+#13+'end:vcard');
    StrLCopy(Value, PChar(s), Size);
  end;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		THTMLClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_HTML: TClipFormat = 0;

function THTMLClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_HTML = 0) then
    CF_HTML := RegisterClipboardFormat('HTML Format');
  Result := CF_HTML;
end;

function THTMLClipboardFormat.HasData: boolean;
begin
  Result := inherited HasData and IsHTML(HTML.Text);
end;

function THTMLClipboardFormat.Assign(Source: TCustomDataFormat): boolean;
begin
  Result := True;
  if (Source is TTextDataFormat) then
    HTML.Text := MakeHTML(TTextDataFormat(Source).Text)
  else
    Result := inherited Assign(Source);
end;

function THTMLClipboardFormat.AssignTo(Dest: TCustomDataFormat): boolean;
begin
  Result := True;
  if (Dest is TTextDataFormat) then
    TTextDataFormat(Dest).Text := HTML.Text
  else
    Result := inherited AssignTo(Dest);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TRFC822ClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_RFC822: TClipFormat = 0;

function TRFC822ClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_RFC822 = 0) then
    CF_RFC822 := RegisterClipboardFormat('Internet Message (rfc822/rfc1522)'); // *** DO NOT LOCALIZE ***
  Result := CF_RFC822;
end;

function TRFC822ClipboardFormat.Assign(Source: TCustomDataFormat): boolean;
begin
  Result := True;
  if (Source is TTextDataFormat) then
    Text.Text := TTextDataFormat(Source).Text
  else
    Result := inherited Assign(Source);
end;

function TRFC822ClipboardFormat.AssignTo(Dest: TCustomDataFormat): boolean;
begin
  Result := True;
  if (Dest is TTextDataFormat) then
    TTextDataFormat(Dest).Text := Text.Text
  else
    Result := inherited AssignTo(Dest);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TURLDataFormat
//
////////////////////////////////////////////////////////////////////////////////
function TURLDataFormat.Assign(Source: TClipboardFormat): boolean;
var
  s			: string;
begin
  Result := False;
  (*
  ** TURLClipboardFormat
  *)
  if (Source is TURLClipboardFormat) then
  begin
    if (FURL = '') then
      FURL := TURLClipboardFormat(Source).URL;
    Result := True;
  end else
  (*
  ** TTextClipboardFormat
  *)
  if (Source is TTextClipboardFormat) then
  begin
    if (FURL = '') then
    begin
      s := TTextClipboardFormat(Source).Text;
      // Convert from text if the string looks like an URL
      if (pos('://', s) > 1) then
      begin
        FURL := s;
        Result := True;
      end;
    end;
  end else
  (*
  ** TFileClipboardFormat
  *)
  if (Source is TFileClipboardFormat) then
  begin
    if (FURL = '') then
    begin
      s := TFileClipboardFormat(Source).Files[0];
      // Convert from Internet Shortcut file format.
      if (CompareText(ExtractFileExt(s), InternetShortcutExt) = 0) and
        (GetURLFromFile(s, FURL)) then
      begin
        if (FTitle = '') then
          FTitle := ChangeFileExt(ExtractFileName(s), '');
        Result := True;
      end;
    end;
  end else
  (*
  ** TFileContentsClipboardFormat
  *)
  if (Source is TFileContentsClipboardFormat) then
  begin
    if (FURL = '') then
    begin
      s := TFileContentsClipboardFormat(Source).Data;
      Result := GetURLFromString(s, FURL);
    end;
  end else
  (*
  ** TFileGroupDescritorClipboardFormat
  *)
  if (Source is TFileGroupDescritorClipboardFormat) then
  begin
    if (FTitle = '') then
    begin
      if (TFileGroupDescritorClipboardFormat(Source).FileGroupDescriptor^.cItems > 0) then
      begin
        // Extract the title of an Internet Shortcut
        s := TFileGroupDescritorClipboardFormat(Source).FileGroupDescriptor^.fgd[0].cFileName;
        if (CompareText(ExtractFileExt(s), InternetShortcutExt) = 0) then
        begin
          FTitle := ChangeFileExt(s, '');
          Result := True;
        end;
      end;
    end;
  end else
  (*
  ** TNetscapeBookmarkClipboardFormat
  *)
  if (Source is TNetscapeBookmarkClipboardFormat) then
  begin
    if (FURL = '') then
      FURL := TNetscapeBookmarkClipboardFormat(Source).URL;
    if (FTitle = '') then
      FTitle := TNetscapeBookmarkClipboardFormat(Source).Title;
    Result := True;
  end else
  (*
  ** TNetscapeImageClipboardFormat
  *)
  if (Source is TNetscapeImageClipboardFormat) then
  begin
    if (FURL = '') then
      FURL := TNetscapeImageClipboardFormat(Source).URL;
    if (FTitle = '') then
      FTitle := TNetscapeImageClipboardFormat(Source).Title;
    Result := True;
  end else
    Result := inherited Assign(Source);
end;

function TURLDataFormat.AssignTo(Dest: TClipboardFormat): boolean;
var
  FGD			: TFileGroupDescriptor;
  s			: string;
begin
  Result := True;
  (*
  ** TURLClipboardFormat
  *)
  if (Dest is TURLClipboardFormat) then
  begin
    TURLClipboardFormat(Dest).URL := FURL;
  end else
  (*
  ** TTextClipboardFormat
  *)
  if (Dest is TTextClipboardFormat) then
  begin
    TTextClipboardFormat(Dest).Text := FURL;
  end else
  (*
  ** TFileContentsClipboardFormat
  *)
  if (Dest is TFileContentsClipboardFormat) then
  begin
    TFileContentsClipboardFormat(Dest).Data := InternetShortcut + #13#10 +
      'URL='+FURL + #13#10;
  end else
  (*
  ** TFileGroupDescritorClipboardFormat
  *)
  if (Dest is TFileGroupDescritorClipboardFormat) then
  begin
    FillChar(FGD, SizeOf(FGD), 0);
    FGD.cItems := 1;
    if (FTitle = '') then
      s := FURL
    else
      s := FTitle;
    StrLCopy(@FGD.fgd[0].cFileName[0], PChar(ConvertURLToFilename(s)),
      SizeOf(FGD.fgd[0].cFileName));
    FGD.fgd[0].dwFlags := FD_LINKUI;
    TFileGroupDescritorClipboardFormat(Dest).CopyFrom(@FGD);
  end else
  (*
  ** TNetscapeBookmarkClipboardFormat
  *)
  if (Dest is TNetscapeBookmarkClipboardFormat) then
  begin
    TNetscapeBookmarkClipboardFormat(Dest).URL := FURL;
    TNetscapeBookmarkClipboardFormat(Dest).Title := FTitle;
  end else
  (*
  ** TNetscapeImageClipboardFormat
  *)
  if (Dest is TNetscapeImageClipboardFormat) then
  begin
    TNetscapeImageClipboardFormat(Dest).URL := FURL;
    TNetscapeImageClipboardFormat(Dest).Title := FTitle;
  end else
    Result := inherited AssignTo(Dest);
end;

procedure TURLDataFormat.Clear;
begin
  Changing;
  FURL := '';
  FTitle := '';
end;

procedure TURLDataFormat.SetTitle(const Value: string);
begin
  Changing;
  FTitle := Value;
end;

procedure TURLDataFormat.SetURL(const Value: string);
begin
  Changing;
  FURL := Value;
end;

function TURLDataFormat.HasData: boolean;
begin
  Result := (FURL <> '') or (FTitle <> '');
end;

function TURLDataFormat.NeedsData: boolean;
begin
  Result := (FURL = '') or (FTitle = '');
end;


////////////////////////////////////////////////////////////////////////////////
//
//		THTMLDataFormat
//
////////////////////////////////////////////////////////////////////////////////
function THTMLDataFormat.Assign(Source: TClipboardFormat): boolean;
begin
  Result := True;

  if (Source is THTMLClipboardFormat) then
    FHTML.Assign(THTMLClipboardFormat(Source).HTML)

  else
    Result := inherited Assign(Source);
end;

function THTMLDataFormat.AssignTo(Dest: TClipboardFormat): boolean;
begin
  Result := True;

  if (Dest is THTMLClipboardFormat) then
    THTMLClipboardFormat(Dest).HTML.Assign(FHTML)

  else
    Result := inherited AssignTo(Dest);
end;

procedure THTMLDataFormat.Clear;
begin
  Changing;
  FHTML.Clear;
end;

constructor THTMLDataFormat.Create(AOwner: TDragDropComponent);
begin
  inherited Create(AOwner);
  FHTML := TStringList.Create;
end;

destructor THTMLDataFormat.Destroy;
begin
  FHTML.Free;
  inherited Destroy;
end;

function THTMLDataFormat.HasData: boolean;
begin
  Result := (FHTML.Count > 0);
end;

function THTMLDataFormat.NeedsData: boolean;
begin
  Result := (FHTML.Count = 0);
end;

procedure THTMLDataFormat.SetHTML(const Value: TStrings);
begin
  FHTML.Assign(Value);
end;

////////////////////////////////////////////////////////////////////////////////
//
//		TOutlookMailDataFormat
//
////////////////////////////////////////////////////////////////////////////////
constructor TOutlookMailDataFormat.Create(AOwner: TDragDropComponent);
begin
  inherited Create(AOwner);
  FStorages := TStorageInterfaceList.Create;
  FStorages.OnChanging := DoOnChanging;
end;

destructor TOutlookMailDataFormat.Destroy;
begin
  Clear;
  FStorages.Free;
  inherited Destroy;
end;

procedure TOutlookMailDataFormat.Clear;
begin
  Changing;
  FStorages.Clear;
end;

function TOutlookMailDataFormat.Assign(Source: TClipboardFormat): boolean;
begin
  Result := True;

  if (Source is TFileContentsStorageClipboardFormat) then
    FStorages.Assign(TFileContentsStorageClipboardFormat(Source).Storages)

  else
    Result := inherited Assign(Source);
end;

function TOutlookMailDataFormat.AssignTo(Dest: TClipboardFormat): boolean;
begin
  Result := True;

  if (Dest is TFileContentsStorageClipboardFormat) then
    TFileContentsStorageClipboardFormat(Dest).Storages.Assign(FStorages)

  else
    Result := inherited AssignTo(Dest);
end;

function TOutlookMailDataFormat.HasData: boolean;
begin
  Result := (FStorages.Count > 0);
end;

function TOutlookMailDataFormat.NeedsData: boolean;
begin
  Result := (FStorages.Count = 0);
end;



////////////////////////////////////////////////////////////////////////////////
//
//		TDropURLTarget
//
////////////////////////////////////////////////////////////////////////////////

constructor TDropURLTarget.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  DragTypes := [dtCopy, dtLink];
  GetDataOnEnter := True;

  FURLFormat := TURLDataFormat.Create(Self);
end;

destructor TDropURLTarget.Destroy;
begin
  FURLFormat.Free;
  inherited Destroy;
end;

function TDropURLTarget.GetTitle: string;
begin
  Result := FURLFormat.Title;
end;

function TDropURLTarget.GetURL: string;
begin
  Result := FURLFormat.URL;
end;

function TDropURLTarget.GetPreferredDropEffect: LongInt;
begin
  Result := GetPreferredDropEffect;
  if (Result = DROPEFFECT_NONE) then
    Result := DROPEFFECT_LINK;
end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDropURLSource
//
////////////////////////////////////////////////////////////////////////////////
constructor TDropURLSource.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  DragTypes := [dtCopy, dtLink];
  PreferredDropEffect := DROPEFFECT_LINK;

  FURLFormat := TURLDataFormat.Create(Self);
end;

destructor TDropURLSource.Destroy;
begin
  FURLFormat.Free;
  inherited Destroy;
end;

function TDropURLSource.GetTitle: string;
begin
  Result := FURLFormat.Title;
end;

procedure TDropURLSource.SetTitle(const Value: string);
begin
  FURLFormat.Title := Value;
end;

function TDropURLSource.GetURL: string;
begin
  Result := FURLFormat.URL;
end;

procedure TDropURLSource.SetURL(const Value: string);
begin
  FURLFormat.URL := Value;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		Initialization/Finalization
//
////////////////////////////////////////////////////////////////////////////////
initialization
  // Data format registration
  TURLDataFormat.RegisterDataFormat;
  THTMLDataFormat.RegisterDataFormat;
  // Clipboard format registration
  TURLDataFormat.RegisterCompatibleFormat(TNetscapeBookmarkClipboardFormat, 0, csSourceTarget, [ddRead]);
  TURLDataFormat.RegisterCompatibleFormat(TNetscapeImageClipboardFormat, 1, csSourceTarget, [ddRead]);
  TURLDataFormat.RegisterCompatibleFormat(TFileContentsClipboardFormat, 2, csSourceTarget, [ddRead]);
  TURLDataFormat.RegisterCompatibleFormat(TFileGroupDescritorClipboardFormat, 2, csSourceTarget, [ddRead]);
  TURLDataFormat.RegisterCompatibleFormat(TURLClipboardFormat, 2, csSourceTarget, [ddRead]);
  TURLDataFormat.RegisterCompatibleFormat(TTextClipboardFormat, 3, csSourceTarget, [ddRead]);
  TURLDataFormat.RegisterCompatibleFormat(TFileClipboardFormat, 4, [csTarget], [ddRead]);

  THTMLDataFormat.RegisterCompatibleFormat(THTMLClipboardFormat, 0, csSourceTarget, [ddRead]);

  TTextDataFormat.RegisterCompatibleFormat(TRFC822ClipboardFormat, 1, csSourceTarget, [ddRead]);
  TTextDataFormat.RegisterCompatibleFormat(THTMLClipboardFormat, 2, csSourceTarget, [ddRead]);

finalization
  // Clipboard format unregistration
  TNetscapeBookmarkClipboardFormat.UnregisterClipboardFormat;
  TNetscapeImageClipboardFormat.UnregisterClipboardFormat;
  TURLClipboardFormat.UnregisterClipboardFormat;
  TVCardClipboardFormat.UnregisterClipboardFormat;
  THTMLClipboardFormat.UnregisterClipboardFormat;
  TRFC822ClipboardFormat.UnregisterClipboardFormat;

  // Target format unregistration
  TURLDataFormat.UnregisterDataFormat;
end.

