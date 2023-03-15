unit DropURLSource;

  // -----------------------------------------------------------------------------
  // Project:         Drag and Drop Source Components
  // Component Names: TDropURLSource
  // Module:          DropURLSource
  // Description:     Implements Dragging & Dropping of URLs
  //                  FROM your application to another.
  // Version:	       3.3
  // Date:            30-OCT-1998
  // Target:          Win32, Delphi 3 & 4
  // Author:          Angus Johnson,   ajohnson@rpi.net.au
  // Copyright        ©1998 Angus Johnson
  // -----------------------------------------------------------------------------
  // You are free to use this source but please give me credit for my work.
  // if you make improvements or derive new components from this code,
  // I would very much like to see your improvements. FEEDBACK IS WELCOME.
  // -----------------------------------------------------------------------------

  // History:
  // dd/mm/yy  Version  Changes
  // --------  -------  ----------------------------------------
  // 16.11.98  3.3      * Module header added.
  //                    * Improved component icon.
  // 22.10.98  3.2      * Initial release.
  //                     (Ver. No coincides with Component Suite Ver. No.)
  // -----------------------------------------------------------------------------

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  dropsource, ActiveX, ClipBrd, ShlObj;

type
  TDropURLSource = class(TDropSource)
  private
    fURL: String;
  protected
    function DoGetData(const FormatEtcIn: TFormatEtc; OUT Medium: TStgMedium):HRESULT; Override;
  public
    constructor Create(aOwner: TComponent); Override;
    function CopyToClipboard: boolean; Override;
  published
    property URL: String Read fURL Write fURL;
  end;

procedure Register;

implementation

var
  CF_URL: UINT; //see initialization.

procedure Register;
begin
  RegisterComponents('DragDrop', [TDropURLSource]);
end;

//******************* ConvertURLToFilename *************************
function ConvertURLToFilename(url: string): string;
const
  Invalids = '\/:?*<>,|''" ';
var
  i: integer;
begin
  if lowercase(copy(url,1,7)) = 'http://' then
    url := copy(url,8,128) // limit to 120 chars.
  else if lowercase(copy(url,1,6)) = 'ftp://' then
    url := copy(url,7,127)
  else if lowercase(copy(url,1,7)) = 'mailto:' then
    url := copy(url,8,128)
  else if lowercase(copy(url,1,5)) = 'file:' then
    url := copy(url,6,126);

  if url = '' then url := 'untitled';
  result := url;
  for i := 1 to length(result) do
    if result[i] = '/'then
    begin
      result := copy(result,1,i-1);
      break;
    end
    else if pos(result[i],Invalids) <> 0 then
      result[i] := '_';
   appendstr(result,'.url');
end;

// -----------------------------------------------------------------------------
//			TDropURLSource
// -----------------------------------------------------------------------------

//******************* TDropURLSource.Create *************************
constructor TDropURLSource.Create(aOwner: TComponent);
begin
  inherited Create(aOwner);
  fURL := '';
  DragTypes := [dtLink]; // Only dtLink allowed

  AddFormatEtc(CF_URL, NIL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL);
  AddFormatEtc(CF_FILEGROUPDESCRIPTOR, NIL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL);
  AddFormatEtc(CF_FILECONTENTS, NIL, DVASPECT_CONTENT, 0, TYMED_HGLOBAL);
  AddFormatEtc(CF_TEXT, NIL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL);
end;


//******************* TDropURLSource.CopyToClipboard *************************
function TDropURLSource.CopyToClipboard: boolean;
var
  FormatEtcIn: TFormatEtc;
  Medium: TStgMedium;
begin
  FormatEtcIn.cfFormat := CF_URL;
  FormatEtcIn.dwAspect := DVASPECT_CONTENT;
  FormatEtcIn.tymed := TYMED_HGLOBAL;
  if fURL = '' then result := false
  else if GetData(formatetcIn,Medium) = S_OK then
  begin
    Clipboard.SetAsHandle(CF_URL,Medium.hGlobal);
    result := true;
  end else result := false;
end;

//******************* TDropURLSource.DoGetData *************************
function TDropURLSource.DoGetData(const FormatEtcIn: TFormatEtc; OUT Medium: TStgMedium):HRESULT;
const
  URLPrefix = '[InternetShortcut]'#10'URL=';
var
  pFGD: PFileGroupDescriptor;
  pText: PChar;
begin

  Medium.tymed := 0;
  Medium.UnkForRelease := NIL;
  Medium.hGlobal := 0;

  //--------------------------------------------------------------------------
  if ((FormatEtcIn.cfFormat = CF_URL) or (FormatEtcIn.cfFormat = CF_TEXT)) and
    (FormatEtcIn.dwAspect = DVASPECT_CONTENT) and
    (FormatEtcIn.tymed and TYMED_HGLOBAL <> 0) then
  begin
    Medium.hGlobal := GlobalAlloc(GMEM_SHARE or GHND, Length(fURL)+1);
    if (Medium.hGlobal = 0) then
      result := E_OUTOFMEMORY
    else
    begin
      medium.tymed := TYMED_HGLOBAL;
      pText := PChar(GlobalLock(Medium.hGlobal));
      try
        StrCopy(pText, PChar(fURL));
      finally
        GlobalUnlock(Medium.hGlobal);
      end;
      result := S_OK;
    end;
  end
  //--------------------------------------------------------------------------
  else if (FormatEtcIn.cfFormat = CF_FILECONTENTS) and
    (FormatEtcIn.dwAspect = DVASPECT_CONTENT) and
    (FormatEtcIn.tymed and TYMED_HGLOBAL <> 0) then
  begin
    Medium.hGlobal := GlobalAlloc(GMEM_SHARE or GHND, Length(URLPrefix + fURL)+1);
    if (Medium.hGlobal = 0) then
      result := E_OUTOFMEMORY
    else
    begin
      medium.tymed := TYMED_HGLOBAL;
      pText := PChar(GlobalLock(Medium.hGlobal));
      try
        StrCopy(pText, PChar(URLPrefix + fURL));
      finally
        GlobalUnlock(Medium.hGlobal);
      end;
      result := S_OK;
    end;
  end
  //--------------------------------------------------------------------------
  else if (FormatEtcIn.cfFormat = CF_FILEGROUPDESCRIPTOR) and
    (FormatEtcIn.dwAspect = DVASPECT_CONTENT) and
    (FormatEtcIn.tymed and TYMED_HGLOBAL <> 0) then
  begin
    Medium.hGlobal := GlobalAlloc(GMEM_SHARE or GHND, SizeOf(TFileGroupDescriptor));
    if (Medium.hGlobal = 0) then
    begin
      result := E_OUTOFMEMORY;
      Exit;
    end;
    medium.tymed := TYMED_HGLOBAL;
    pFGD := pointer(GlobalLock(Medium.hGlobal));
    try
      with pFGD^ do
      begin
        cItems := 1;
        fgd[0].dwFlags := FD_LINKUI;
        StrPCopy(fgd[0].cFileName,ConvertURLToFilename(fURL));
      end;
    finally
      GlobalUnlock(Medium.hGlobal);
    end;
    result := S_OK;
  //--------------------------------------------------------------------------
  end else
    result := DV_E_FORMATETC;
end;

initialization
  CF_URL := RegisterClipboardFormat('UniformResourceLocator');

end.
