unit DropURLTarget;

  // -----------------------------------------------------------------------------
  // Project:         Drag and Drop Target Components
  // Component Names: TDropURLTarget
  // Module:          DropURLTarget
  // Description:     Implements Dragging & Dropping of URLs
  //                  TO your application from another.
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
  DropSource, DropTarget, ActiveX;

type
  TDropURLTarget = class(TDropTarget)
  private
    URLFormatEtc,
    FileContentsFormatEtc: TFormatEtc;
    fURL: String;
  protected
    procedure ClearData; override;
    function DoGetData: boolean; override;
    function HasValidFormats: boolean; override;
  public
    constructor Create(AOwner: TComponent); override;
    property URL: String Read fURL Write fURL;
  end;

procedure Register;

implementation

var
  CF_URL: UINT; //see initialization.

procedure Register;
begin
  RegisterComponents('DragDrop', [TDropURLTarget]);
end;

//******************* GetURLFromFile *************************
function GetURLFromFile(const Filename: string; var URL: string): boolean;
var
  URLfile: textfile;
  str: string;
  i: integer;
begin
  result := false;
  AssignFile(URLFile, Filename);
  try
    Reset(URLFile);
    ReadLn(URLFile, str);
    CloseFile(URLFile);
    if (copy(str,1,18) <> '[InternetShortcut]') then
      exit;
    i := pos('=',str);
    if (i <> 23) and (i <> 24) then exit; // Netscape and IE are different!
    result := true;
    URL := copy(str,i+1,250);
  except
  end;
end;

// -----------------------------------------------------------------------------
//			TDropURLTarget
// -----------------------------------------------------------------------------

//******************* TDropURLTarget.Create *************************
constructor TDropURLTarget.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  DragTypes := [dtLink]; //Only allow links.
  GetDataOnEnter := true;
  with URLFormatEtc do
  begin
    cfFormat := CF_URL;
    ptd := nil;
    dwAspect := DVASPECT_CONTENT;
    lindex := -1;
    tymed := TYMED_HGLOBAL;
  end;
  with FileContentsFormatEtc do
  begin
    cfFormat := CF_FILECONTENTS;
    ptd := nil;
    dwAspect := DVASPECT_CONTENT;
    lindex := 0;
    tymed := TYMED_HGLOBAL;
  end;
end;

//This demonstrates how to enumerate all DataObject formats.
//******************* TDropURLTarget.HasValidFormats *************************
function TDropURLTarget.HasValidFormats: boolean;
var
  GetNum, GotNum: longint;
  FormatEnumerator: IEnumFormatEtc;
  tmpFormatEtc: TformatEtc;
begin
  result := false;
  //Enumerate available DataObject formats
  //to see if any one of the wanted format is available...
  if (fDataObj.EnumFormatEtc(DATADIR_GET,FormatEnumerator) <> S_OK) or
     (FormatEnumerator.Reset <> S_OK) then
    exit;
  GetNum := 1; //get one at a time...
  while (FormatEnumerator.Next(GetNum, tmpFormatEtc, @GotNum) = S_OK) and
        (GetNum = GotNum) do
    with tmpFormatEtc do
      if (ptd = nil) and (dwAspect = DVASPECT_CONTENT) and
         {(lindex <> -1) or} (tymed and TYMED_HGLOBAL <> 0) and
         ((cfFormat = CF_URL) or (cfFormat = CF_FILECONTENTS) or
         (cfFormat = CF_HDROP) or (cfFormat = CF_TEXT)) then
      begin
        result := true;
        break;
      end;
end;

//******************* TDropURLTarget.ClearData *************************
procedure TDropURLTarget.ClearData;
begin
  fURL := '';
end;

//******************* TDropURLTarget.DoGetData *************************
function TDropURLTarget.DoGetData: boolean;
var
  medium: TStgMedium;
  cText: pchar;
  tmpFiles: TStringList;
begin
  fURL := '';
  result := false;
  //--------------------------------------------------------------------------
  if (fDataObj.GetData(URLFormatEtc, medium) = S_OK) then
  begin
    try
      if (medium.tymed <> TYMED_HGLOBAL) then
        exit;
      cText := PChar(GlobalLock(medium.HGlobal));
      fURL := cText;
      GlobalUnlock(medium.HGlobal);
      result := true;
    finally
      ReleaseStgMedium(medium);
    end;
  end
  //--------------------------------------------------------------------------
  else if (fDataObj.GetData(TextFormatEtc, medium) = S_OK) then
  begin
    try
      if (medium.tymed <> TYMED_HGLOBAL) then
        exit;
      cText := PChar(GlobalLock(medium.HGlobal));
      fURL := cText;
      GlobalUnlock(medium.HGlobal);
      result := true;
    finally
      ReleaseStgMedium(medium);
    end;
  end
  //--------------------------------------------------------------------------
  else if (fDataObj.GetData(FileContentsFormatEtc, medium) = S_OK) then
  begin
    try
      if (medium.tymed <> TYMED_HGLOBAL) then
        exit;
      cText := PChar(GlobalLock(medium.HGlobal));
      fURL := cText;
      fURL := copy(fURL,24,250);
      GlobalUnlock(medium.HGlobal);
      result := true;
    finally
      ReleaseStgMedium(medium);
    end;
  end
  //--------------------------------------------------------------------------
  else if (fDataObj.GetData(HDropFormatEtc, medium) = S_OK) then
  begin
    try
      if (medium.tymed <> TYMED_HGLOBAL) then exit;
      tmpFiles := TStringList.create;
      try
        if GetFilesFromHGlobal(medium.HGlobal,TStrings(tmpFiles)) then
        begin
          if (lowercase(ExtractFileExt(tmpFiles[0])) = '.url') and
             GetURLFromFile(tmpFiles[0], fURL) then
            result := true;
        end;
      finally
        tmpFiles.free;
      end;
    finally
      ReleaseStgMedium(medium);
    end;
  end;
end;

initialization
  CF_URL := RegisterClipboardFormat('UniformResourceLocator');

end.
