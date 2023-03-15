unit DropURLTarget;

// -----------------------------------------------------------------------------
// Project:         Drag and Drop Target Components
// Component Names: TDropURLTarget
// Module:          DropURLTarget
// Description:     Implements Dragging & Dropping of URLs
//                  TO your application from another.
// Version:	       3.6
// Date:            21-APR-1999
// Target:          Win32, Delphi3, Delphi4, C++ Builder 3, C++ Builder 4
// Authors:         Angus Johnson,   ajohnson@rpi.net.au
//                  Anders Melander, anders@melander.dk
//                                   http://www.melander.dk
//                  Graham Wideman,  graham@sdsu.edu
//                                   http://www.wideman-one.com
// Copyright        ©1997-99 Angus Johnson, Anders Melander & Graham Wideman
// -----------------------------------------------------------------------------

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DropSource, DropTarget, ActiveX, ShlObj;

type
  TDropURLTarget = class(TDropTarget)
  private
    URLFormatEtc, 
    FileContentsFormatEtc,
    FGDFormatEtc: TFormatEtc;
    fURL: String;
    fTitle: String;
  protected
    procedure ClearData; override;
    function DoGetData: boolean; override;
    function HasValidFormats: boolean; override;
  public
    constructor Create(AOwner: TComponent); override;
    property URL: String Read fURL Write fURL;
    property Title: String Read fTitle Write fTitle;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('DragDrop', [TDropURLTarget]);
end;
// ----------------------------------------------------------------------------- 

function GetURLFromFile(const Filename: string; var URL: string): boolean;
var
  URLfile: textfile;
  str: string;
  i: integer;
begin
  //OK, just to get confusing...
  //URL file contents in 2 possible formats...
  //1. '[InternetShortcut]'#13#10'URL=http://....';
  //2. '[InternetShortcut]'#10'URL=http://....';
  result := false;
  AssignFile(URLFile, Filename);
  try
    Reset(URLFile);
    try
      ReadLn(URLFile, str);
      if (copy(str,1,18) <> '[InternetShortcut]') then exit; //error
      if length(str) = 18 then ReadLn(URLFile, str);
      i := pos('=',str);
      if (i = 0) then exit;
      URL := copy(str,i+1,250);
      result := true;
    finally
      CloseFile(URLFile);
    end;
  except
  end;
end;

// -----------------------------------------------------------------------------
//			TDropURLTarget
// -----------------------------------------------------------------------------

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
  with FGDFormatEtc do
  begin
    cfFormat := CF_FILEGROUPDESCRIPTOR;
    ptd := nil;
    dwAspect := DVASPECT_CONTENT;
    lindex := -1;
    tymed := TYMED_HGLOBAL;
  end;
end;
// ----------------------------------------------------------------------------- 

//This demonstrates how to enumerate all DataObject formats.
function TDropURLTarget.HasValidFormats: boolean;
var
  GetNum, GotNum: longint;
  FormatEnumerator: IEnumFormatEtc;
  tmpFormatEtc: TformatEtc;
begin
  result := false;
  //Enumerate available DataObject formats
  //to see if any one of the wanted formats is available...
  if (DataObject.EnumFormatEtc(DATADIR_GET,FormatEnumerator) <> S_OK) or
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
// ----------------------------------------------------------------------------- 

procedure TDropURLTarget.ClearData;
begin
  fURL := '';
end;
// ----------------------------------------------------------------------------- 

function TDropURLTarget.DoGetData: boolean;
var
  medium: TStgMedium;
  cText: pchar;
  tmpFiles: TStringList;
  pFGD: PFileGroupDescriptor;
begin
  fURL := '';
  fTitle := '';
  result := false;
  //--------------------------------------------------------------------------
  if (DataObject.GetData(URLFormatEtc, medium) = S_OK) then
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
  else if (DataObject.GetData(TextFormatEtc, medium) = S_OK) then
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
  else if (DataObject.GetData(FileContentsFormatEtc, medium) = S_OK) then
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
  else if (DataObject.GetData(HDropFormatEtc, medium) = S_OK) then
  begin
    try
      if (medium.tymed <> TYMED_HGLOBAL) then exit;
      tmpFiles := TStringList.create;
      try
        if GetFilesFromHGlobal(medium.HGlobal,TStrings(tmpFiles)) and
          (lowercase(ExtractFileExt(tmpFiles[0])) = '.url') and
             GetURLFromFile(tmpFiles[0], fURL) then
        begin
            fTitle := extractfilename(tmpFiles[0]);
            delete(fTitle,length(fTitle)-3,4); //deletes '.url' extension
            result := true;
        end;
      finally
        tmpFiles.free;
      end;
    finally
      ReleaseStgMedium(medium);
    end;
  end;

  if (DataObject.GetData(FGDFormatEtc, medium) = S_OK) then
  begin
    try
      if (medium.tymed <> TYMED_HGLOBAL) then exit;
      pFGD := pointer(GlobalLock(medium.HGlobal));
      fTitle := pFGD^.fgd[0].cFileName;
      delete(fTitle,length(fTitle)-3,4); //deletes '.url' extension
    finally
      ReleaseStgMedium(medium);
    end;
  end
  else if fTitle = '' then fTitle := fURL;
end;
// ----------------------------------------------------------------------------- 
// ----------------------------------------------------------------------------- 

end.
