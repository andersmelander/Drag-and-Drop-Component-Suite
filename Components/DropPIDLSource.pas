unit DropPIDLSource;

// -----------------------------------------------------------------------------
// Project:         Drag and Drop Source Components
// Component Names: TDropPIDLSource
// Module:          DropPIDLSource
// Description:     Implements Dragging & Dropping of PIDLs
//                  FROM your application to another.
// Version:	       3.5
// Date:            30-MAR-1999
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
  dropsource, ActiveX, ClipBrd, ShlObj;

type
  TDropPIDLSource = class(TDropSource)
  private
    fPIDLs: TStrings; //NOTE: contains folder PIDL as well as file PIDLs
    function GetFilename(index: integer): string; //used internally
  protected
    function DoGetData(const FormatEtcIn: TFormatEtc; OUT Medium: TStgMedium):HRESULT; Override;
    function CutOrCopyToClipboard: boolean; Override;
  public
    constructor Create(aOwner: TComponent); Override;
    destructor Destroy; Override;
    procedure CopyFolderPidlToList(pidl: PItemIDList);
    procedure CopyFilePidlToList(pidl: PItemIDList);
  end;

procedure Register;

//Exported as also used by DropPIDLTarget...
function PidlToString(pidl: PItemIDList): String;
function JoinPidlStrings(pidl1,pidl2: string): String;

var
  CF_FILENAMEMAP: UINT;

implementation

procedure Register;
begin
  RegisterComponents('DragDrop', [TDropPIDLSource]);
end;

// -----------------------------------------------------------------------------
//			Miscellaneous Functions...
// -----------------------------------------------------------------------------

function GetSizeOfPidl(pidl: PItemIDList): integer;
var
  i: integer;
begin
  result := SizeOf(Word);
  repeat
    i := pSHItemID(pidl)^.cb;
    inc(result,i);
    inc(longint(pidl),i);
  until i = 0;
end;

function PidlToString(pidl: PItemIDList): String;
var
  PidlLength: integer;
begin
  PidlLength := GetSizeOfPidl(pidl);
  setlength(result,PidlLength);
  Move(pidl^,pchar(result)^,PidlLength);
end;

function JoinPidlStrings(pidl1,pidl2: string): String;
var
  PidlLength: integer;
begin
  if Length(pidl1) <= 2 then PidlLength := 0
  else PidlLength := Length(pidl1)-2;
  setlength(result,PidlLength+length(pidl2));
  if PidlLength > 0 then Move(pidl1[1],result[1],PidlLength);
  Move(pidl2[1],result[PidlLength+1],length(pidl2));
end;

// -----------------------------------------------------------------------------
//			TDropPIDLSource
// -----------------------------------------------------------------------------

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
constructor TDropPIDLSource.Create(aOwner: TComponent);
begin
  inherited Create(aOwner);
  fPIDLs := TStringList.create;
  AddFormatEtc(CF_HDROP, NIL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL);
  AddFormatEtc(CF_IDLIST, NIL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL);
  AddFormatEtc(CF_PREFERREDDROPEFFECT, NIL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL);
end;

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
destructor TDropPIDLSource.Destroy;
begin
  fPIDLs.free;
  inherited Destroy;
end;

//this function is used internally by DoGetData()...
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
function TDropPIDLSource.GetFilename(index: integer): string;
var
  PidlStr: string;
  buff: array [0..MAX_PATH] of char;
begin
  if (index < 1) or (index >= fPIDLs.count) then result := ''
  else
  begin
    PidlStr := JoinPidlStrings(fPIDLs[0], fPIDLs[index]);
    SHGetPathFromIDList(PItemIDList(pChar(PidlStr)),buff);
    result := buff;
  end;
end;

//Note: Once the PIDL has been copied into the list it can be 'freed'.
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
procedure TDropPIDLSource.CopyFolderPidlToList(pidl: PItemIDList);
begin
  fPIDLs.clear;
  fPIDLs.add(PidlToString(pidl));
end;

//Note: Once the PIDL has been copied into the list it can be 'freed'.
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
procedure TDropPIDLSource.CopyFilePidlToList(pidl: PItemIDList);
begin
  if fPIDLs.count < 1 then exit; //no folder pidl has been added!
  fPIDLs.add(PidlToString(pidl));
end;

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
function TDropPIDLSource.CutOrCopyToClipboard: boolean;
var
  FormatEtcIn: TFormatEtc;
  Medium: TStgMedium;
begin
  FormatEtcIn.cfFormat := CF_IDLIST;
  FormatEtcIn.dwAspect := DVASPECT_CONTENT;
  FormatEtcIn.tymed := TYMED_HGLOBAL;
  if (fPIDLs.count < 2) then result := false
  else if GetData(formatetcIn,Medium) = S_OK then
  begin
    Clipboard.SetAsHandle(CF_IDLIST,Medium.hGlobal);
    result := true;
  end else result := false;
end;

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
function TDropPIDLSource.DoGetData(const FormatEtcIn: TFormatEtc; OUT Medium: TStgMedium):HRESULT;
var
  i, MemSpace, CidaSize, Offset: integer;
  pCIDA: PIDA;
  pInt: ^UINT;
  pOffset: PChar;
  DropEffect: ^DWORD;
  dropfiles: pDropFiles;
  fFiles: string;
  pFileList: PChar;
begin
  Medium.tymed := 0;
  Medium.UnkForRelease := NIL;
  Medium.hGlobal := 0;
  if fPIDLs.count < 2 then result := E_UNEXPECTED
  //--------------------------------------------------------------------------
  else if (FormatEtcIn.cfFormat = CF_HDROP) and
    (FormatEtcIn.dwAspect = DVASPECT_CONTENT) and
    (FormatEtcIn.tymed and TYMED_HGLOBAL <> 0) then
  begin
    fFiles := '';
    for i := 1 to fPIDLs.Count-1 do
      appendstr(fFiles,GetFilename(i)+#0);
    appendstr(fFiles,#0);

    Medium.hGlobal :=
      GlobalAlloc(GMEM_SHARE or GMEM_ZEROINIT, SizeOf(TDropFiles)+length(fFiles));
    if (Medium.hGlobal = 0) then
      result:=E_OUTOFMEMORY
    else
    begin
      Medium.tymed := TYMED_HGLOBAL;
      dropfiles := GlobalLock(Medium.hGlobal);
      try
        dropfiles^.pfiles := SizeOf(TDropFiles);
        dropfiles^.fwide := False;
        longint(pFileList) := longint(dropfiles)+SizeOf(TDropFiles);
        move(fFiles[1],pFileList^,length(fFiles));
      finally
        GlobalUnlock(Medium.hGlobal);
      end;
      result := S_OK;
    end;
  end
  //--------------------------------------------------------------------------
  else if (FormatEtcIn.cfFormat = CF_IDLIST) and
    (FormatEtcIn.dwAspect = DVASPECT_CONTENT) and
    (FormatEtcIn.tymed and TYMED_HGLOBAL <> 0) then
  begin
    CidaSize := sizeof(UINT)*(1+fPIDLs.Count); //size of CIDA structure
    MemSpace := CidaSize;
    for i := 0 to fPIDLs.Count-1 do
      Inc(MemSpace, Length(fPIDLs[i]));
    Medium.hGlobal := GlobalAlloc(GMEM_SHARE or GMEM_ZEROINIT, MemSpace);

    if (Medium.hGlobal = 0) then
      result := E_OUTOFMEMORY
    else
    begin
      medium.tymed := TYMED_HGLOBAL;
      pCIDA := PIDA(GlobalLock(Medium.hGlobal));
      try
        pCIDA^.cidl := fPIDLs.count-1; //don't count folder
        pInt := @(pCIDA^.aoffset); //points to aoffset[0];
        pOffset := pChar(pCIDA);
        inc(pOffset,CidaSize); //pOffset now points to where the Folder PIDL will be stored
        offset := CidaSize;
        for i := 0 to fPIDLs.Count-1 do
        begin
          pInt^ := offset; //put offset into aoffset[i]
          Move(pointer(fPIDLs[i])^,pOffset^,length(fPIDLs[i])); //put the PIDL into pOffset
          inc(offset,length(fPIDLs[i])); //increase the offset by the size of the last pidl
          inc(pInt); //increment the aoffset pointer
          inc(pOffset,length(fPIDLs[i])); //move pOffset ready for the next PIDL
        end;
      finally
        GlobalUnlock(Medium.hGlobal);
      end;
     result := S_OK;
    end;
  end
  //--------------------------------------------------------------------------
  else if (FormatEtcIn.cfFormat = CF_PREFERREDDROPEFFECT) and
    (FormatEtcIn.dwAspect = DVASPECT_CONTENT) and
    (FormatEtcIn.tymed and TYMED_HGLOBAL <> 0) then
  begin
    Medium.tymed := TYMED_HGLOBAL;
    Medium.hGlobal := GlobalAlloc(GMEM_SHARE or GMEM_ZEROINIT, SizeOf(DWORD));
    if Medium.hGlobal = 0 then
      result:=E_OUTOFMEMORY
    else
    begin
      DropEffect := GlobalLock(Medium.hGlobal);
      try
        DropEffect^ := FeedbackEffect;
      finally
        GlobalUnLock(Medium.hGlobal);
      end;
      result := S_OK;
    end;
  end
  //--------------------------------------------------------------------------
  else
    result := DV_E_FORMATETC;
end;
// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------

end.
