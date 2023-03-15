unit DropPIDLSource;

// -----------------------------------------------------------------------------
// Project:         Drag and Drop Source Components
// Component Names: TDropPIDLSource
// Module:          DropPIDLSource
// Description:     Implements Dragging & Dropping of PIDLs
//                  FROM your application to another.
// Version:	       3.4
// Date:            19-FEB-1999
// Target:          Win32, Delphi 3 & 4, CB3
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
function PidlToString(pidl: PItemIDList): String;

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

// -----------------------------------------------------------------------------
//			TDropPIDLSource
// -----------------------------------------------------------------------------

//******************* TDropPIDLSource.Create *************************
constructor TDropPIDLSource.Create(aOwner: TComponent);
begin
  inherited Create(aOwner);
  fPIDLs := TStringList.create;
  AddFormatEtc(CF_IDLIST, NIL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL);
  AddFormatEtc(CF_PREFERREDDROPEFFECT, NIL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL);
end;

//******************* TDropPIDLSource.Destroy *************************
destructor TDropPIDLSource.Destroy;
begin
  fPIDLs.free;
  inherited Destroy;
end;

//Note: Once the PIDL has been copied into the list it can be 'freed'.
//******************* TDropPIDLSource.CopyFolderPidlToList *************************
procedure TDropPIDLSource.CopyFolderPidlToList(pidl: PItemIDList);
begin
  fPIDLs.clear;
  fPIDLs.add(PidlToString(pidl));
end;

//Note: Once the PIDL has been copied into the list it can be 'freed'.
//******************* TDropPIDLSource.CopyFilePidlToList *************************
procedure TDropPIDLSource.CopyFilePidlToList(pidl: PItemIDList);
begin
  if fPIDLs.count < 1 then exit; //no folder pidl has been added!
  fPIDLs.add(PidlToString(pidl));
end;

//******************* TDropPIDLSource.CutOrCopyToClipboard *************************
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

//******************* TDropPIDLSource.DoGetData *************************
function TDropPIDLSource.DoGetData(const FormatEtcIn: TFormatEtc; OUT Medium: TStgMedium):HRESULT;
var
  i, MemSpace, CidaSize, Offset: integer;
  pCIDA: PIDA;
  pInt: ^UINT;
  pOffset: PChar;
  DropEffect: ^DWORD;
begin
  Medium.tymed := 0;
  Medium.UnkForRelease := NIL;
  Medium.hGlobal := 0;
  if fPIDLs.count < 2 then result := E_UNEXPECTED
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
 else
    result := DV_E_FORMATETC;
end;
// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------

//initialization
//  CF_IDLIST is 'registered' in DropSource

end.
