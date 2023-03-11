unit DropSource3;

// -----------------------------------------------------------------------------
//
//			*** NOT FOR RELEASE ***
//
// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite
// Module:          DropSource3
// Description:     Deprecated TDropSource class.
//                  Provided for compatibility with previous versions of the
//                  Drag and Drop Component Suite.
// Version:         4.0
// Date:            25-JUN-2000
// Target:          Win32, Delphi 3-6 and C++ Builder 3-5
// Authors:         Angus Johnson, ajohnson@rpi.net.au
//                  Anders Melander, anders@melander.dk, http://www.melander.dk
// Copyright        © 1997-2000 Angus Johnson & Anders Melander
// -----------------------------------------------------------------------------


interface

uses
  DragDrop,
  DropSource,
  ActiveX,
  Classes;

{$include DragDrop.inc}

const
  MAXFORMATS = 20;

type
  // TODO -oanme -cStopShip : Verify that TDropSource can be used for pre v4 components.
  TDropSource = class(TCustomDropSource)
  private
    FDataFormats: array[0..MAXFORMATS-1] of TFormatEtc;
    FDataFormatsCount: integer;

  protected
    // IDataObject implementation
    function QueryGetData(const FormatEtc: TFormatEtc): HRESULT; stdcall;

    // TCustomDropSource implementation
    function HasFormat(const FormatEtc: TFormatEtc): boolean; override;
    function GetEnumFormatEtc(dwDirection: LongInt): IEnumFormatEtc; override;

    // New functions...
    procedure AddFormatEtc(cfFmt: TClipFormat; pt: PDVTargetDevice;
      dwAsp, lInd, tym: longint); virtual;

  public
    constructor Create(AOwner: TComponent); override;
  end;

implementation

uses
  ShlObj,
  SysUtils,
  Windows;

// -----------------------------------------------------------------------------
//			TEnumFormatEtc
// -----------------------------------------------------------------------------

type

  pFormatList = ^TFormatList;
  TFormatList = array[0..255] of TFormatEtc;

  TEnumFormatEtc = class(TInterfacedObject, IEnumFormatEtc)
  private
    FFormatList: pFormatList;
    FFormatCount: Integer;
    FIndex: Integer;
  public
    constructor Create(FormatList: pFormatList; FormatCount, Index: Integer);
    { IEnumFormatEtc }
    function Next(Celt: LongInt; out Elt; pCeltFetched: pLongInt): HRESULT; stdcall;
    function Skip(Celt: LongInt): HRESULT; stdcall;
    function Reset: HRESULT; stdcall;
    function Clone(out Enum: IEnumFormatEtc): HRESULT; stdcall;
  end;
// -----------------------------------------------------------------------------

constructor TEnumFormatEtc.Create(FormatList: pFormatList;
            FormatCount, Index: Integer);
begin
  inherited Create;
  FFormatList := FormatList;
  FFormatCount := FormatCount;
  FIndex := Index;
end;
// -----------------------------------------------------------------------------

function TEnumFormatEtc.Next(Celt: LongInt;
  out Elt; pCeltFetched: pLongInt): HRESULT;
var
  i: Integer;
begin
  i := 0;
  WHILE (i < Celt) and (FIndex < FFormatCount) do
  begin
    TFormatList(Elt)[i] := FFormatList[fIndex];
    Inc(FIndex);
    Inc(i);
  end;
  if pCeltFetched <> NIL then pCeltFetched^ := i;
  if i = Celt then result := S_OK else result := S_FALSE;
end;
// -----------------------------------------------------------------------------

function TEnumFormatEtc.Skip(Celt: LongInt): HRESULT;
begin
  if Celt <= FFormatCount - FIndex then
  begin
    FIndex := FIndex + Celt;
    result := S_OK;
  end else
  begin
    FIndex := FFormatCount;
    result := S_FALSE;
  end;
end;
// -----------------------------------------------------------------------------

function TEnumFormatEtc.ReSet: HRESULT;
begin
  fIndex := 0;
  result := S_OK;
end;
// -----------------------------------------------------------------------------

function TEnumFormatEtc.Clone(out Enum: IEnumFormatEtc): HRESULT;
begin
  enum := TEnumFormatEtc.Create(FFormatList, FFormatCount, FIndex);
  result := S_OK;
end;

// -----------------------------------------------------------------------------
//			TDropSource
// -----------------------------------------------------------------------------

constructor TDropSource.Create(AOwner: TComponent);
begin
  inherited Create(aOwner);
  FDataFormatsCount := 0;
end;
// -----------------------------------------------------------------------------

function TDropSource.QueryGetData(const FormatEtc: TFormatEtc): HRESULT; stdcall;
var
  i: integer;
begin
  result:= S_OK;
  for i := 0 to FDataFormatsCount-1 do
    with FDataFormats[i] do
    begin
      if (FormatEtc.cfFormat = cfFormat) and
         (FormatEtc.dwAspect = dwAspect) and
         (FormatEtc.tymed and tymed <> 0) then exit; //result:= S_OK;
    end;
  result:= E_FAIL;
end;
// -----------------------------------------------------------------------------

function TDropSource.GetEnumFormatEtc(dwDirection: Integer): IEnumFormatEtc;
begin
  if (dwDirection = DATADIR_GET) then
    Result := TEnumFormatEtc.Create(pFormatList(@FDataFormats), FDataFormatsCount, 0)
  else
    result := nil;
end;
// -----------------------------------------------------------------------------

procedure TDropSource.AddFormatEtc(cfFmt: TClipFormat;
  pt: PDVTargetDevice; dwAsp, lInd, tym: longint);
begin
  if fDataFormatsCount = MAXFORMATS then exit;

  FDataFormats[fDataFormatsCount].cfFormat := cfFmt;
  FDataFormats[fDataFormatsCount].ptd := pt;
  FDataFormats[fDataFormatsCount].dwAspect := dwAsp;
  FDataFormats[fDataFormatsCount].lIndex := lInd;
  FDataFormats[fDataFormatsCount].tymed := tym;
  inc(FDataFormatsCount);
end;
// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------

function TDropSource.HasFormat(const FormatEtc: TFormatEtc): boolean;
begin
  Result := True;
 { TODO -oanme -cStopShip : TDropSource.HasFormat needs implementation }
end;

initialization
  OleInitialize(NIL);
  ShGetMalloc(ShellMalloc);

finalization
  ShellMalloc := nil;
  OleUninitialize;
end.
