unit DropComboTarget;

(*
 * Drag and Drop Component Suite
 *
 * Copyright (c) Angus Johnson & Anders Melander
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 *)

interface

uses
  Classes,
  Graphics,
  ActiveX,
  DragDrop,
  DropTarget,
  DragDropFormats,
  DragDropInternet,
  DragDropGraphics,
  DragDropFile,
  DragDropText;

type
  // Note: mfCustom is used to support DataFormatAdapters.
  TComboFormatType = (mfText, mfFile, mfURL, mfBitmap, mfMetaFile, mfData, mfCustom);
  TComboFormatTypes = set of TComboFormatType;

const
  AllComboFormats = [mfText, mfFile, mfURL, mfBitmap, mfMetaFile, mfData];

{$HPPEMIT '// Work around for stupid #define in wingdi.h of "GetMetaFile" as "GetMetaFileA".'}
{$HPPEMIT '#pragma option push'}
{$HPPEMIT '#pragma warn -8017'}
{$HPPEMIT '#define GetMetaFile  GetMetaFile'}
{$HPPEMIT '#pragma option pop'}
type
////////////////////////////////////////////////////////////////////////////////
//
//              TDropComboTarget
//
////////////////////////////////////////////////////////////////////////////////
  TDropComboTarget = class(TCustomDropMultiTarget)
  private
    FFileFormat: TFileDataFormat;
    FFileMapFormat: TFileMapDataFormat;
    FURLFormat: TURLDataFormat;
    FBitmapFormat: TBitmapDataFormat;
    FMetaFileFormat: TMetaFileDataFormat;
    FTextFormat: TTextDataFormat;
    FDataFormat: TDataStreamDataFormat;
    FFormats: TComboFormatTypes;
  protected
    procedure DoAcceptFormat(const DataFormat: TCustomDataFormat;
      var Accept: boolean); override;
    function GetFiles: TUnicodeStrings;
    function GetTitle: string;
    function GetURL: AnsiString;
    function GetBitmap: TBitmap;
    function GetMetaFile: TMetaFile;
    function GetText: string;
    function GetFileMaps: TUnicodeStrings;
    function GetStreams: TStreamList;
  public
    constructor Create(AOwner: TComponent); override;
    property Files: TUnicodeStrings read GetFiles;
    property FileMaps: TUnicodeStrings read GetFileMaps;
    property URL: AnsiString read GetURL;
    property Title: string read GetTitle;
    property Bitmap: TBitmap read GetBitmap;
    property MetaFile: TMetaFile read GetMetaFile;
    property Text: string read GetText;
    property Data: TStreamList read GetStreams;
  published
    property OnAcceptFormat;
    property Formats: TComboFormatTypes read FFormats write FFormats default AllComboFormats;
  end;



implementation

////////////////////////////////////////////////////////////////////////////////
//
//              TDropComboTarget
//
////////////////////////////////////////////////////////////////////////////////
constructor TDropComboTarget.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FFileFormat := TFileDataFormat.Create(Self);
  FURLFormat := TURLDataFormat.Create(Self);
  FBitmapFormat := TBitmapDataFormat.Create(Self);
  FMetaFileFormat := TMetaFileDataFormat.Create(Self);
  FTextFormat := TTextDataFormat.Create(Self);
  FFileMapFormat := TFileMapDataFormat.Create(Self);
  FDataFormat := TDataStreamDataFormat.Create(Self);
  FFormats := AllComboFormats;
end;

procedure TDropComboTarget.DoAcceptFormat(const DataFormat: TCustomDataFormat;
  var Accept: boolean);
begin
  if (Accept) then
  begin
    if (DataFormat is TFileDataFormat) or (DataFormat is TFileMapDataFormat) then
      Accept := (mfFile in FFormats)
    else if (DataFormat is TURLDataFormat) then
      Accept := (mfURL in FFormats)
    else if (DataFormat is TBitmapDataFormat) then
      Accept := (mfBitmap in FFormats)
    else if (DataFormat is TMetaFileDataFormat) then
      Accept := (mfMetaFile in FFormats)
    else if (DataFormat is TTextDataFormat) then
      Accept := (mfText in FFormats)
    else if (DataFormat is TDataStreamDataFormat) then
      Accept := (mfData in FFormats)
    else
      Accept := (mfCustom in FFormats)
  end;

  if (Accept) then
    inherited DoAcceptFormat(DataFormat, Accept);

end;

function TDropComboTarget.GetBitmap: TBitmap;
begin
  Result := FBitmapFormat.Bitmap;
end;

function TDropComboTarget.GetFileMaps: TUnicodeStrings;
begin
  Result := FFileMapFormat.FileMaps;
end;

function TDropComboTarget.GetFiles: TUnicodeStrings;
begin
  Result := FFileFormat.Files;
end;

function TDropComboTarget.GetMetaFile: TMetaFile;
begin
  Result := FMetaFileFormat.MetaFile;
end;

function TDropComboTarget.GetStreams: TStreamList;
begin
  Result := FDataFormat.Streams;
end;

function TDropComboTarget.GetText: string;
begin
  Result := FTextFormat.Text;
end;

function TDropComboTarget.GetTitle: string;
begin
  Result := FURLFormat.Title;
end;

function TDropComboTarget.GetURL: AnsiString;
begin
  Result := FURLFormat.URL;
end;

end.
