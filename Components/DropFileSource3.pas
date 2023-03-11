unit DropFileSource3;

// -----------------------------------------------------------------------------
//
//			*** NOT FOR RELEASE ***
//
//                     *** INTERNAL USE ONLY ***
//
// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite
// Module:          DropFileSource3
// Description:     Test case for deprecated TDropSource class.
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
  DragDropPIDL,
  DragDropFormats,
  DragDropFile,
  DropSource3,
  ActiveX, Classes;

{$include DragDrop.inc}

type
  TDropFileSourceX = class(TDropSource)
  private
    fFiles: TStrings;
    fMappedNames: TStrings;
    FFileClipboardFormat: TFileClipboardFormat;
    FPIDLClipboardFormat: TPIDLClipboardFormat;
    FPreferredDropEffectClipboardFormat: TPreferredDropEffectClipboardFormat;
    FFilenameMapClipboardFormat: TFilenameMapClipboardFormat;
    FFilenameMapWClipboardFormat: TFilenameMapWClipboardFormat;

    procedure SetFiles(files: TStrings);
    procedure SetMappedNames(names: TStrings);
  protected
    function DoGetData(const FormatEtcIn: TFormatEtc;
      out Medium: TStgMedium):HRESULT; override;
    function CutOrCopyToClipboard: boolean; override;
  public
    constructor Create(aOwner: TComponent); override;
    destructor Destroy; override;
  published
    property Files: TStrings read fFiles write SetFiles;
    //MappedNames is only needed if files need to be renamed during a drag op
    //eg dragging from 'Recycle Bin'.
    property MappedNames: TStrings read fMappedNames write SetMappedNames;
  end;

procedure Register;

implementation

uses
  Windows,
  ShlObj,
  SysUtils,
  ClipBrd;

procedure Register;
begin
  RegisterComponents(DragDropComponentPalettePage, [TDropFileSourceX]);
end;

// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------

constructor TDropFileSourceX.Create(aOwner: TComponent);
begin
  inherited Create(aOwner);
  fFiles := TStringList.Create;
  fMappedNames := TStringList.Create;

  FFileClipboardFormat := TFileClipboardFormat.Create;
  FPIDLClipboardFormat := TPIDLClipboardFormat.Create;
  FPreferredDropEffectClipboardFormat := TPreferredDropEffectClipboardFormat.Create;
  FFilenameMapClipboardFormat := TFilenameMapClipboardFormat.Create;
  FFilenameMapWClipboardFormat := TFilenameMapWClipboardFormat.Create;

  AddFormatEtc(FFileClipboardFormat.GetClipboardFormat, NIL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL);
  AddFormatEtc(FPIDLClipboardFormat.GetClipboardFormat, NIL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL);
  AddFormatEtc(FPreferredDropEffectClipboardFormat.GetClipboardFormat, NIL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL);
  AddFormatEtc(FFilenameMapClipboardFormat.GetClipboardFormat, NIL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL);
  AddFormatEtc(FFilenameMapWClipboardFormat.GetClipboardFormat, NIL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL);
end;
// -----------------------------------------------------------------------------

destructor TDropFileSourceX.destroy;
begin
  FFileClipboardFormat.Free;
  FPIDLClipboardFormat.Free;
  FPreferredDropEffectClipboardFormat.Free;
  FFilenameMapClipboardFormat.Free;
  FFilenameMapWClipboardFormat.Free;
  fFiles.Free;
  fMappedNames.free;
  inherited Destroy;
end;
// -----------------------------------------------------------------------------

procedure TDropFileSourceX.SetFiles(files: TStrings);
begin
  fFiles.assign(files);
end;
// -----------------------------------------------------------------------------

procedure TDropFileSourceX.SetMappedNames(names: TStrings);
begin
  fMappedNames.assign(names);
end;
// -----------------------------------------------------------------------------

function TDropFileSourceX.CutOrCopyToClipboard: boolean;
var
  FormatEtcIn: TFormatEtc;
  Medium: TStgMedium;
begin
  FormatEtcIn.cfFormat := CF_HDROP;
  FormatEtcIn.dwAspect := DVASPECT_CONTENT;
  FormatEtcIn.tymed := TYMED_HGLOBAL;
  if (Files.count = 0) then result := false
  else if GetData(formatetcIn,Medium) = S_OK then
  begin
    Clipboard.SetAsHandle(CF_HDROP,Medium.hGlobal);
    result := true;
  end else result := false;
end;
// -----------------------------------------------------------------------------

function TDropFileSourceX.DoGetData(const FormatEtcIn: TFormatEtc;
         out Medium: TStgMedium):HRESULT;
begin
  Medium.tymed := 0;
  Medium.UnkForRelease := NIL;
  Medium.hGlobal := 0;

  result := E_UNEXPECTED;
  if fFiles.count = 0 then
    exit;

  //--------------------------------------------------------------------------
  if FFileClipboardFormat.AcceptFormat(FormatEtcIn) then
  begin
    FFileClipboardFormat.Files.Assign(FFiles);
    if FFileClipboardFormat.SetDataToMedium(FormatEtcIn, Medium) then
      result := S_OK;
  end else
  //--------------------------------------------------------------------------
  if FFilenameMapClipboardFormat.AcceptFormat(FormatEtcIn) then
  begin
    FFilenameMapClipboardFormat.FileMaps.Assign(fMappedNames);
    if FFilenameMapClipboardFormat.SetDataToMedium(FormatEtcIn, Medium) then
      result := S_OK;
  end else
  //--------------------------------------------------------------------------
  if FFilenameMapWClipboardFormat.AcceptFormat(FormatEtcIn) then
  begin
    FFilenameMapWClipboardFormat.FileMaps.Assign(fMappedNames);
    if FFilenameMapWClipboardFormat.SetDataToMedium(FormatEtcIn, Medium) then
      result := S_OK;
  end else
  //--------------------------------------------------------------------------
  if FPIDLClipboardFormat.AcceptFormat(FormatEtcIn) then
  begin
    FPIDLClipboardFormat.Filenames.Assign(FFiles);
    if FPIDLClipboardFormat.SetDataToMedium(FormatEtcIn, Medium) then
      result := S_OK;
  end else
  //--------------------------------------------------------------------------
  //This next format does not work for Win95 but should for Win98, WinNT ...
  //It stops the shell from prompting (with a popup menu) for the choice of
  //Copy/Move/Shortcut when performing a file 'Shortcut' onto Desktop or Explorer.
  if FPreferredDropEffectClipboardFormat.AcceptFormat(FormatEtcIn) then
  begin
    FPreferredDropEffectClipboardFormat.Value := FeedbackEffect;
    if FPreferredDropEffectClipboardFormat.SetDataToMedium(FormatEtcIn, Medium) then
      result := S_OK;
  end else
    result := DV_E_FORMATETC;
end;

end.
