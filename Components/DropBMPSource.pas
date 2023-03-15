unit DropBMPSource;

// -----------------------------------------------------------------------------
// Project:         Drag and Drop Source Components
// Component Names: TDropBMPSource
// Module:          DropBMPSource
// Description:     Implements Dragging & Dropping of Bitmaps
//                  FROM your application to another.
// Version:	       3.4
// Date:            17-FEB-1999
// Target:          Win32, Delphi 3 & 4, CB3
// Authors:         Angus Johnson,   ajohnson@rpi.net.au
//                  Anders Melander, anders@melander.dk
//                                   http://www.melander.dk
//                  Graham Wideman,  graham@sdsu.edu
//                                   http://www.wideman-one.com
// Copyright        ©1997-99 Angus Johnson, Anders Melander & Graham Wideman
// -----------------------------------------------------------------------------

// -----------------------------------------------------------------------------
// Acknowledgement:
// Thanks to Dieter Steinwedel for some help with DIBs.
// http://godard.oec.uni-osnabrueck.de/student_home/dsteinwe/delphi/DietersDelphiSite.htm
// -----------------------------------------------------------------------------

interface

uses
  Windows, SysUtils, Classes, Graphics, Controls, dropsource, ActiveX, ClipBrd;

type
  TDropBMPSource = class(TDropSource)
  private
    fBitmap: TBitmap;
    procedure SetBitmap(Bmp: TBitmap);
  protected
    function DoGetData(const FormatEtcIn: TFormatEtc; OUT Medium: TStgMedium):HRESULT; Override;
  public
    constructor Create(aOwner: TComponent); Override;
    destructor Destroy; Override;
    function CutOrCopyToClipboard: boolean; Override;
  published
    property Bitmap: TBitmap Read fBitmap Write SetBitmap;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('DragDrop', [TDropBMPSource]);
end;

// -----------------------------------------------------------------------------
//     Miscellaneous DIB Functions
//    (Modified from Graphic.pas - ©Inprise Corporation.)
// -----------------------------------------------------------------------------

//******************* BytesPerScanline *************************
function BytesPerScanline(PixelsPerScanline, BitsPerPixel, Alignment: Longint): Longint;
begin
  Dec(Alignment);
  Result := ((PixelsPerScanline * BitsPerPixel) + Alignment) and not Alignment;
  Result := Result div 8;
end;

//******************* InitializeBitmapInfoHeader *************************
procedure InitializeBitmapInfoHeader(Bitmap: HBITMAP; var BI: TBitmapInfoHeader);
var
  DS: TDIBSection;
  Bytes: Integer;
begin
  FillChar(BI, sizeof(BI), 0);
  DS.dsbmih.biSize := 0;
  Bytes := GetObject(Bitmap, SizeOf(DS), @DS);
  if Bytes = 0 then exit;
  if (Bytes >= (sizeof(DS.dsbm) + sizeof(DS.dsbmih))) and
    (DS.dsbmih.biSize >= sizeof(DS.dsbmih)) then
    BI := DS.dsbmih
  else
  begin
    with BI, DS.dsbm do
    begin
      biSize := SizeOf(BI);
      biWidth := bmWidth;
      biHeight := bmHeight;
    end;
  end;
  BI.biBitCount := DS.dsbm.bmBitsPixel * DS.dsbm.bmPlanes;
  BI.biPlanes := 1;
  if BI.biSizeImage = 0 then
    BI.biSizeImage := BytesPerScanLine(BI.biWidth, BI.biBitCount, 32) * Abs(BI.biHeight);
end;

//******************* InternalGetDIBSizes *************************
procedure InternalGetDIBSizes(Bitmap: HBITMAP; var InfoHeaderSize: Integer;
  var ImageSize: DWORD);
var
  BI: TBitmapInfoHeader;
begin
  InitializeBitmapInfoHeader(Bitmap, BI);
  if BI.biSize = 0 then exit; //error

  if BI.biBitCount > 8 then
  begin
    InfoHeaderSize := SizeOf(TBitmapInfoHeader);
    if (BI.biCompression and BI_BITFIELDS) <> 0 then
      Inc(InfoHeaderSize, 12);
  end
  else
    InfoHeaderSize := SizeOf(TBitmapInfoHeader) + SizeOf(TRGBQuad) *
      (1 shl BI.biBitCount);
  ImageSize := BI.biSizeImage;
end;

//******************* InternalGetDIB *************************
function InternalGetDIB(Bitmap: HBITMAP;
   Palette: HPALETTE; var BitmapInfo; var Bits): Boolean;
var
  OldPal: HPALETTE;
  DC: HDC;
begin
  Result := false;
  InitializeBitmapInfoHeader(Bitmap, TBitmapInfoHeader(BitmapInfo));
  if TBitmapInfoHeader(BitmapInfo).biSize = 0 then exit; //error

  OldPal := 0;
  DC := CreateCompatibleDC(0);
  try
    if Palette <> 0 then
    begin
      OldPal := SelectPalette(DC, Palette, False);
      RealizePalette(DC);
    end;
    Result := GetDIBits(DC, Bitmap, 0, TBitmapInfoHeader(BitmapInfo).biHeight, @Bits,
      TBitmapInfo(BitmapInfo), DIB_RGB_COLORS) <> 0;
  finally
    if OldPal <> 0 then SelectPalette(DC, OldPal, False);
    DeleteDC(DC);
  end;
end;

//******************* GetHGlobalDIBFromBitmap *************************
function GetHGlobalDIBFromBitmap(Src: HBITMAP; Pal: HPALETTE): HGlobal;
var HeaderSize: Integer;
    ImageSize: DWORD;
    DIBHeader, DIBBits: Pointer;
    HasFailed: boolean;
begin
     Result := 0;
     if Src=0 then exit;
     InternalGetDIBSizes(Src, HeaderSize, ImageSize);
     Result:=GlobalAlloc(GMEM_MOVEABLE or GMEM_ZEROINIT or GMEM_SHARE, HeaderSize+integer(ImageSize));
     if Result=0 then exit; //Out of memory.
     try
       DIBHeader:=GlobalLock(Result);
       if DIBHeader = nil then
       begin
         GlobalFree(Result);
         Result := 0;
         exit;
       end;
       HasFailed := false;
       try
         DIBBits:=Pointer(Longint(DIBHeader) + HeaderSize);
         if not InternalGetDIB(Src, Pal, DIBHeader^, DIBBits^) then
           HasFailed := true;
       finally
         GlobalUnlock(Result);
         if HasFailed then
         begin
           GlobalFree(Result);
           Result := 0;
         end;
       end;
     except
       GlobalFree(Result);
       Result := 0;
     end;
end;

// -----------------------------------------------------------------------------
//			TDropBMPSource
// -----------------------------------------------------------------------------

//******************* TDropBMPSource.Create *************************
constructor TDropBMPSource.Create(aOwner: TComponent);
begin
  inherited Create(aOwner);
  fbitmap := TBitmap.create;
  DragTypes := [dtCopy]; // Default to Copy

  AddFormatEtc(CF_BITMAP, NIL, DVASPECT_CONTENT, -1, TYMED_GDI);
  AddFormatEtc(CF_PALETTE, NIL, DVASPECT_CONTENT, -1, TYMED_GDI);
  AddFormatEtc(CF_DIB, NIL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL);
end;

//******************* TDropBMPSource.Destroy *************************
destructor TDropBMPSource.destroy;
begin
  fBitmap.Free;
  inherited Destroy;
end;

//******************* TDropBMPSource.SetBitmap *************************
procedure TDropBMPSource.SetBitmap(Bmp: TBitmap);
begin
  fBitmap.assign(Bmp);
end;

//******************* TDropBMPSource.CutOrCopyToClipboard *************************
function TDropBMPSource.CutOrCopyToClipboard: boolean;
var
  data: HGlobal;
begin
  result := false;
  if fBitmap.empty then exit;
  try
    data := GetHGlobalDIBFromBitmap(fBitmap.handle,fBitmap.palette);
    if data = 0 then exit;
    Clipboard.SetAsHandle(CF_DIB,data);
    result := true;
  except
    raise Exception.create('Unable to copy BMP to clipboard.');
  end;
end;

//******************* TDropBMPSource.DoGetData *************************
function TDropBMPSource.DoGetData(const FormatEtcIn: TFormatEtc; OUT Medium: TStgMedium):HRESULT;
var
  fmt: WORD;
  pal: HPALETTE;
begin

  Medium.tymed := 0;
  Medium.UnkForRelease := nil;
  Medium.HGlobal := 0;
  //--------------------------------------------------------------------------
  if not fBitmap.empty and
     (FormatEtcIn.cfFormat = CF_DIB) and
     (FormatEtcIn.dwAspect = DVASPECT_CONTENT) and
     (FormatEtcIn.tymed and TYMED_HGLOBAL <> 0) then
  begin
    try
      Medium.HGlobal := GetHGlobalDIBFromBitmap(fBitmap.handle,fBitmap.palette);
      if Medium.HGlobal <> 0 then
      begin
        Medium.tymed := TYMED_HGLOBAL;
        result := S_OK
      end else
        result := E_OUTOFMEMORY;
    except
      result := E_OUTOFMEMORY;
    end;
  end
  //--------------------------------------------------------------------------
  else if not fBitmap.empty and
     (FormatEtcIn.cfFormat = CF_BITMAP) and
     (FormatEtcIn.dwAspect = DVASPECT_CONTENT) and
     (FormatEtcIn.tymed and TYMED_GDI <> 0) then
  begin
    try
      //This next line just gets a copy of the bitmap handle...
      fBitmap.SaveToClipboardFormat(fmt, THandle(Medium.hBitmap), pal);
      if pal <> 0 then DeleteObject(pal);
      if Medium.hBitmap <> 0 then
      begin
        Medium.tymed := TYMED_GDI;
        result := S_OK
      end else
        result := E_OUTOFMEMORY;
    except
      result := E_OUTOFMEMORY;
    end;
  end
  //--------------------------------------------------------------------------
  else if not fBitmap.empty and
     (FormatEtcIn.cfFormat = CF_PALETTE) and
     (FormatEtcIn.dwAspect = DVASPECT_CONTENT) and
     (FormatEtcIn.tymed and TYMED_GDI <> 0) then
  begin
    try
      Medium.hBitmap := CopyPalette(fBitmap.palette);
      if Medium.hBitmap <> 0 then
      begin
        Medium.tymed := TYMED_GDI;
        result := S_OK
      end else
        result := E_OUTOFMEMORY;
    except
      result := E_OUTOFMEMORY;
    end {try}
  end else
    result := DV_E_FORMATETC;
end;

end.
