unit DropBMPTarget;

// -----------------------------------------------------------------------------
// Project:         Drag and Drop Source Components
// Component Names: TDropBMPSource
// Module:          DropBMPSource
// Description:     Implements Dragging & Dropping of Bitmaps
//                  FROM your application to another.
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

// -----------------------------------------------------------------------------
// Acknowledgement:
// Thanks to Dieter Steinwedel for some help with DIBs.
// http://godard.oec.uni-osnabrueck.de/student_home/dsteinwe/delphi/DietersDelphiSite.htm
// -----------------------------------------------------------------------------

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  dropsource, DropTarget, ActiveX;

type
  TDropBMPTarget = class(TDropTarget)
  private
    fBitmap: TBitmap;
  protected
    procedure ClearData; override;
    function DoGetData: boolean; override;
    function HasValidFormats: boolean; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    property Bitmap: TBitmap Read fBitmap;
  end;

procedure Register;

implementation

const
  DIBFormatEtc: TFormatEtc = (cfFormat: CF_DIB;
    ptd: nil; dwAspect: DVASPECT_CONTENT; lindex: -1; tymed: TYMED_HGLOBAL);
  BMPFormatEtc: TFormatEtc = (cfFormat: CF_BITMAP;
    ptd: nil; dwAspect: DVASPECT_CONTENT; lindex: -1; tymed: TYMED_GDI);
  PalFormatEtc: TFormatEtc = (cfFormat: CF_PALETTE;
    ptd: nil; dwAspect: DVASPECT_CONTENT; lindex: -1; tymed: TYMED_GDI);

procedure Register;
begin
  RegisterComponents('DragDrop', [TDropBMPTarget]);
end;

// -----------------------------------------------------------------------------
//	Miscellaneous DIB Functions
//  (Modified from Graphics.pas - © Inprise Corporation.)
// -----------------------------------------------------------------------------

procedure ByteSwapColors(var Colors; Count: Integer);
// convert RGB to BGR and vice-versa.  TRGBQuad <-> TPaletteEntry
// Note: Intel 386s processors not supported.
begin
  asm
        MOV   EDX, Colors
        MOV   ECX, Count
        DEC   ECX
        JS    @@END
  @@1:  MOV   EAX, [EDX+ECX*4]
        BSWAP EAX
        SHR   EAX,8
        MOV   [EDX+ECX*4],EAX
        DEC   ECX
        JNS   @@1
  @@END:
  end;
end;
// ----------------------------------------------------------------------------- 

function PaletteFromDIBColorTable(ColorTable: pointer; ColorCount: Integer): HPalette;
var
  Pal: TMaxLogPalette;
  DC: HDC;
  SysPalSize: Integer;
begin
  Result := 0;
  if (ColorCount <= 2) then Exit;
  Pal.palVersion := $300;
  Pal.palNumEntries := ColorCount;
  Move(ColorTable, Pal.palPalEntry, ColorCount * 4);

  DC := GetDC(0);
  try
    SysPalSize := GetDeviceCaps(DC, SIZEPALETTE);
    if (ColorCount = 16) and (SysPalSize >= 16) then
    begin
      GetSystemPaletteEntries(DC, 0, 8, Pal.palPalEntry);
      GetSystemPaletteEntries(DC, SysPalSize - 8, 8, Pal.palPalEntry[8]);
    end
    else
      ByteSwapColors(Pal.palPalEntry, ColorCount);
  finally
    ReleaseDC(0, DC);
  end;
  Result := CreatePalette(PLogPalette(@Pal)^);
end;
// ----------------------------------------------------------------------------- 

procedure CopyDIBToBitmap(Bitmap:TBitmap; ImagePtr: pointer);
var
  BitmapInfoHeader: PBitmapInfoHeader;
  BitmapInfo: PBitmapInfo;
  ColorCount: integer;
  DIBits: Pointer;
  Pal: HPALETTE;
  DC: HDC;
  OldPal: HPALETTE;
begin
  OldPal := 0;
  BitmapInfo := ImagePtr;
  BitmapInfoHeader := ImagePtr;
  with BitmapInfoHeader^ do
    if biClrUsed <> 0 then
      ColorCount := biClrUsed
    else if biBitCount > 8 then
      ColorCount := 0
    else
      ColorCount := (1 shl biBitCount);

  longint(DIBits) := longint(ImagePtr) +
      SizeOf(TBitmapInfoHeader) + (ColorCount * SizeOf(TRgbQuad));

  {Create the palette }
  Pal := PaletteFromDIBColorTable(@BitmapInfo.bmiColors, ColorCount);
  try
    DC := GetDC(0);
    if DC <> 0 then
    try
      if Pal <> 0 then
      begin
        OldPal := SelectPalette(DC, Pal, False);
        RealizePalette(DC);
      end;
      try
         Bitmap.Handle:=CreateDIBitmap(DC, BitmapInfo^.bmiHeader,
             CBM_INIT, DIBits, BitmapInfo^, DIB_RGB_COLORS);
      finally
         if OldPal <> 0 then SelectPalette(DC, OldPal, False);
      end;
    finally
      ReleaseDC(0, DC);
    end;
  finally
    if Pal <> 0 then DeleteObject(Pal);
  end;
end;

// -----------------------------------------------------------------------------
//			TDropBMPTarget
// -----------------------------------------------------------------------------

constructor TDropBMPTarget.Create( AOwner: TComponent );
begin
   inherited Create( AOwner );
   fBitmap := TBitmap.Create;
end;
// ----------------------------------------------------------------------------- 

destructor TDropBMPTarget.Destroy;
begin
  fBitmap.Free;
  inherited Destroy;
end;
// ----------------------------------------------------------------------------- 

function TDropBMPTarget.HasValidFormats: boolean;
begin
  result := (DataObject.QueryGetData(DIBFormatEtc) = S_OK) or
                (DataObject.QueryGetData(BMPFormatEtc) = S_OK);
end;
// ----------------------------------------------------------------------------- 

procedure TDropBMPTarget.ClearData;
begin
  fBitmap.handle := 0;
end;
// ----------------------------------------------------------------------------- 

function TDropBMPTarget.DoGetData: boolean;
var
  medium, medium2: TStgMedium;
  DIBData: pointer;
begin
  result := false;
  //--------------------------------------------------------------------------
  if (DataObject.GetData(DIBFormatEtc, medium) = S_OK) then
  begin
    if (medium.tymed = TYMED_HGLOBAL) then
    begin
      DIBData := GlobalLock(medium.HGlobal);
      try
        CopyDIBToBitmap(fBitmap, DIBData);
        result := true;
      finally
        GlobalUnlock(medium.HGlobal);
      end;
    end;
    ReleaseStgMedium(medium);
  end
  //--------------------------------------------------------------------------
  else if (DataObject.GetData(BMPFormatEtc, medium) = S_OK) then
  begin
    try
      if (medium.tymed <> TYMED_GDI) then exit;
      if (DataObject.GetData(PalFormatEtc, medium2) = S_OK) then
      begin
        try
          fBitmap.LoadFromClipboardFormat(CF_BITMAP, Medium.hBitmap, Medium2.hBitmap);
        finally
          ReleaseStgMedium(medium2);
        end;
      end
      else
        fBitmap.LoadFromClipboardFormat(CF_BITMAP, Medium.hBitmap, 0);
      result := true;
    finally
      ReleaseStgMedium(medium);
    end;
  end
  //--------------------------------------------------------------------------
  else
    result := false;
end;
// ----------------------------------------------------------------------------- 
// ----------------------------------------------------------------------------- 

end.
