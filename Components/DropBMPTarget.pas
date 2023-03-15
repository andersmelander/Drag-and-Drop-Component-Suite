unit DropBMPTarget;

  // -----------------------------------------------------------------------------
  // Project:         Drag and Drop Source Components
  // Component Names: TDropBMPSource
  // Module:          DropBMPSource
  // Description:     Implements Dragging & Dropping of Bitmaps
  //                  FROM your application to another.
  // Version:         3.3
  // Date:            16-NOV-1998
  // Target:          Win32, Delphi 3 & 4
  // Author:          Angus Johnson,   ajohnson@rpi.net.au
  // Copyright        �1998 Angus Johnson
  // -----------------------------------------------------------------------------
  // You are free to use this source but please give me credit for my work.
  // if you make improvements or derive new components from this code,
  // I would very much like to see your improvements. FEEDBACK IS WELCOME.
  // -----------------------------------------------------------------------------

  // Acknowledgements:
  // Thanks to Dieter Steinwedel for some help with DIBs.
  // http://godard.oec.uni-osnabrueck.de/student_home/dsteinwe/delphi/DietersDelphiSite.htm
  // -----------------------------------------------------------------------------

  // History:
  // dd/mm/yy  Version  Changes
  // --------  -------  ----------------------------------------
  // 16.11.98  3.3      * Added DIB support.
  // 22.10.98  3.2      * Initial release.
  //                     (Ver. No coincides with Component Suite Ver. No.)
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
//  (Modified from Graphics.pas - � Inprise Corporation.)
// -----------------------------------------------------------------------------

//******************* ByteSwapColors *************************
procedure ByteSwapColors(var Colors; Count: Integer);
// convert RGB to BGR and vice-versa.  TRGBQuad <-> TPaletteEntry
// Note: Intel 386s processors no longer supported.
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

//******************* PaletteFromDIBColorTable *************************
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

//******************* CopyDIBToBitmap *************************
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
end;

// -----------------------------------------------------------------------------
//			TDropBMPTarget
// -----------------------------------------------------------------------------

//******************* TDropBMPTarget.Create *************************
constructor TDropBMPTarget.Create( AOwner: TComponent );
begin
   inherited Create( AOwner );
   fBitmap := TBitmap.Create;
end;

//******************* TDropBMPTarget.Destroy *************************
destructor TDropBMPTarget.Destroy;
begin
  fBitmap.Free;
  inherited Destroy;
end;

//******************* TDropBMPTarget.HasValidFormats *************************
function TDropBMPTarget.HasValidFormats: boolean;
begin
  result := (fDataObj.QueryGetData(DIBFormatEtc) = S_OK) or
                (fDataObj.QueryGetData(BMPFormatEtc) = S_OK);
end;

//******************* TDropBMPTarget.ClearData *************************
procedure TDropBMPTarget.ClearData;
begin
  fBitmap.handle := 0;
end;

//******************* TDropBMPTarget.DoGetData *************************
function TDropBMPTarget.DoGetData: boolean;
var
  medium, medium2: TStgMedium;
  DIBData: pointer;
begin
  result := false;
  //--------------------------------------------------------------------------
  if (fDataObj.GetData(DIBFormatEtc, medium) = S_OK) then
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
  else if (fDataObj.GetData(BMPFormatEtc, medium) = S_OK) then
  begin
    try
      if (medium.tymed <> TYMED_GDI) then exit;
      if (fDataObj.GetData(PalFormatEtc, medium2) = S_OK) then
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

end.
