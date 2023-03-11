//---------------------------------------------------------------------------
// This demo application was contributed by Jonathan Arnold.
//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "main.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DragDrop"
#pragma link "DragDropText"
#pragma link "DropSource"
#pragma link "DropTarget"
#pragma resource "*.dfm"
TFormAutoScroll *FormAutoScroll;
//---------------------------------------------------------------------------
__fastcall TFormAutoScroll::TFormAutoScroll(TComponent* Owner): TForm(Owner)
{
  int i;

  StringGrid1->ColCount = 'Z'-'A'+1;
  StringGrid1->RowCount = 50+1;
  // Populate header cells.
  for (i = 0; i < StringGrid1->RowCount; ++i)
    StringGrid1->Cells[0][1+i] = AnsiString(i+1);
  for (i = 0; i < StringGrid1->ColCount; ++i)
    StringGrid1->Cells[1+i][0] = char(i+'A');

  // Populate the grid with data.
  for (i = 1; i < StringGrid1->RowCount; ++i)
    for (int j = 1; j < StringGrid1->ColCount; ++j)
      StringGrid1->Cells[j][i] = AnsiString(char('A'+j-1))+AnsiString(i);

  // Slow auto scroll down to 1 scroll every 100mS.
  // The default value is 50.
  DragDropScrollInterval = 100;
}
//---------------------------------------------------------------------------
void __fastcall TFormAutoScroll::PanelSourceMouseDown(TObject *Sender,
  TMouseButton Button, TShiftState Shift, int X, int Y)
{
  if (DragDetectPlus(reinterpret_cast<THandle>(Handle), Point(X, Y))) {
    // Drag the current time of day.
    DropTextSource1->Text = DateTimeToStr(Now());
    DropTextSource1->Execute();
  }
}
//---------------------------------------------------------------------------
void __fastcall TFormAutoScroll::DropTextTarget1DragOver(TObject *Sender,
      TShiftState ShiftState, TPoint &APoint, int &Effect)
{
   // Determine if the cursor if over a data cell. If it isn't, we do not accept
   // a drop.
   int CellX, CellY;

   StringGrid1->MouseToCell(APoint.x, APoint.y, CellX, CellY);
   if ((CellX < 1) || (CellY < 1))
     Effect = DROPEFFECT_NONE;
}
//---------------------------------------------------------------------------
void __fastcall TFormAutoScroll::DropTextTarget1Drop(TObject *Sender,
      TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  TDropTextTarget* tdtt = dynamic_cast<TDropTextTarget*>(Sender);
  if (!tdtt)
    return;
  // Determine which cell we dropped on and fill it with the dragged text.
  int CellX, CellY;
  StringGrid1->MouseToCell(APoint.x, APoint.y, CellX, CellY);
  StringGrid1->Cells[CellX][CellY] = tdtt->Text;
}
//---------------------------------------------------------------------------
void __fastcall TFormAutoScroll::DropTextTarget1Enter(TObject *Sender,
      TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  /*
  ** Set up a custom no-scroll zone:
  ** 1. Get the grids client rect.
  ** 2. Move the top left corner to the first data cell.
  ** 3. Shrink the rect 1/3 or the grid cell size.
  */
  TRect CustomNoScrollZone = StringGrid1->ClientRect;
  TRect FirstCell = StringGrid1->CellRect(StringGrid1->LeftCol,StringGrid1->TopRow);
  CustomNoScrollZone.Top = FirstCell.Top;
  CustomNoScrollZone.Left = FirstCell.Left;
  InflateRect(&CustomNoScrollZone, -StringGrid1->DefaultColWidth / 3,
    -StringGrid1->DefaultRowHeight / 3);

  TCustomDropTarget* tcdt = dynamic_cast<TCustomDropTarget*>(Sender);
  if (tcdt)
    tcdt->NoScrollZone = CustomNoScrollZone;
}
//---------------------------------------------------------------------------
