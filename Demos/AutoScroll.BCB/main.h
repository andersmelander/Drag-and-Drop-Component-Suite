//---------------------------------------------------------------------------

#ifndef mainH
#define mainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "DragDrop.hpp"
#include "DragDropText.hpp"
#include "DropSource.hpp"
#include "DropTarget.hpp"
#include <ComCtrls.hpp>
#include <ExtCtrls.hpp>
#include <Grids.hpp>
//---------------------------------------------------------------------------
class TFormAutoScroll : public TForm
{
__published:	// IDE-managed Components
    TPanel *Panel3;
    TRichEdit *RichEdit1;
    TPanel *Panel1;
    TStringGrid *StringGrid1;
    TPanel *PanelSource;
    TDropTextSource *DropTextSource1;
    TDropTextTarget *DropTextTarget1;
    void __fastcall PanelSourceMouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
    void __fastcall DropTextTarget1DragOver(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
    void __fastcall DropTextTarget1Drop(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
    void __fastcall DropTextTarget1Enter(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
private:	// User declarations
public:		// User declarations
    __fastcall TFormAutoScroll(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFormAutoScroll *FormAutoScroll;
//---------------------------------------------------------------------------
#endif
