//---------------------------------------------------------------------------

#ifndef TargetH
#define TargetH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "DragDrop.hpp"
#include "DragDropText.hpp"
#include "DropTarget.hpp"
#include <ExtCtrls.hpp>
//---------------------------------------------------------------------------
class TFormTarget : public TForm
{
__published:	// IDE-managed Components
    TPanel *Panel2;
    TPanel *PanelDest;
    TPanel *Panel5;
    TDropTextTarget *DropTextTarget1;
    void __fastcall DropTextTarget1Drop(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
private:	// User declarations
    TGenericDataFormat* TimeDataFormatTarget;
public:		// User declarations
    __fastcall TFormTarget(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFormTarget *FormTarget;
//---------------------------------------------------------------------------
#endif
