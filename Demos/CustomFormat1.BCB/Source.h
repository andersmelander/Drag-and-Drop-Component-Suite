//---------------------------------------------------------------------------

#ifndef SourceH
#define SourceH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "DragDropText.hpp"
#include "DropSource.hpp"
#include <ExtCtrls.hpp>
#include "DragDrop.hpp"
//---------------------------------------------------------------------------
class TFormSource : public TForm
{
__published:	// IDE-managed Components
    TPanel *Panel3;
    TMemo *Memo1;
    TPanel *Panel1;
    TPanel *PanelSource;
    TPanel *Panel4;
    TDropTextSource *DropTextSource1;
    TTimer *Timer1;
    void __fastcall PanelSourceMouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
    void __fastcall Timer1Timer(TObject *Sender);
private:	// User declarations
    TGenericDataFormat* TimeDataFormatSource;
public:		// User declarations
    __fastcall TFormSource(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFormSource *FormSource;
//---------------------------------------------------------------------------
#endif
