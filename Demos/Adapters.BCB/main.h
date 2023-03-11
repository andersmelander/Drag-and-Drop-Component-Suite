//---------------------------------------------------------------------------

#ifndef mainH
#define mainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "DragDropText.hpp"
#include "DropTarget.hpp"

// Note: We have to include the appropriate header files for the File
// and URL data format.  The DragDropFile header contains the
// TFileDataFormat class and the DragDropInternet header contains the
// TURLDataFormat class.
#include "DragDropFormats.hpp"
#include "DragDropInternet.hpp"
#include "DragDropFile.hpp"

#include <ExtCtrls.hpp>
#include "DragDrop.hpp"
//---------------------------------------------------------------------------
class TFormMain : public TForm
{
__published:	// IDE-managed Components
    TPanel *Panel1;
    TLabel *Label1;
    TGroupBox *GroupBox1;
    TMemo *MemoText;
    TGroupBox *GroupBox3;
    TMemo *MemoFile;
    TGroupBox *GroupBox2;
    TMemo *MemoURL;
    TPanel *Panel2;
    TDropTextTarget *DropTextTarget1;
    TDataFormatAdapter *DataFormatAdapterFile;
    TDataFormatAdapter *DataFormatAdapterURL;
    void __fastcall DropTextTarget1Drop(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
private:	// User declarations
public:		// User declarations
    __fastcall TFormMain(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFormMain *FormMain;
//---------------------------------------------------------------------------
#endif
