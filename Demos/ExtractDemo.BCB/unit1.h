//---------------------------------------------------------------------------

#ifndef Unit1H
#define Unit1H
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "DragDrop.hpp"
#include "DragDropFile.hpp"
#include "DropSource.hpp"
#include <ComCtrls.hpp>
#include <ExtCtrls.hpp>
//---------------------------------------------------------------------------
class TFormMain : public TForm
{
__published:	// IDE-managed Components
    TPanel *Panel1;
    TLabel *Label2;
    TButton *ButtonClose;
    TListView *ListView1;
    TDropFileSource *DropFileSource1;
    TStatusBar *StatusBar1;
    void __fastcall ListView1MouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
    void __fastcall ButtonCloseClick(TObject *Sender);
    void __fastcall DropFileSource1AfterDrop(TObject *Sender,
          TDragResult DragResult, bool Optimized);
    void __fastcall DropFileSource1Drop(TObject *Sender,
          TDragType DragType, bool &ContinueDrop);
private:	// User declarations
    AnsiString TempPath; // path to temp folder
    TStringList* ExtractedFiles;
    void ExtractFile(int FileIndex, AnsiString Filename);
    void RemoveFile(int FileIndex);
public:		// User declarations
    __fastcall TFormMain(TComponent* Owner);
	__fastcall ~TFormMain();
};
//---------------------------------------------------------------------------
extern PACKAGE TFormMain *FormMain;
//---------------------------------------------------------------------------
#endif
 
