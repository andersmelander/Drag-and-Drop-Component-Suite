//---------------------------------------------------------------------------

#ifndef mainH
#define mainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "DropSource.hpp"
#include <ComCtrls.hpp>
#include <ExtCtrls.hpp>
#include "DragDrop.hpp"
#include "DragDropFormats.hpp"
//---------------------------------------------------------------------------
class TFormMain : public TForm
{
__published:	// IDE-managed Components
    TPanel *Panel1;
    TStatusBar *StatusBar1;
    TButton *ButtonAbort;
    TProgressBar *ProgressBar1;
    TPanel *Panel2;
    TPanel *Panel3;
    TPaintBox *PaintBoxPie;
    TPanel *Panel4;
    TLabel *Label2;
    TLabel *Label3;
    TRadioButton *RadioButtonNormal;
    TRadioButton *RadioButtonAsync;
    TPanel *Panel5;
    TLabel *Label1;
    TTimer *Timer1;
    TDropEmptySource *DropEmptySource1;
    TDataFormatAdapter *DataFormatAdapterSource;
        TLabel *Label4;
    void __fastcall Timer1Timer(TObject *Sender);
    void __fastcall DropEmptySource1AfterDrop(TObject *Sender,
          TDragResult DragResult, bool Optimized);
    void __fastcall DropEmptySource1Drop(TObject *Sender,
          TDragType DragType, bool &ContinueDrop);
    void __fastcall FormDestroy(TObject *Sender);
    void __fastcall ButtonAbortClick(TObject *Sender);
    void __fastcall OnMouseDown(TObject *Sender, TMouseButton Button,
          TShiftState Shift, int X, int Y);
private:	// User declarations
	void DrawPie( int Percent );
	int Tick;
	bool EvenOdd;
	bool DoAbort;
	void __fastcall OnGetStream(TFileContentsStreamOnDemandClipboardFormat* Sender,
								int Index,  _di_IStream &AStream);
	void __fastcall OnProgress( TObject* Sender, int Count, int MaxCount );
public:		// User declarations
    __fastcall TFormMain(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFormMain *FormMain;
//---------------------------------------------------------------------------

// TFakeStream is a read-only stream which produces its contents on-the-run.
// It is used for this demo so we can simulate transfer of very large and
// arbitrary amounts of data without using any memory.
typedef void __fastcall (__closure *TStreamProgressEvent)(System::TObject* Sender, int Count, int MaxCount );

class TFakeStream : public TStream
{
private:
	typedef TStream Inherited;
	
	int FSize, FPosition, FMaxCount;
	TStreamProgressEvent FProgress;
	bool FAbort;
public:
	TFakeStream(int _Size, int _MaxCount);
	int __fastcall Read(void* Buffer, int Count);
	int __fastcall Seek(int Offset, Word Origin);
#if defined(RTLVersion)
  #if (RTLVersion < 14)
        __int64 __fastcall Seek(const __int64 Offset, TSeekOrigin Origin);
  #endif
#endif"
	void __fastcall SetSize(int NewSize);
        void __fastcall SetSize(const __int64 NewSize);
	int __fastcall Write(const void* Buffer, int Count);
	void Abort() { FAbort=true; };
	__property TStreamProgressEvent OnProgress = {read=FProgress, write=FProgress};
};
	
#endif

