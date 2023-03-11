//---------------------------------------------------------------------------

#ifndef mainH
#define mainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "DragDropInternet.hpp"
#include "DropTarget.hpp"
#include <ComCtrls.hpp>
#include <ExtCtrls.hpp>
#include "DragDrop.hpp"
//---------------------------------------------------------------------------
class TFormMain : public TForm
{
__published:	// IDE-managed Components
    TPanel *Panel1;
    TLabel *Label1;
    TLabel *Label2;
    TEdit *Edit1;
    TPanel *Panel2;
    TPageControl *PageControl1;
    TTabSheet *TabSheetHeader;
    TListView *ListViewHeader;
    TTabSheet *TabSheetContents;
    TMemo *MemoContent;
    TTabSheet *TabSheetRaw;
    TMemo *MemoRaw;
    TPanel *Panel3;
    TButton *Button1;
    TDropURLTarget *DropURLTarget1;
    void __fastcall DropURLTarget1Drop(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
    void __fastcall DropURLTarget1Enter(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
    void __fastcall Button1Click(TObject *Sender);
private:	// User declarations
	void ParseRFC822(const AnsiString Msg);
public:		// User declarations
    __fastcall TFormMain(TComponent* Owner);
};

// A custom Clipboard format
class TNetscapeMessageClipboardFormat : public TCustomTextClipboardFormat
{
private:
	TClipFormat CF_NETSCAPEMESSAGE;
	
public:
	__fastcall TNetscapeMessageClipboardFormat() : TCustomTextClipboardFormat(),
										CF_NETSCAPEMESSAGE(0)
	{}
	
	__fastcall TClipFormat GetClipboardFormat()
	{
		if ( CF_NETSCAPEMESSAGE == 0 )
			CF_NETSCAPEMESSAGE = RegisterClipboardFormat("Netscape Message");
		return CF_NETSCAPEMESSAGE;
	}
};

//---------------------------------------------------------------------------
extern PACKAGE TFormMain *FormMain;
//---------------------------------------------------------------------------
#endif
