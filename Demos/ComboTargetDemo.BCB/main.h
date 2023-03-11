//---------------------------------------------------------------------------

#ifndef mainH
#define mainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "DropComboTarget.hpp"
#include "DropTarget.hpp"
#include <ComCtrls.hpp>
#include <ExtCtrls.hpp>
#include "DragDrop.hpp"
//---------------------------------------------------------------------------
class TFormMain : public TForm
{
__published:	// IDE-managed Components
    TPanel *Panel2;
    TPanel *PanelDropZone;
    TLabel *Label3;
    TPanel *Panel1;
    TGroupBox *GroupBox1;
    TCheckBox *CheckBoxText;
    TCheckBox *CheckBoxFiles;
    TCheckBox *CheckBoxURLs;
    TCheckBox *CheckBoxBitmaps;
    TCheckBox *CheckBoxMetaFiles;
    TCheckBox *CheckBoxData;
    TPageControl *PageControl1;
    TTabSheet *TabSheetText;
    TMemo *MemoText;
    TTabSheet *TabSheetFiles;
    TSplitter *Splitter1;
    TListBox *ListBoxFiles;
    TListBox *ListBoxMaps;
    TTabSheet *TabSheetBitmap;
    TScrollBox *ScrollBox2;
    TImage *ImageBitmap;
    TTabSheet *TabSheetURL;
    TLabel *Label1;
    TLabel *Label2;
    TEdit *EditURLURL;
    TEdit *EditURLTitle;
    TTabSheet *TabSheetData;
    TListView *ListViewData;
    TTabSheet *TabSheetMetaFile;
    TScrollBox *ScrollBox1;
    TImage *ImageMetaFile;
    TDropComboTarget *DropComboTarget1;
    void __fastcall DropComboTarget1Drop(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
    void __fastcall ListViewDataDblClick(TObject *Sender);
    void __fastcall CheckBoxTextClick(TObject *Sender);
    void __fastcall CheckBoxFilesClick(TObject *Sender);
    void __fastcall CheckBoxURLsClick(TObject *Sender);
    void __fastcall CheckBoxBitmapsClick(TObject *Sender);
    void __fastcall CheckBoxMetaFilesClick(TObject *Sender);
    void __fastcall CheckBoxDataClick(TObject *Sender);
private:	// User declarations
public:		// User declarations
    __fastcall TFormMain(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFormMain *FormMain;
//---------------------------------------------------------------------------
#endif
