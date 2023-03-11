//---------------------------------------------------------------------------

#ifndef DropFileH
#define DropFileH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "cdiroutl.h"
#include "DragDrop.hpp"
#include "DragDropFile.hpp"
#include "DropSource.hpp"
#include <ComCtrls.hpp>
#include <ExtCtrls.hpp>
#include <FileCtrl.hpp>
#include <Grids.hpp>
#include <ImgList.hpp>
#include <Menus.hpp>
#include <Outline.hpp>
#include <SyncObjs.hpp>
#include "DropTarget.hpp"
//---------------------------------------------------------------------------

// This thread is used to watch for and
// display changes in DirectoryOutline.directory
class TDirectoryThread : public TThread
{
private:
    TListView* fListView;
    AnsiString fDirectory;
    TEvent* FWakeupEvent; //Used to signal change of directory or terminating
    TStrings* FFiles;
protected:
    void ScanDirectory();
    void __fastcall UpdateListView();
    void SetDirectory(AnsiString Value);
	void ProcessFilenameChanges(HANDLE fcHandle);
public:
    TDirectoryThread(TListView* ListView, AnsiString Dir);
    __fastcall ~TDirectoryThread();
    void __fastcall Execute(); //override
    void WakeUp();
    __property AnsiString Directory={read=fDirectory, write=SetDirectory};
};

class TFormFile : public TForm
{
__published:	// IDE-managed Components
    TDriveComboBox *DriveComboBox;
    TMemo *Memo1;
    TListView *ListView1;
    TButton *btnClose;
    TStatusBar *StatusBar1;
    TPanel *Panel1;
    TDropFileTarget *DropFileTarget1;
    TDropFileSource *DropSource1;
    TImageList *ImageList1;
    TDropDummy *DropDummy1;
    TPopupMenu *PopupMenu1;
    TMenuItem *MenuCopy;
    TMenuItem *MenuCut;
    TMenuItem *N1;
    TMenuItem *MenuPaste;
    TCDirectoryOutline *DirectoryOutline;
    void __fastcall ListView1CustomDrawItem(TCustomListView *Sender,
          TListItem *Item, TCustomDrawState State, bool &DefaultDraw);
    void __fastcall ListView1MouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
    void __fastcall btnCloseClick(TObject *Sender);
    void __fastcall DriveComboBoxChange(TObject *Sender);
    void __fastcall DropSource1AfterDrop(TObject *Sender,
          TDragResult DragResult, bool Optimized);
    void __fastcall DropSource1Feedback(TObject *Sender, int Effect,
          bool &UseDefaultCursors);
    void __fastcall DropSource1Paste(TObject *Sender, TDragResult Action,
          bool DeleteOnPaste);
    void __fastcall PopupMenu1Popup(TObject *Sender);
    void __fastcall DropFileTarget1Drop(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
    void __fastcall DropFileTarget1Enter(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
    void __fastcall DropFileTarget1GetDropEffect(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
    void __fastcall DirectoryOutlineChange(TObject *Sender);
    void __fastcall MenuCutOrCopyClick(TObject *Sender);
    void __fastcall MenuPasteClick(TObject *Sender);
private:	// User declarations

	// CUSTOM CURSORS:
	// The cursors in DropCursors.res are exactly the same as the default cursors.
	// Use DropCursors.res as a template if you wish to customise your own cursors.
	// For this demo we've created Cursors.res - some coloured cursors.
	enum {crCopy=101, crMove, crLink, crCopyScroll, crMoveScroll, crLinkScroll};
	
	AnsiString SourcePath;
	bool IsEXEfile;
	TDirectoryThread* DirectoryThread;
public:		// User declarations
    __fastcall TFormFile(TComponent* Owner);
    __fastcall ~TFormFile();
};
//---------------------------------------------------------------------------
extern PACKAGE TFormFile *FormFile;
//---------------------------------------------------------------------------
#endif
