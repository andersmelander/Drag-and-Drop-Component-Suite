//---------------------------------------------------------------------------
// This demo application was contributed by Jonathan Arnold.
//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "DropFile.h"
#include <comobj.hpp> // For OleCheck()
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "cdiroutl"
#pragma link "DragDrop"
#pragma link "DragDropFile"
#pragma link "DropSource"
#pragma link "DropTarget"
#pragma link "DropTarget"
#pragma resource "*.dfm"
TFormFile *FormFile;

void CreateLink(AnsiString SourceFile, AnsiString ShortCutName)
{
  // CreateLink code from Harold Howe's http://www.bcbdev.com

  // IShellLink allows us to create the shortcut.
  // IPersistFile saves the link to the hard drive.
  IShellLink* pLink;
  IPersistFile* pPersistFile;

  // First, we have to initialize the COM library.
  // Note: This is actually done for us by the drag/drop library, but for
  // completeness it is also included here. No harm done.
  if (SUCCEEDED(CoInitialize(NULL))) {
    try {

      // If CoInitialize doesn't fail, then instantiate an
      // IShellLink object by calling CoCreateInstance.
      if(SUCCEEDED(CoCreateInstance(CLSID_ShellLink, NULL,
        CLSCTX_INPROC_SERVER, IID_IShellLink, (void **) &pLink))) {

        // if that succeeds, then fill in the shortcut attributes
        pLink->SetPath(SourceFile.c_str());
        pLink->SetWorkingDirectory( ExtractFilePath(SourceFile).c_str() );
        pLink->SetShowCmd(SW_SHOW);

        // Make sure we have a unique name
        AnsiString tmpShortCutName = ChangeFileExt(ShortCutName, ".lnk");
        int i=0;
        while (FileExists(tmpShortCutName )) { // not unique, append number until it is
          tmpShortCutName = ChangeFileExt( ShortCutName, "" ) + "(" +
            AnsiString(++i) + ").lnk";
        }

        // Now we need to save the shortcut to the hard drive. The
        // IShellLink object also implements the IPersistFile interface.
        // Get the IPersistFile part of the object using QueryInterface.
        if(SUCCEEDED(pLink->QueryInterface(IID_IPersistFile,
          (void **)&pPersistFile))) {

          // If that succeeds, then call the Save method of the
          // IPersistFile object to write the shortcut to the desktop.
          WideString strShortCutLocation(tmpShortCutName);
          pPersistFile->Save(strShortCutLocation.c_bstr(), TRUE);
          pPersistFile->Release();
        }
        pLink->Release();
      }
    }
    __finally {
      // Calls to CoInitialize need a corresponding CoUninitialize call
      CoUninitialize();
    }
  }
}

// TODO : IncludeTrailingBackslash needs implentation for C++Builder 4.
/*
AnsiString IncludeTrailingBackslash(const AnsiString Path)
{
  return Path;
}
*/

//---------------------------------------------------------------------------
__fastcall TFormFile::TFormFile(TComponent* Owner): TForm(Owner)
{
  // Load custom cursors...
  Screen->Cursors[crCopy] = ::LoadCursor(HInstance, "CUR_DRAG_COPY");
  Screen->Cursors[crMove] = ::LoadCursor(HInstance, "CUR_DRAG_MOVE");
  Screen->Cursors[crLink] = ::LoadCursor(HInstance, "CUR_DRAG_LINK");
  Screen->Cursors[crCopyScroll] = ::LoadCursor(HInstance, "CUR_DRAG_COPY_SCROLL");
  Screen->Cursors[crMoveScroll] = ::LoadCursor(HInstance, "CUR_DRAG_MOVE_SCROLL");
  Screen->Cursors[crLinkScroll] = ::LoadCursor(HInstance, "CUR_DRAG_LINK_SCROLL");
}
//---------------------------------------------------------------------------
__fastcall TFormFile::~TFormFile()
{
  if (DirectoryThread) {
    DirectoryThread->Terminate();
    DirectoryThread->WakeUp();
    DirectoryThread->WaitFor();
    delete DirectoryThread;
  }
}
//---------------------------------------------------------------------------

void __fastcall TFormFile::ListView1CustomDrawItem(TCustomListView *Sender,
  TListItem *Item, TCustomDrawState State, bool &DefaultDraw)
{
  // Items which have been "cut to clipboard" are drawn differently.
  if (Item->Data)
    Sender->Canvas->Font->Style = TFontStyles() << fsStrikeOut;
}
//---------------------------------------------------------------------------

void __fastcall TFormFile::ListView1MouseDown(TObject *Sender,
  TMouseButton Button, TShiftState Shift, int X, int Y)
{
  // If no files selected then exit...
  if (ListView1->SelCount == 0)
    return;

  // Wait for user to move cursor before we start the drag/drop.
  if (DragDetectPlus(reinterpret_cast<THandle>(Handle), Point(X,Y))) {

    StatusBar1->SimpleText = "";
    DropSource1->Files->Clear();

    // Fill DropSource1->Files with selected files in ListView1
    for ( int i = 0; i < ListView1->Items->Count; ++i ) {
      if (ListView1->Items->Item[i]->Selected) {
        AnsiString Filename = IncludeTrailingBackslash(DirectoryOutline->Directory) +
          ListView1->Items->Item[i]->Caption;
        DropSource1->Files->Add(Filename);

        // The TDropFileSource->MappedNames list can be used to indicate to the
        // drop target that the files should be renamed once they have been
        // copied. This is the technique used when dragging files from the
        // recycle bin.
        // DropSource1->MappedNames->Add("NewFileName"+inttostr(i+1));
      }
    }

    // Temporarily disable the list view as a drop target.
    DropFileTarget1->DragTypes = TDragTypes();

    TDragResult Res;

    try {
      // OK, now we are all set to go. Let"s start the drag...
      Res = DropSource1->Execute();
    }
    __finally {
      // Enable the list view as a drop target again.
      DropFileTarget1->DragTypes = TDragTypes() << dtCopy << dtMove << dtLink;
    }

    // Note:
    // The target is responsible, from this point on, for the
    // copying/moving/linking of the file but the target reports
    // back to the source what (should have) happened via the
    // returned value of Execute.

    // Feedback in StatusBar1 what happened...
    switch (Res) {
      case drDropCopy:
        StatusBar1->SimpleText = "Copied successfully";
        break;
      case drDropMove:
        StatusBar1->SimpleText = "Moved successfully";
        break;
      case drDropLink:
        StatusBar1->SimpleText = "Linked successfully";
        break;
      case drCancel:
        StatusBar1->SimpleText = "Drop was cancelled";
        break;
      case drOutMemory:
        StatusBar1->SimpleText = "Drop cancelled - out of memory";
        break;
      default:
        StatusBar1->SimpleText = "Drop cancelled - unknown reason";
        break;
    }
  }
}
//---------------------------------------------------------------------------
void __fastcall TFormFile::btnCloseClick(TObject *Sender)
{
  Close();
}
//---------------------------------------------------------------------------
void __fastcall TFormFile::DriveComboBoxChange(TObject *Sender)
{
  // Manual synchronization to work around bug in TDirectoryOutline.
  DirectoryOutline->Drive = DriveComboBox->Drive;
}
//---------------------------------------------------------------------------
void __fastcall TFormFile::PopupMenu1Popup(TObject *Sender)
{
  MenuCopy->Enabled = (ListView1->SelCount > 0);
  MenuCut->Enabled = MenuCopy->Enabled;

  // Open the clipboard as an IDataObject
  IDataObject* DataObject;
  OleCheck(OleGetClipboard(&DataObject));
  try {
    // Enable paste menu if the clipboard contains data in any of
    // the supported formats
    MenuPaste->Enabled = DropFileTarget1->HasValidFormats(DataObject);
  }
  __finally {
    DataObject = NULL;
  }
}
//---------------------------------------------------------------------------

void __fastcall TFormFile::DirectoryOutlineChange(TObject *Sender)
{
  if (!DirectoryThread)
    DirectoryThread = new TDirectoryThread(ListView1, DirectoryOutline->Directory);
  else
    DirectoryThread->Directory = DirectoryOutline->Directory;
}
//---------------------------------------------------------------------------

void __fastcall TFormFile::MenuCutOrCopyClick(TObject *Sender)
{
  if (ListView1->SelCount == 0) {
    StatusBar1->SimpleText = "No files have been selected!";
    return;
  }

  DropSource1->Files->Clear();
  for (int i = 0; i < ListView1->Items->Count; ++i) {
    if (ListView1->Items->Item[i]->Selected) {
      AnsiString Filename =
        IncludeTrailingBackslash(DirectoryOutline->Directory) +
          ListView1->Items->Item[i]->Caption;
      DropSource1->Files->Add(Filename);

      // Flag item as "cut" so it can be drawn differently.
      if (Sender == MenuCut)
        ListView1->Items->Item[i]->Data = (void*)true;
      else
        ListView1->Items->Item[i]->Data = (void*)false;
    }
    else
      ListView1->Items->Item[i]->Data = 0;
  }

  ListView1->Invalidate();

  // Transfer data to clipboard.
  AnsiString Operation;
  bool Status;

  if (Sender == MenuCopy) {
    Status = DropSource1->CopyToClipboard();
    Operation = "copied";
  }
  else if (Sender == MenuCut) {
    Status = DropSource1->CutToClipboard();
    Operation = "cut";
  }
  else
    Status = false;

  if (Status) {
    StatusBar1->SimpleText = Format("%d file(s) %s to clipboard.",
      ARRAYOFCONST((DropSource1->Files->Count, Operation)));
  }
}
//---------------------------------------------------------------------------

void __fastcall TFormFile::MenuPasteClick(TObject *Sender)
{
  // PasteFromClipboard fires an OnDrop event, so we don't need to do
  // anything special here.
  DropFileTarget1->PasteFromClipboard();
}
//---------------------------------------------------------------------------

//--------------------------
// SOURCE events...
//--------------------------
void __fastcall TFormFile::DropSource1AfterDrop(TObject *Sender,
  TDragResult DragResult, bool Optimized)
{
  // Delete source files if target performed an unoptimized drag/move
  // operation (target copies files, source deletes them).
  if ((DragResult == drDropMove) &&  !Optimized)
    for (int i = 0; i < DropSource1->Files->Count; ++i)
      DeleteFile(DropSource1->Files->Strings[i]);
}
//---------------------------------------------------------------------------
void __fastcall TFormFile::DropSource1Feedback(TObject *Sender, int Effect,
  bool &UseDefaultCursors)
{
  UseDefaultCursors = false; // We want to use our own.
  switch (Effect) {
    case DROPEFFECT_COPY:
      ::SetCursor(Screen->Cursors[crCopy]);
      break;
    case DROPEFFECT_MOVE:
      ::SetCursor(Screen->Cursors[crMove]);
      break;
    case DROPEFFECT_LINK:
      ::SetCursor(Screen->Cursors[crLink]);
      break;
    case DROPEFFECT_SCROLL | DROPEFFECT_COPY:
      ::SetCursor(Screen->Cursors[crCopyScroll]);
      break;
    case DROPEFFECT_SCROLL | DROPEFFECT_MOVE:
      ::SetCursor(Screen->Cursors[crMoveScroll]);
      break;
    case DROPEFFECT_SCROLL | DROPEFFECT_LINK:
      ::SetCursor(Screen->Cursors[crLinkScroll]);
      break;
    default:
      UseDefaultCursors = true; // Use default NoDrop
      break;
  }
}
//---------------------------------------------------------------------------

void __fastcall TFormFile::DropSource1Paste(TObject *Sender,
  TDragResult Action, bool DeleteOnPaste)
{
  StatusBar1->SimpleText = "Target pasted file(s)";

  // Delete source files if target performed a paste/move operation and
  // requested the source to "Delete on paste".
  if (DeleteOnPaste)
    for (int i = 0; i < DropSource1->Files->Count; ++i)
      DeleteFile(DropSource1->Files->Strings[i]);
}
//---------------------------------------------------------------------------

//--------------------------
// TARGET events...
//--------------------------

void __fastcall TFormFile::DropFileTarget1Enter(TObject *Sender,
  TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  // Note: GetDataOnEnter has been set to True.
  // A side effect of this is that TDropFileTarget can't be used to accept drops
  // from WinZip. Although the file names are received correctly, the files
  // aren't extracted and thus can't be copied/moved.
  // This is caused by a quirk in WinZip; Apparently WinZip doesn't like
  // IDataObject.GetData to be called before IDropTarget.Drop is called.

  // Save the location (path) of the files being dragged.
  // Also flags if an EXE file is being dragged.
  // This info will be used to set the default (ie. no Shift or Ctrl Keys
  // pressed) drag behaviour (COPY, MOVE or LINK).
  if (DropFileTarget1->Files->Count > 0) {
    SourcePath = ExtractFilePath(DropFileTarget1->Files->Strings[0]);
    IsEXEfile = (DropFileTarget1->Files->Count == 1) &&
      (AnsiCompareText(ExtractFileExt(DropFileTarget1->Files->Strings[0]),
        ".exe") == 0);
  }
}
//---------------------------------------------------------------------------

void __fastcall TFormFile::DropFileTarget1Drop(TObject *Sender,
  TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  int SuccessCnt = 0;
  AnsiString NewPath = IncludeTrailingBackslash(DirectoryOutline->Directory);

  // Filter out the DROPEFFECT_SCROLL flag if set...
  // (ie: when dropping a file while the target window is scrolling)
  Effect = Effect & ~DROPEFFECT_SCROLL;

  // Now, 'Effect' should equal one of the following:
  // DROPEFFECT_COPY, DROPEFFECT_MOVE or DROPEFFECT_LINK.
  // Note however, that if we call TDropTarget->PasteFromClipboard, Effect
  // can be a combination of the above drop effects.

  for (int i = 0; i < DropFileTarget1->Files->Count; ++i) {
    // Name mapping occurs when dragging files from Recycled Bin...
    // In most situations Name Mapping can be ignored entirely.
    AnsiString NewFilename;
    if (i < DropFileTarget1->MappedNames->Count)
      NewFilename = NewPath+DropFileTarget1->MappedNames->Strings[i];
    else
      NewFilename = NewPath+ExtractFileName(DropFileTarget1->Files->Strings[i]);

    if (!FileExists(NewFilename)) {
      try {
        if (Effect & DROPEFFECT_COPY) {
          Effect = DROPEFFECT_COPY;
          // Copy the file.
          if (::CopyFile(DropFileTarget1->Files->Strings[i].c_str(),
            NewFilename.c_str(), true))
          SuccessCnt++;
        }
        else if (Effect & DROPEFFECT_MOVE) {
          Effect = DROPEFFECT_MOVE;
          // Move the file.
          if (RenameFile(DropFileTarget1->Files->Strings[i], NewFilename))
            SuccessCnt++;
        }
      }
      catch ( ... ) {
        // Ignore errors.
      }
    }

    if (Effect & DROPEFFECT_LINK) {
      Effect = DROPEFFECT_LINK;
      // Create a shell link to the file.
      CreateLink(DropFileTarget1->Files->Strings[i], NewFilename);
      SuccessCnt++;
    }
  }

  switch (Effect) {
    case DROPEFFECT_MOVE:
      StatusBar1->SimpleText =
        Format("%d file(s) were moved.   Files dropped at point (%d,%d).",
          ARRAYOFCONST((SuccessCnt, (int)APoint.x, (int)APoint.y)));
      break;
    case DROPEFFECT_COPY:
      StatusBar1->SimpleText =
        Format("%d file(s) were copied.   Files dropped at point (%d,%d).",
          ARRAYOFCONST((SuccessCnt, (int)APoint.x, (int)APoint.y)));
      break;
    case DROPEFFECT_LINK:
      StatusBar1->SimpleText =
        Format("%d file(s) were linked.   Files dropped at point (%d,%d).",
          ARRAYOFCONST((SuccessCnt, (int)APoint.x, (int)APoint.y)));
      break;
  }
}
//---------------------------------------------------------------------------

void __fastcall TFormFile::DropFileTarget1GetDropEffect(TObject *Sender,
  TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  // Note: The 'Effect' parameter (on event entry) is the
  // set of effects allowed by both the source and target.
  // Use this event when you wish to override the Default behaviour...

  // Save the value of the auto scroll flag.
  // As an alternative we could implement our own auto scroll logic here.
  int Scroll = Effect & DROPEFFECT_SCROLL;

  // Reject the drop if source and target paths are the same (DROPEFFECT_NONE).
  if (IncludeTrailingBackslash(DirectoryOutline->Directory) == SourcePath)
    Effect = DROPEFFECT_NONE;
  // else if Ctrl+Shift are pressed then create a link (DROPEFFECT_LINK).
  else if (ShiftState.Contains(ssShift) && ShiftState.Contains(ssCtrl) &&
    (Effect & DROPEFFECT_LINK))
    Effect = DROPEFFECT_LINK;
  // else if Shift is pressed then move (DROPEFFECT_MOVE).
  else if (ShiftState.Contains(ssShift) && (Effect & DROPEFFECT_MOVE) )
    Effect = DROPEFFECT_MOVE;
  // else if Ctrl is pressed then copy (DROPEFFECT_COPY).
  else if (ShiftState.Contains(ssCtrl) && (Effect & DROPEFFECT_COPY) )
    Effect = DROPEFFECT_COPY;
    // else if dragging a single EXE file then default to link (DROPEFFECT_LINK).
  else if (IsEXEfile && (Effect & DROPEFFECT_LINK) )
    Effect = DROPEFFECT_LINK;
  // else if source and target drives are the same then default to MOVE (DROPEFFECT_MOVE).
  else if (!SourcePath.IsEmpty() && (DirectoryOutline->Directory[1] == SourcePath[1]) &&
    (Effect & DROPEFFECT_MOVE) )
    Effect = DROPEFFECT_MOVE;
  // otherwise just use whatever we can get away with.
  else if (Effect & DROPEFFECT_COPY)
    Effect = DROPEFFECT_COPY;
  else if (Effect & DROPEFFECT_MOVE)
    Effect = DROPEFFECT_MOVE;
  else if (Effect & DROPEFFECT_LINK)
    Effect = DROPEFFECT_LINK;
  else
    Effect = DROPEFFECT_NONE;

  // Restore auto scroll flag.
  Effect = Effect | Scroll;
}
//---------------------------------------------------------------------------

//----------------------------------------------------------------------------
// TDirectoryThread
// This thread monitors the current directory for changes and updates the
// ListView whenever the directory is changed (files added, renamed or deleted).
//----------------------------------------------------------------------------

// OK, we're showing off... This is a little overkill for a demo,
// but still you can see what can be done.
TDirectoryThread::TDirectoryThread(TListView* ListView, AnsiString Dir): TThread(true)
{
  fListView = ListView;
  Priority = tpLowest;
  fDirectory = Dir;
  FWakeupEvent = new TEvent(0, false, false, 0);
  FFiles = new TStringList();

  Resume();
}

__fastcall TDirectoryThread::~TDirectoryThread()
{
  delete FWakeupEvent;
  delete FFiles;
}

void TDirectoryThread::WakeUp()
{
  FWakeupEvent->SetEvent();
}

void TDirectoryThread::SetDirectory(AnsiString Value)
{
  if (Value == fDirectory)
    return;

  fDirectory = Value;
  WakeUp();
}

void TDirectoryThread::ScanDirectory()
{
  TSearchRec sr;

  FFiles->Clear();
  int res = FindFirst(IncludeTrailingBackslash(fDirectory)+"*.*", 0, sr);
  try {
    while ((res == 0) && !Terminated) {
      if ((sr.Name != ".") && (sr.Name != ".."))
        FFiles->Add(sr.Name.LowerCase());
      res = FindNext(sr);
    }
  }
  __finally {
    FindClose(&sr);
  }
}

void __fastcall TDirectoryThread::UpdateListView()
{
  fListView->Items->BeginUpdate();
  try {
    fListView->Items->Clear();
    for (int i = 0; i < FFiles->Count; ++i) {
      TListItem* NewItem = fListView->Items->Add();
      NewItem->Caption = FFiles->Strings[i];
    }
    if (fListView->Items->Count > 0)
      fListView->ItemFocused = fListView->Items->Item[0];
  }
  __finally {
    fListView->Items->EndUpdate();
  }
  FFiles->Clear();
}

void __fastcall TDirectoryThread::Execute()
{
  // OUTER LOOP - which will exit only when terminated ...
  // directory changes will be processed within this OUTER loop
  // (file changes will be processed within the INNER loop)
  while (!Terminated) {
    ScanDirectory();
    Synchronize(UpdateListView);

    //Monitor directory for file changes
    HANDLE fFileChangeHandle =
    ::FindFirstChangeNotification(fDirectory.c_str(), false,
      FILE_NOTIFY_CHANGE_FILE_NAME);
    if (fFileChangeHandle == INVALID_HANDLE_VALUE)
      //Can't monitor filename changes! Just wait for change of directory or terminate
      ::WaitForSingleObject(FWakeupEvent, INFINITE);
    else
    {
      try { //This function performs an INNER loop...
        ProcessFilenameChanges(fFileChangeHandle);
      }
      __finally {
        FindCloseChangeNotification(fFileChangeHandle);
      }
    }
  }
}

void TDirectoryThread::ProcessFilenameChanges(HANDLE fcHandle)
{
  DWORD WaitResult;
  HANDLE HandleArray[] = { (HANDLE)FWakeupEvent->Handle, fcHandle };

  // INNER LOOP -
  // which will exit only if terminated or the directory is changed
  // filename changes will be processed within this loop
  while (!Terminated) {

    // wait for either filename or directory change, or terminate...
    WaitResult = ::WaitForMultipleObjects(2, HandleArray, false, INFINITE);

    if (WaitResult == WAIT_OBJECT_0 + 1) {
      //filename has changed
      do { //collect all immediate filename changes...
        FindNextChangeNotification(fcHandle);
      } while (!Terminated && (WaitForSingleObject(fcHandle, 0) == WAIT_OBJECT_0));
      if (Terminated)
        break;

      // OK, now update (before restarting inner loop)...
      ScanDirectory();
      Synchronize(UpdateListView);
    }
    else
    { // Either directory changed or terminated ...
      //collect all (almost) immediate directory changes before exiting...
      while ((!Terminated) &&
        (WaitForSingleObject(FWakeupEvent, 100) == WAIT_OBJECT_0))
        ;
      break;
    }
  }
}

