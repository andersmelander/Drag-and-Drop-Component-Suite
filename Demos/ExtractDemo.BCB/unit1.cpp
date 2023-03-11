//---------------------------------------------------------------------------
// This demo application was contributed by Jonathan Arnold.
//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include <fstream>
#include <FileCtrl.hpp>
#include "Unit1.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DragDrop"
#pragma link "DragDropFile"
#pragma link "DropSource"
#pragma resource "*.dfm"
TFormMain *FormMain;

//---------------------------------------------------------------------------
// TODO : IncludeTrailingBackslash needs implentation for C++Builder 4.
/*
AnsiString IncludeTrailingBackslash(const AnsiString Path)
{
  return Path;
}
*/

//---------------------------------------------------------------------------
__fastcall TFormMain::TFormMain(TComponent* Owner): TForm(Owner)
{
  // Get path to temporary directory
  char tPath[_MAX_PATH];
  ::GetTempPath(sizeof(tPath), tPath);
  TempPath = tPath;
  IncludeTrailingBackslash(TempPath);

  // List of all extracted files
  ExtractedFiles = new TStringList();
}
//---------------------------------------------------------------------------
__fastcall TFormMain::~TFormMain()
{
  // Before we exit, we make sure that we aren't leaving any extracted
  // files behind. Since it is the drop target's responsibility to
  // clean up after an optimized drag/move operation, we might get away with
  // just deleting all drag/copied files, but since many ill behaved drop
  // targets doesn't clean up after them selves, we will do it for them
  // here. If you trust your drop target to clean up after itself, you can skip
  // this step.
  // Note that this means that you shouldn't exit this application before
  // the drop target has had a chance of actually copy/move the files.
  for (int i = 0; i < ExtractedFiles->Count; ++i) {
    if (FileExists(ExtractedFiles->Strings[i])) {
      try {
	DeleteFile(ExtractedFiles->Strings[i]);
      }
      catch( ... ) {
        // Ignore any errors we might get
      }
    }
  }

  // Note: We should also remove any folders we created, but this example
  // doesn't do that.

  delete ExtractedFiles;
}
//---------------------------------------------------------------------------

////////////////////////////////////////////////////////////////////////////////
//
//		MouseDown handler.
//
////////////////////////////////////////////////////////////////////////////////
// Does drag detection, sets up the filename list and starts the drag operation.
////////////////////////////////////////////////////////////////////////////////
void __fastcall TFormMain::ListView1MouseDown(TObject *Sender,
  TMouseButton Button, TShiftState Shift, int X, int Y)
{
  // If no files selected then exit...
  if (ListView1->SelCount == 0)
    return;

  if (DragDetectPlus(reinterpret_cast<THandle>(((TWinControl*)Sender)->Handle),
    Point(X,Y))) {

    // Clear any filenames left from a previous drag operation...
    DropFileSource1->Files->Clear();

    // "Extracting" files here would be much simpler but is often
    // very inefficient as many drag ops are cancelled before the
    // files are actually dropped. Instead we delay the extracting,
    // until we know the user really wants the files, but load the
    // filenames into DropFileSource1->Files as if they already exist...

    // Add root files and top level subfolders...
    for (int i = 0; i < ListView1->Items->Count; ++i) {
      if (ListView1->Items->Item[i]->Selected) {
        // Note that it isn't nescessary to list files and folders in
	// sub folders. It is sufficient to list the top level sub folders,
	// since the drag target will copy/move the sub folders and
        // everything they contain.
	// Some target applications might not be able to handle this
	// optimization or it might not suit your purposes. In that case,
	// simply remove all the code between [A] and [B] below .

	// [A]
	int j = ListView1->Items->Item[i]->Caption.Pos("\\");
	if (j > 0) {
	  // Item is a subfolder...

          // Get the top level subfolder.
	  AnsiString s = ListView1->Items->Item[i]->Caption.SubString(1, j-1);

	  // Add folder if it hasn"t already been done.
          if (DropFileSource1->Files->IndexOf(TempPath + s) == -1)
            DropFileSource1->Files->Add(TempPath + s);
        }
        else
          // [B]
	  // Item is a file in the root folder...
	  DropFileSource1->Files->Add(TempPath + ListView1->Items->Item[i]->Caption);
      }
    }

    // Start the drag operation...
    DropFileSource1->Execute();
  }
}

//---------------------------------------------------------------------------
void __fastcall TFormMain::ButtonCloseClick(TObject *Sender)
{
  Close();
}
//---------------------------------------------------------------------------
////////////////////////////////////////////////////////////////////////////////
//
//		OnAfterDrop handler.
//
////////////////////////////////////////////////////////////////////////////////
// Executes after the target has returned from its OnDrop event handler.
////////////////////////////////////////////////////////////////////////////////
void __fastcall TFormMain::DropFileSource1AfterDrop(TObject *Sender,
  TDragResult DragResult, bool Optimized)
{
  // If the user performed a move operation, we now delete the selected files
  // from the archive.
  if (DragResult == drDropMove) {
    for (int i = ListView1->Items->Count-1; i >= 0; --i) {
      if (ListView1->Items->Item[i]->Selected)
        RemoveFile(i);
    }
  }

  // If the user performed an unoptimized move operation, we must delete the
  // files that were extracted.
  if ((DragResult == drDropMove) && !Optimized) {
    for (int i = 0; i < DropFileSource1->Files->Count; ++i) {
      if (FileExists(DropFileSource1->Files->Strings[i])) {
        try {
          DeleteFile(DropFileSource1->Files->Strings[i]);
            // Remove the files we just deleted from the "to do" list.
            int j = ExtractedFiles->IndexOf(DropFileSource1->Files->Strings[i]);
            if (j != -1)
              ExtractedFiles->Delete(j);
        }
        catch(...) {
          // Ignore any errors we might get.
        }
      }
    }
  }
  // Note: We should also remove any folders we created, but this example
  // doesn"t do that.
}
//---------------------------------------------------------------------------
////////////////////////////////////////////////////////////////////////////////
//
//		OnDrop handler.
//
////////////////////////////////////////////////////////////////////////////////
// Executes when the user drops the files on a drop target.
////////////////////////////////////////////////////////////////////////////////
void __fastcall TFormMain::DropFileSource1Drop(TObject *Sender,
  TDragType DragType, bool &ContinueDrop)
{
  // If the user actually dropped the filenames somewhere, we would now
  // have to extract the files from the archive. The files should be
  // extracted to the same path and filename as the ones we specified
  // in the drag operation. Otherwise the drop source will not be able
  // to find the files.

  // "Extract" all the selected files into the temporary folder tree...
  for (int i = ListView1->Items->Count-1; i >= 0; --i) {
    if (ListView1->Items->Item[i]->Selected)
      ExtractFile(i, TempPath + ListView1->Items->Item[i]->Caption);
  }

  // As soon as this method returns, the destination"s (e->g. Explorer"s)
  // DropTarget->OnDrop event will trigger and the destination will
  // start copying/moving the files.
}
//---------------------------------------------------------------------------
////////////////////////////////////////////////////////////////////////////////
//
//		Extract file from archive.
//
////////////////////////////////////////////////////////////////////////////////
// This method extracts a single file from the archive and saves it to disk.
// In a "real world" application, you would create (e->g. unzip, download etc.)
// your physical files here.
////////////////////////////////////////////////////////////////////////////////
void  TFormMain::ExtractFile(int FileIndex, AnsiString Filename)
{
  // Of course, this is a demo so we"ll just make phoney files here...
  AnsiString path = ExtractFilePath(Filename);
  if (path != "")
    ForceDirectories(path); // Create all the folders

  // Now, create an empty file
  std::fstream fs(Filename.c_str(), std::ios_base::out|std::ios_base::trunc);

  // Remember that we have extracted this file
  if (ExtractedFiles->IndexOf(Filename) == -1)
    ExtractedFiles->Add(Filename);
}

////////////////////////////////////////////////////////////////////////////////
//
//		Delete file from archive.
//
////////////////////////////////////////////////////////////////////////////////
// This method removes a single file from the archive.
////////////////////////////////////////////////////////////////////////////////
void TFormMain::RemoveFile(int FileIndex)
{
  // This is just a demo, so we"ll just remove the filename from the ListView...
  ListView1->Items->Delete(FileIndex);
}
