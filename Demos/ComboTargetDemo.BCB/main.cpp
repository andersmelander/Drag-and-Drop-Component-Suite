//---------------------------------------------------------------------------
// This demo application was contributed by Jonathan Arnold.
//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "main.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DragDrop"
#pragma link "DropComboTarget"
#pragma link "DropTarget"
#pragma link "DragDrop"
#pragma resource "*.dfm"
TFormMain *FormMain;
//---------------------------------------------------------------------------
__fastcall TFormMain::TFormMain(TComponent* Owner): TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::DropComboTarget1Drop(TObject *Sender,
  TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  // Clear all formats.
  EditURLURL->Text = "";
  EditURLTitle->Text = "";
  MemoText->Lines->Clear();
  ImageBitmap->Picture->Assign(0);
  ImageMetaFile->Picture->Assign(0);
  ListBoxFiles->Items->Clear();
  ListBoxMaps->Items->Clear();
  ListViewData->Items->Clear();

  // Extract and display dropped data
  for (int i = 0; i <= DropComboTarget1->Data->Count-1; ++i) {
    Name = DropComboTarget1->Data->Names->Names[i];
    if (Name == "")
      Name = AnsiString(i)+"->dat";
    TStream* Stream = new TFileStream(ExtractFilePath(Application->ExeName)+Name, fmCreate);
    try {
      TListItem* listItem = ListViewData->Items->Add();
      listItem->Caption = Name;
      listItem->SubItems->Add(DropComboTarget1->Data->Streams[i]->Size);

      // Copy dropped data to stream (in this case a file stream)
      Stream->CopyFrom(DropComboTarget1->Data->Streams[i],
      DropComboTarget1->Data->Streams[i]->Size);
    }
    __finally {
      delete Stream;
    }
  }

  // Copy the rest of the dropped formats->
  ListBoxFiles->Items->Assign(DropComboTarget1->Files);
  ListBoxMaps->Items->Assign(DropComboTarget1->FileMaps);
  EditURLURL->Text = DropComboTarget1->URL;
  EditURLTitle->Text = DropComboTarget1->Title;
  ImageBitmap->Picture->Assign(DropComboTarget1->Bitmap);
  ImageMetaFile->Picture->Assign(DropComboTarget1->MetaFile);
  MemoText->Lines->Text = DropComboTarget1->Text;

  // Determine which formats were dropped->
  TabSheetFiles->TabVisible = (ListBoxFiles->Items->Count > 0);
  TabSheetURL->TabVisible = (EditURLURL->Text != "") || (EditURLTitle->Text != "");
  TabSheetBitmap->TabVisible = (ImageBitmap->Picture->Graphic) &&
    (!ImageBitmap->Picture->Graphic->Empty);
  TMetafile* tmf = dynamic_cast<TMetafile*>(ImageMetaFile->Picture->Graphic);
  TabSheetMetaFile->TabVisible = tmf && tmf->Handle;
  TabSheetText->TabVisible = (MemoText->Lines->Count > 0);
  TabSheetData->TabVisible = (ListViewData->Items->Count > 0);
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::ListViewDataDblClick(TObject *Sender)
{
  // Launch an extracted data file if user double clicks on it.
  Screen->Cursor = crAppStart;
  try {
    Application->ProcessMessages(); //otherwise cursor change will be missed
    ShellExecute(0, 0,
      (ExtractFilePath(Application->ExeName)+
        ((TListView*)(Sender))->Selected->Caption).c_str(),
      0, 0, SW_NORMAL);
  }
  __finally {
    Screen->Cursor = crDefault;
  }
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::CheckBoxTextClick(TObject *Sender)
{
  // Enable or disable format according to users selection
  TCheckBox* cbox = static_cast<TCheckBox*>(Sender);
  if (cbox->Checked)
    DropComboTarget1->Formats = DropComboTarget1->Formats << mfText;
  else
    DropComboTarget1->Formats = DropComboTarget1->Formats >> mfText;
}
//---------------------------------------------------------------------------

void __fastcall TFormMain::CheckBoxFilesClick(TObject *Sender)
{
  // Enable or disable format according to users selection
  TCheckBox* cbox = static_cast<TCheckBox*>(Sender);
  if (cbox->Checked)
    DropComboTarget1->Formats = DropComboTarget1->Formats << mfFile;
  else
    DropComboTarget1->Formats = DropComboTarget1->Formats >> mfFile;
}
//---------------------------------------------------------------------------

void __fastcall TFormMain::CheckBoxURLsClick(TObject *Sender)
{
  // Enable or disable format according to users selection
  TCheckBox* cbox = static_cast<TCheckBox*>(Sender);
  if (cbox->Checked)
    DropComboTarget1->Formats = DropComboTarget1->Formats << mfURL;
  else
    DropComboTarget1->Formats = DropComboTarget1->Formats >> mfURL;
}
//---------------------------------------------------------------------------

void __fastcall TFormMain::CheckBoxBitmapsClick(TObject *Sender)
{
  // Enable or disable format according to users selection
  TCheckBox* cbox = static_cast<TCheckBox*>(Sender);
  if (cbox->Checked)
    DropComboTarget1->Formats = DropComboTarget1->Formats << mfBitmap;
  else
    DropComboTarget1->Formats = DropComboTarget1->Formats >> mfBitmap;
}
//---------------------------------------------------------------------------

void __fastcall TFormMain::CheckBoxMetaFilesClick(TObject *Sender)
{
  // Enable or disable format according to users selection
  TCheckBox* cbox = static_cast<TCheckBox*>(Sender);
  if (cbox->Checked)
    DropComboTarget1->Formats = DropComboTarget1->Formats << mfMetaFile;
  else
    DropComboTarget1->Formats = DropComboTarget1->Formats >> mfMetaFile;
}
//---------------------------------------------------------------------------

void __fastcall TFormMain::CheckBoxDataClick(TObject *Sender)
{
  // Enable or disable format according to users selection
  TCheckBox* cbox = static_cast<TCheckBox*>(Sender);
  if (cbox->Checked)
    DropComboTarget1->Formats = DropComboTarget1->Formats << mfData;
  else
    DropComboTarget1->Formats = DropComboTarget1->Formats >> mfData;
}
//---------------------------------------------------------------------------

