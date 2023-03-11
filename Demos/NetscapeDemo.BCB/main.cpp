//---------------------------------------------------------------------------

#include <vcl.h>
#include <ole2.hpp>
#include <utilcls.h>
#include <registry.hpp>
#pragma hdrstop

#undef GetClipboardFormatName

#include "main.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DragDrop"
#pragma link "DragDropInternet"
#pragma link "DropTarget"
#pragma link "DragDrop"
#pragma resource "*.dfm"
TFormMain *FormMain;
//---------------------------------------------------------------------------
__fastcall TFormMain::TFormMain(TComponent* Owner): TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::DropURLTarget1Drop(TObject *Sender,
  TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  Edit1->Text = DropURLTarget1->URL;

  Variant Netscape = CreateOleObject("Netscape.Network.1");

  try {
    // Avoid leaking a BSTR, which happens if we just pass c_str in
    TOleString url(DropURLTarget1->URL.c_str());

    if (Netscape.OleFunction("open", (BSTR)url, 0, 0, 0, 0) == 0)
       throw Exception("Failed to open URL: "+DropURLTarget1->URL);

    int Size;
    WideString Buffer;
    Buffer.SetLength(255);
    AnsiString s;

    TVariant wb = Buffer.c_bstr();

    // Note: Netscape Messenger must be running in order to read the mail.
    do {
      Size = Netscape.OleFunction("Read", (BSTR*)wb, 255);
      s = s + Buffer;
    } while (Size > 0);


    if ((Size == -1) && (Netscape.OleFunction("GetStatus") != 0))
      throw Exception("Failed to read from URL. Status: "+
        AnsiString(Netscape.OleFunction("GetStatus")));

    ParseRFC822(s);

  }
  __finally {
    Netscape = Unassigned;
  }
}
//---------------------------------------------------------------------------

void __fastcall TFormMain::DropURLTarget1Enter(TObject *Sender,
  TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  // Reject drop unless the "Netscape Message" format is present in the data
  // object. Even though this format doesn't contain any usefull information,
  // its presence does indicate that the drop is a Netscape message and not a
  // regular URL.
  TNetscapeMessageClipboardFormat* NetscapeClip = new TNetscapeMessageClipboardFormat();
  try {
    if (!NetscapeClip->HasValidFormats(DropURLTarget1->DataObject))
      Effect = DROPEFFECT_NONE;
  }
  __finally {
    delete NetscapeClip;
  }
  // Another way to do the same could be to set GetDataOnEnter = True and
  // then parse the URL here. If the URL starts with "mailbox:" the drop
  // contains a mail message.
}
//---------------------------------------------------------------------------

void __fastcall TFormMain::Button1Click(TObject *Sender)
{
  AnsiString s;

  // TODO : Needs conditional for C++Builder 4.
  TRegistry* reg = new TRegistry(KEY_READ);
  try {
    reg->RootKey = HKEY_LOCAL_MACHINE;

    // Note: The following registry keys should work for all versions of
    // Navigator 4->x. Previous versions has not been tested/verified.
    if (reg->OpenKeyReadOnly("\\SOFTWARE\\Netscape\\Netscape Navigator")) {
      s = reg->ReadString("CurrentVersion");
      if (reg->OpenKeyReadOnly("\\SOFTWARE\\Netscape\\Netscape Navigator\\"+s+"\\Main")) {
        s = reg->ReadString("Install Directory");
        if (!s.IsEmpty())
          s = s+"\\Program\\netscape.exe";
      }
    }
    else {
      // Note: The following registry keys are valid for Navigator 6, but the
      // ShellExecute haven't been tested with version 6.
      if (reg->OpenKeyReadOnly("\\SOFTWARE\\Netscape\\Netscape 6")) {
        s = reg->ReadString("CurrentVersion");
        if (reg->OpenKeyReadOnly("\\SOFTWARE\\Netscape\\Netscape 6\\"+s+"\\Main"))
          s = reg->ReadString("PathToExe");
        else
          s = "";
      }
    }
  }
  __finally {
    delete reg;
  }

  if (!s.IsEmpty())
    ShellExecute(Handle, "Open", s.c_str(), Edit1->Text.c_str(), "", SW_SHOWNORMAL);
}
//---------------------------------------------------------------------------

void TFormMain::ParseRFC822(const AnsiString Msg)
{
  /*
  ** Parse a text message in RFC 822 format.
  **
  ** Note: This is a very simple implementation. I have not read the RFC822
  ** specifications and do not know if what I do here is correct. If you need
  ** a RFC822 decoder you should probably look elsewhere for a reference
  ** implementation.
  */
  ListViewHeader->Items->Clear();
  MemoContent->Lines->Clear();
  MemoRaw->Lines->Text = Msg;
  TStringList* Lines = new TStringList();
  try {
    Lines->Text = Msg;
    int i = 0;
    while (i < Lines->Count) {
      AnsiString s = Lines->Strings[i];
      int n  = s.Pos(":");
      if (n == 0) {
        if (s.IsEmpty())
          i++;
        break;
      }

      TListItem* item = ListViewHeader->Items->Add();
      item->Caption = s.SubString(1, n-1);
      AnsiString Value = s.SubString(n+1, s.Length());
      i++;
      while (i < Lines->Count) {
        s = Lines->Strings[i];
        if (!s.IsEmpty() && s.IsDelimiter("\t ", 1)) {
          Value = Value+" "+s.TrimLeft();
          i++;
        }
        else
          break;
      }
      item->SubItems->Add(Value);
    }

    while (i < Lines->Count) {
      AnsiString s = Lines->Strings[i];
      i++;
      while ((s.SubString(s.Length()-2, 3) == "=20") && (i < Lines->Count)) {
        s = s.SubString(1, s.Length()-3)+Lines->Strings[i];
        i++;
      }
      MemoContent->Lines->Add(s);
    }
  }
  __finally {
    delete Lines;
  }
}

