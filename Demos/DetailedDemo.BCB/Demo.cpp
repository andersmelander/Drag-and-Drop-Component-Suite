//---------------------------------------------------------------------------
// This demo application was contributed by Jonathan Arnold.
//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Demo.h"
#include "DropFile.h"
#include "DropText.h"
#include "DropURL.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFormDemo *FormDemo;
//---------------------------------------------------------------------------
__fastcall TFormDemo::TFormDemo(TComponent* Owner): TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TFormDemo::ButtonTextClick(TObject *Sender)
{
  // Run D&D Text test
  TFormText* ft = new TFormText(this);
  try {
    ft->ShowModal();
  }
  __finally {
    delete ft;
  }
}
//---------------------------------------------------------------------------
void __fastcall TFormDemo::ButtonFileClick(TObject *Sender)
{
  // Run D&D file test
  TFormFile* ft = new TFormFile(this);
  try {
    ft->ShowModal();
  }
  __finally {
    delete ft;
  }
}
//---------------------------------------------------------------------------
void __fastcall TFormDemo::ButtonURLClick(TObject *Sender)
{
  // Run D&D URL test
  TFormURL* ft = new TFormURL(this);
  try {
    ft->ShowModal();
  }
  __finally {
    delete ft;
  }
}
//---------------------------------------------------------------------------
void __fastcall TFormDemo::ButtonExitClick(TObject *Sender)
{
  Close();
}
//---------------------------------------------------------------------------
void __fastcall TFormDemo::Label6Click(TObject *Sender)
{
  // Send email
  Screen->Cursor = crAppStart;
  try {
    Application->ProcessMessages();	// Show cursor change
    ::ShellExecute( 0, NULL, ("mailto:"+((TLabel*)Sender)->Caption).c_str(),
      0, 0, SW_NORMAL );
  }
  __finally {
    Screen->Cursor = crDefault;
  }
}
//---------------------------------------------------------------------------

void __fastcall TFormDemo::Label7Click(TObject *Sender)
{
  // Fire off URL
  Screen->Cursor = crAppStart;
  try {
    Application->ProcessMessages();	// Show cursor change
    ::ShellExecute( 0, NULL, ((TLabel*)Sender)->Caption.c_str(),
      0, 0, SW_NORMAL );
  }
  __finally {
    Screen->Cursor = crDefault;
  }
}
//---------------------------------------------------------------------------

