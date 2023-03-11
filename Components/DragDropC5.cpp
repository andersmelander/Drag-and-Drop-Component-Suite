//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop
USEUNIT("DragDrop.pas");
USEUNIT("DragDropFile.pas");
USERES("DragDropFile.dcr");
USEUNIT("DragDropFormats.pas");
USEUNIT("DragDropGraphics.pas");
USERES("DragDropGraphics.dcr");
USEUNIT("DragDropPIDL.pas");
USERES("DragDropPIDL.dcr");
USEUNIT("DragDropText.pas");
USERES("DragDropText.dcr");
USEUNIT("DropSource.pas");
USERES("DropSource.dcr");
USEUNIT("DropTarget.pas");
USERES("DropTarget.dcr");
USEUNIT("DragDropInternet.pas");
USERES("DragDropInternet.dcr");
USEPACKAGE("Vcl50.bpi");
USEUNIT("DropComboTarget.pas");
USERES("DropComboTarget.dcr");
USEUNIT("DragDropHandler.pas");
USERES("DragDropHandler.dcr");
//---------------------------------------------------------------------------
#pragma package(smart_init)
//---------------------------------------------------------------------------
//   Package source.
//---------------------------------------------------------------------------
int WINAPI DllEntryPoint(HINSTANCE hinst, unsigned long reason, void*)
{
	return 1;
}
//---------------------------------------------------------------------------
