//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop
USERES("dragdropC4.res");
USEUNIT("DropSource.pas");
USERES("DropSource.dcr");
USEUNIT("DropTarget.pas");
USERES("DropTarget.dcr");
USEUNIT("DragDropDesign.pas");
USEUNIT("DragDropFormats.pas");
USEUNIT("DragDropGraphics.pas");
USERES("DragDropGraphics.dcr");
USEUNIT("DragDropInternet.pas");
USERES("DragDropInternet.dcr");
USEUNIT("DragDropText.pas");
USERES("DragDropText.dcr");
USEUNIT("DropComboTarget.pas");
USERES("DropComboTarget.dcr");
USEUNIT("DragDrop.pas");
USERES("DragDrop.dcr");
USEUNIT("DragDropPIDL.pas");
USERES("DragDropPIDL.dcr");
USEPACKAGE("vcl40.bpi");
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
