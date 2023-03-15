//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop
USERES("dragdropC4.res");
USEUNIT("DropSource.pas");
USERES("DropSource.dcr");
USEUNIT("DropURLTarget.pas");
USERES("DropURLTarget.dcr");
USEUNIT("DropBMPTarget.pas");
USERES("DropBMPTarget.dcr");
USEUNIT("DropPIDLSource.pas");
USERES("DropPIDLSource.dcr");
USEUNIT("DropPIDLTarget.pas");
USERES("DropPIDLTarget.dcr");
USEUNIT("DropTarget.pas");
USERES("DropTarget.dcr");
USEUNIT("DropURLSource.pas");
USERES("DropURLSource.dcr");
USEUNIT("DropBMPSource.pas");
USERES("DropBMPSource.dcr");
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
