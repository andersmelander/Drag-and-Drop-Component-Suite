//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop
USERES("ExtractDemoBCB4.res");
USEFORM("unit1.cpp", FormMain);
USEFILE("..\ExtractDemo\readme.txt");
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
	try
	{
		Application->Initialize();
		Application->CreateForm(__classid(TFormMain), &FormMain);
		Application->Run();
	}
	catch (Exception &exception)
	{
		Application->ShowException(&exception);
	}
	return 0;
}
//---------------------------------------------------------------------------
