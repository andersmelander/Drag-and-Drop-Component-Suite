//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop
USERES("AdapterBCB4.res");
USEFORM("main.cpp", FormMain);
USEFILE("..\Adapters\readme.txt");
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
