//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop
USERES("AutoScrollBCB4.res");
USEFORM("main.cpp", FormAutoScroll);
USEFILE("..\AutoScroll\readme.txt");
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
	try
	{
		Application->Initialize();
		Application->CreateForm(__classid(TFormAutoScroll), &FormAutoScroll);
		Application->Run();
	}
	catch (Exception &exception)
	{
		Application->ShowException(&exception);
	}
	return 0;
}
//---------------------------------------------------------------------------
