//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop
USERES("CustomFormat1BCB4.res");
USEFORM("Target.cpp", FormTarget);
USEFORM("Source.cpp", FormSource);
USEFILE("..\CustomFormat1\readme.txt");
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
	try
	{
		Application->Initialize();
		Application->CreateForm(__classid(TFormSource), &FormSource);
		Application->CreateForm(__classid(TFormTarget), &FormTarget);
		Application->Run();
	}
	catch (Exception &exception)
	{
		Application->ShowException(&exception);
	}
	return 0;
}
//---------------------------------------------------------------------------
