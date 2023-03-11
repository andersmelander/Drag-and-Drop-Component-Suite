//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop
USERES("DragDropDemoBCB4.res");
USEFORM("Demo.cpp", FormDemo);
USEFORM("DropFile.cpp", FormFile);
USEFORM("DropText.cpp", FormText);
USEFORM("DropURL.cpp", FormURL);
USEFILE("..\DetailedDemo\readme.txt");
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
	try
	{
		Application->Initialize();
		Application->CreateForm(__classid(TFormDemo), &FormDemo);
		Application->Run();
	}
	catch (Exception &exception)
	{
		Application->ShowException(&exception);
	}
	return 0;
}
//---------------------------------------------------------------------------
