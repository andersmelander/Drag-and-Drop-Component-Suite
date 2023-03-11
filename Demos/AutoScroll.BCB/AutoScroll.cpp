//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
USERES("AutoScroll.res");
USEFORM("main.cpp", FormAutoScroll);
USE("..\AutoScroll\readme.txt", File);
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
