//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
USERES("ExtractDemo.res");
USEFORM("unit1.cpp", FormMain);
USE("..\ExtractDemo\readme.txt", File);
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
