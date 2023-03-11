//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
USERES("CustomFormat1.res");
USEFORM("Source.cpp", FormSource);
USEFORM("Target.cpp", FormTarget);
USE("..\CustomFormat1\readme.txt", File);
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
  try {
    Application->Initialize();
    Application->CreateForm(__classid(TFormSource), &FormSource);
    Application->CreateForm(__classid(TFormTarget), &FormTarget);
    Application->Run();
  }
  catch (Exception &exception) {
    Application->ShowException(&exception);
  }
  return 0;
}
//---------------------------------------------------------------------------
