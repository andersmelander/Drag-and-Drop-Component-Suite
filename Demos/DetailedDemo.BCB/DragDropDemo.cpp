//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
USERES("DragDropDemo.res");
USEFORM("Demo.cpp", FormDemo);
USEFORM("DropFile.cpp", FormFile);
USEFORM("DropText.cpp", FormText);
USEFORM("DropURL.cpp", FormURL);
USERES("cursors.res");
USE("..\DetailedDemo\readme.txt", File);
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
  try {
    Application->Initialize();
    Application->CreateForm(__classid(TFormDemo), &FormDemo);
    Application->Run();
  }
  catch (Exception &exception) {
    Application->ShowException(&exception);
  }
  return 0;
}
//---------------------------------------------------------------------------
