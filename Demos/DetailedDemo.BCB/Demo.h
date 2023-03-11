//---------------------------------------------------------------------------

#ifndef DemoH
#define DemoH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Buttons.hpp>
#include <ExtCtrls.hpp>
//---------------------------------------------------------------------------
class TFormDemo : public TForm
{
__published:	// IDE-managed Components
    TPanel *Panel1;
    TBitBtn *ButtonText;
    TBitBtn *ButtonExit;
    TBitBtn *ButtonFile;
    TPanel *Panel2;
    TLabel *Label2;
    TLabel *Label4;
    TLabel *Label5;
    TLabel *Label6;
    TLabel *Label7;
    TPanel *Panel3;
    TBitBtn *ButtonURL;
    void __fastcall ButtonTextClick(TObject *Sender);
    void __fastcall ButtonFileClick(TObject *Sender);
    void __fastcall ButtonURLClick(TObject *Sender);
    void __fastcall ButtonExitClick(TObject *Sender);
    void __fastcall Label6Click(TObject *Sender);
    void __fastcall Label7Click(TObject *Sender);
private:	// User declarations
public:		// User declarations
    __fastcall TFormDemo(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFormDemo *FormDemo;
//---------------------------------------------------------------------------
#endif
