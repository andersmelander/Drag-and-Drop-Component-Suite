unit DragDropDesign;
// TODO : Default event for target components should be OnDrop.
// TODO : Add parent form to Target property editor list.
// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite
// Module:          DragDropDesign
// Description:     Contains design-time support for the drag and drop
//                  components.
// Version:         4.1
// Date:            22-JAN-2002
// Target:          Win32, Delphi 4-6, C++Builder 4-6
// Authors:         Anders Melander, anders@melander.dk, http://www.melander.dk
// Copyright        © 1997-2002 Angus Johnson & Anders Melander
// -----------------------------------------------------------------------------

interface

{$include DragDrop.inc}

procedure Register;

implementation

uses
  DragDrop,
{$ifndef VER14_PLUS}
  DsgnIntf,
{$else}
  DesignIntf,
  DesignEditors,
{$endif}
  Classes;

type
  TDataFormatNameEditor = class(TStringProperty)
  public
    function GetAttributes: TPropertyAttributes; override;
    procedure GetValues(Proc: TGetStrProc); override;
  end;

procedure Register;
begin
  RegisterPropertyEditor(TypeInfo(string), TDataFormatAdapter, 'DataFormatName',
    TDataFormatNameEditor);
end;

{ TDataFormatNameEditor }

function TDataFormatNameEditor.GetAttributes: TPropertyAttributes;
begin
  Result := [paValueList, paSortList, paMultiSelect];
end;

type
  TDataFormatClassesCracker = class(TDataFormatClasses);

procedure TDataFormatNameEditor.GetValues(Proc: TGetStrProc);
var
  i: Integer;
begin
  for i := 0 to TDataFormatClassesCracker.Instance.Count-1 do
    Proc(TDataFormatClassesCracker.Instance[i].ClassName);
end;

end.
