unit DragDropDesign;
// -----------------------------------------------------------------------------
//
//			NOT FOR RELEASE
//
// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite
// Module:          DragDropDesign
// Description:     Contains design-time support for the drag and drop
//                  components.
// Version:         4.0
// Date:            25-JUN-2000
// Target:          Win32, Delphi 3-6 and C++ Builder 3-5
// Authors:         Angus Johnson, ajohnson@rpi.net.au
//                  Anders Melander, anders@melander.dk, http://www.melander.dk
// Copyright        © 1997-2000 Angus Johnson & Anders Melander
// -----------------------------------------------------------------------------

interface

procedure Register;

implementation

{$include DragDrop.inc}

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
  i		: Integer;
begin
  for i := 0 to TDataFormatClassesCracker.Instance.Count-1 do
    Proc(TDataFormatClassesCracker.Instance[i].ClassName);
end;

end.
