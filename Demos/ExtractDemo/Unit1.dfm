object FormMain: TFormMain
  Left = 193
  Top = 169
  AutoScroll = False
  Caption = 'Extract/Download Demo'
  ClientHeight = 261
  ClientWidth = 437
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Arial'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  Scaled = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 15
  object ListView1: TListView
    Left = 0
    Top = 41
    Width = 437
    Height = 201
    Align = alClient
    Columns = <
      item
        Caption = 'A list of  '#39'archived'#39'  files...'
        Width = 400
      end>
    ColumnClick = False
    Items.Data = {
      D40000000500000000000000FFFFFFFFFFFFFFFF00000000000000000D526F6F
      7446696C65312E74787400000000FFFFFFFFFFFFFFFF00000000000000000D52
      6F6F7446696C65322E77726900000000FFFFFFFFFFFFFFFF0000000000000000
      13537562466F6C6465725C46696C65332E70617300000000FFFFFFFFFFFFFFFF
      000000000000000013537562466F6C6465725C46696C65342E64666D00000000
      FFFFFFFFFFFFFFFF000000000000000023537562466F6C6465725C4E65737465
      64537562466F6C6465725C46696C65352E637070}
    MultiSelect = True
    ReadOnly = True
    RowSelect = True
    TabOrder = 0
    ViewStyle = vsReport
    OnMouseDown = ListView1MouseDown
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 437
    Height = 41
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    object Label2: TLabel
      Left = 6
      Top = 6
      Width = 276
      Height = 15
      Caption = 'A demo of how to drag files from a zipped archive...'
    end
    object ButtonClose: TButton
      Left = 369
      Top = 5
      Width = 62
      Height = 25
      Cancel = True
      Caption = 'E&xit'
      TabOrder = 0
      OnClick = ButtonCloseClick
    end
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 242
    Width = 437
    Height = 19
    Panels = <>
    SimplePanel = True
    SimpleText = #39'Extract'#39'  files by dragging them to Explorer ...'
  end
  object DropFileSource1: TDropFileSource
    DragTypes = [dtCopy, dtMove]
    OnDrop = DropFileSource1Drop
    OnAfterDrop = DropFileSource1AfterDrop
    Left = 398
    Top = 92
  end
end
