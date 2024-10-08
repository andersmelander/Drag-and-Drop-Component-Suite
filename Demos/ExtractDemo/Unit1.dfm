object FormMain: TFormMain
  Left = 193
  Top = 169
  Caption = 'Extract/Download Demo'
  ClientHeight = 261
  ClientWidth = 437
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  Scaled = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
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
    Items.ItemData = {
      05480100000500000000000000FFFFFFFFFFFFFFFF00000000FFFFFFFF000000
      000D52006F006F007400460069006C00650031002E0074007800740000000000
      FFFFFFFFFFFFFFFF00000000FFFFFFFF000000000D52006F006F007400460069
      006C00650032002E0077007200690000000000FFFFFFFFFFFFFFFF00000000FF
      FFFFFF000000001353007500620046006F006C006400650072005C0046006900
      6C00650033002E0070006100730000000000FFFFFFFFFFFFFFFF00000000FFFF
      FFFF000000001353007500620046006F006C006400650072005C00460069006C
      00650034002E00640066006D0000000000FFFFFFFFFFFFFFFF00000000FFFFFF
      FF000000002353007500620046006F006C006400650072005C004E0065007300
      74006500640053007500620046006F006C006400650072005C00460069006C00
      650035002E00630070007000}
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
      Width = 250
      Height = 13
      Margins.Bottom = 0
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
    ParentFont = True
    SimplePanel = True
    SimpleText = #39'Extract'#39'  files by dragging them to Explorer ...'
    UseSystemFont = False
  end
  object DropFileSource1: TDropFileSource
    DragTypes = [dtCopy, dtMove]
    OnDrop = DropFileSource1Drop
    OnAfterDrop = DropFileSource1AfterDrop
    OnGetData = DropFileSource1GetData
    Left = 398
    Top = 92
  end
end
