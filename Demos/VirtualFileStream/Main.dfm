object FormMain: TFormMain
  Left = 269
  Top = 103
  Width = 381
  Height = 257
  BorderIcons = [biSystemMenu, biMinimize]
  Caption = 'Virtual File Stream (File Contents) Demo'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Arial'
  Font.Style = []
  OldCreateOrder = False
  Position = poDefaultPosOnly
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 373
    Height = 171
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 4
    Caption = ' '
    TabOrder = 0
    object ListView1: TListView
      Left = 4
      Top = 4
      Width = 365
      Height = 163
      Align = alClient
      Columns = <
        item
          Caption = 'File name'
          Width = 100
        end
        item
          AutoSize = True
          Caption = 'Contents'
        end>
      Items.Data = {
        A90000000200000000000000FFFFFFFFFFFFFFFF01000000000000000D466972
        737446696C652E74787428546869732069732074686520636F6E74656E74206F
        66207468652066697273742066696C652E2E2E00000000FFFFFFFFFFFFFFFF01
        000000000000000E5365636F6E6446696C652E7478742E2E2E2E616E64207468
        69732069732074686520636F6E74656E74206F6620746865207365636F6E6420
        66696C652EFFFFFFFF}
      MultiSelect = True
      ReadOnly = True
      PopupMenu = PopupMenu1
      TabOrder = 0
      ViewStyle = vsReport
      OnMouseDown = OnMouseDown
      OnMouseMove = ListView1MouseMove
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 171
    Width = 373
    Height = 59
    Align = alBottom
    BevelOuter = bvNone
    BorderWidth = 4
    Caption = ' '
    TabOrder = 1
    object Label1: TLabel
      Left = 4
      Top = 4
      Width = 365
      Height = 51
      Align = alClient
      Caption = 
        'Drag between the above list and any application which accepts th' +
        'e FileContents format (e.g. the VirtualFile demo application or ' +
        'Explorer and Outlook).'
      ShowAccelChar = False
      WordWrap = True
    end
  end
  object DropDummy1: TDropDummy
    DragTypes = []
    Target = Owner
    Left = 96
    Top = 120
  end
  object DropEmptySource1: TDropEmptySource
    DragTypes = [dtCopy, dtMove]
    OnAfterDrop = DropEmptySource1AfterDrop
    Left = 16
    Top = 88
  end
  object DropEmptyTarget1: TDropEmptyTarget
    DragTypes = [dtCopy, dtLink]
    OnEnter = DropEmptyTarget1Enter
    OnDrop = DropEmptyTarget1Drop
    Target = ListView1
    Left = 56
    Top = 88
  end
  object DataFormatAdapterSource: TDataFormatAdapter
    DragDropComponent = DropEmptySource1
    DataFormatName = 'TVirtualFileStreamDataFormat'
    Left = 16
    Top = 120
  end
  object DataFormatAdapterTarget: TDataFormatAdapter
    DragDropComponent = DropEmptyTarget1
    DataFormatName = 'TVirtualFileStreamDataFormat'
    Left = 56
    Top = 120
  end
  object PopupMenu1: TPopupMenu
    OnPopup = PopupMenu1Popup
    Left = 132
    Top = 120
    object MenuCopy: TMenuItem
      Caption = 'Copy'
      OnClick = MenuCopyClick
    end
    object MenuPaste: TMenuItem
      Caption = 'Paste'
      OnClick = MenuPasteClick
    end
  end
end
