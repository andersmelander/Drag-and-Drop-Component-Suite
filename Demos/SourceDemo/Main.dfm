object Form1: TForm1
  Left = 236
  Top = 162
  Width = 395
  Height = 292
  Caption = 'Simple Source Demo'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Shell Dlg 2'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  Scaled = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 387
    Height = 38
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    object Label1: TLabel
      Left = 3
      Top = 13
      Width = 256
      Height = 13
      Caption = 'Drag files from this Listbox onto Windows Explorer ...'
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 217
    Width = 387
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    Caption = ' '
    TabOrder = 1
    object ButtonClose: TButton
      Left = 157
      Top = 8
      Width = 75
      Height = 25
      Anchors = [akTop]
      Cancel = True
      Caption = '&Close'
      TabOrder = 0
      OnClick = ButtonCloseClick
    end
  end
  object ListView1: TListView
    Left = 0
    Top = 38
    Width = 387
    Height = 179
    Align = alClient
    Columns = <
      item
        Caption = 'Filenames'
        Width = 380
      end>
    ColumnClick = False
    MultiSelect = True
    ReadOnly = True
    PopupMenu = PopupMenu1
    TabOrder = 2
    ViewStyle = vsReport
    OnDragOver = ListView1DragOver
    OnMouseDown = ListView1MouseDown
  end
  object DropFileSource1: TDropFileSource
    DragTypes = [dtCopy, dtMove, dtLink]
    OnGetDragImage = DropFileSource1GetDragImage
    ShowImage = True
    Left = 355
    Top = 224
  end
  object DropDummy1: TDropDummy
    DragTypes = []
    Target = Owner
    Left = 276
    Top = 224
  end
  object PopupMenu1: TPopupMenu
    AutoPopup = False
    Left = 192
    Top = 136
    object Just1: TMenuItem
      Caption = 'Just'
    end
    object a1: TMenuItem
      Caption = 'a'
    end
    object test1: TMenuItem
      Caption = 'test'
    end
    object of1: TMenuItem
      Caption = 'of'
    end
    object popup1: TMenuItem
      Caption = 'popup'
    end
    object menu1: TMenuItem
      Caption = 'menu'
    end
  end
end
